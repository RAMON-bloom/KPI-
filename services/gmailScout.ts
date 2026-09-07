// Best-effort auto-fill for the daily "スカウト返信数" fields: Gmail's search narrows
// candidates by date + a broad "スカウト" subject keyword (cheap, one API call), then each
// candidate's exact Subject header is fetched and matched client-side against the known
// notification-email templates per media. This two-step approach avoids depending on Gmail's
// imperfect Japanese tokenization for the actual classification. RDS is handled separately by
// fetchRdsReplyMatches below, since its notification emails need cross-day dedup by candidate
// name rather than the plain per-message classification used for the other media.

export class GmailPermissionError extends Error {}

export interface ScoutReplyFetchResult {
  counts: Record<string, number>;
  totalMatched: number;
  totalScanned: number;
}

export interface ScoutReplyRangeResult {
  countsByDate: Record<string, Record<string, number>>;
  totalMatched: number;
  totalScanned: number;
}

export type GmailScanProgress = (done: number, total: number) => void;

const MEDIA_SUBJECT_MATCHERS: { mediaId: string; test: (subject: string) => boolean }[] = [
  { mediaId: 'biz', test: (s) => s.includes('ビズリーチ') && s.includes('スカウト返信') },
  { mediaId: 'rds', test: (s) => s.includes('リクルートダイレクトスカウト') && s.includes('スカウト送付') },
  { mediaId: 'doda', test: (s) => s.includes('返信・質問') && s.includes('スカウト') && s.toLowerCase().includes('doda') },
  { mediaId: 'liiga', test: (s) => s.includes('詳しく聞きたい') && s.includes('メッセージをお送りください') },
];

function toGmailDate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}/${m}/${day}`;
}

function toDateISO(epochMs: number): string {
  const d = new Date(epochMs);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

async function gmailFetch(accessToken: string, path: string): Promise<any> {
  const res = await fetch(`https://gmail.googleapis.com/gmail/v1/users/me/${path}`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (res.status === 401 || res.status === 403) {
    throw new GmailPermissionError('Gmailの読み取り権限が許可されていません。');
  }
  if (!res.ok) {
    throw new Error(`Gmail APIエラー（${res.status}）が発生しました。`);
  }
  return res.json();
}

// Fetches message details with a small concurrency cap instead of one-at-a-time, so a
// multi-month backfill (potentially hundreds of messages) doesn't take minutes to classify.
async function mapWithConcurrency<T, R>(
  items: T[],
  limit: number,
  fn: (item: T) => Promise<R>,
  onProgress?: GmailScanProgress
): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let nextIndex = 0;
  let done = 0;
  async function worker() {
    while (true) {
      const i = nextIndex++;
      if (i >= items.length) return;
      results[i] = await fn(items[i]);
      done++;
      onProgress?.(done, items.length);
    }
  }
  const workers = Array.from({ length: Math.max(1, Math.min(limit, items.length)) }, () => worker());
  await Promise.all(workers);
  return results;
}

function decodeGmailBase64Url(data: string): string {
  try {
    const base64 = data.replace(/-/g, '+').replace(/_/g, '/');
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return new TextDecoder('utf-8').decode(bytes);
  } catch {
    return '';
  }
}

function stripHtmlTags(html: string): string {
  return html
    .replace(/<[^>]*>/g, '\n')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
}

// Gmail's payload is a MIME tree (multipart/mixed, multipart/alternative, ...) — walks it
// collecting every text/plain part; only falls back to text/html (tags stripped) if no plain
// part exists anywhere, since HTML markup would otherwise pollute the "候補者氏名：..." line
// extractRdsCandidateName looks for.
function extractMessageBody(payload: any): string {
  if (!payload) return '';
  const plainParts: string[] = [];
  const htmlParts: string[] = [];
  const walk = (part: any) => {
    if (!part) return;
    const mimeType: string = part.mimeType || '';
    if (part.body?.data && mimeType === 'text/plain') {
      plainParts.push(decodeGmailBase64Url(part.body.data));
    } else if (part.body?.data && mimeType === 'text/html') {
      htmlParts.push(stripHtmlTags(decodeGmailBase64Url(part.body.data)));
    }
    (part.parts || []).forEach(walk);
  };
  walk(payload);
  return (plainParts.length > 0 ? plainParts : htmlParts).join('\n');
}

// RDS's reply-notification body has a "候補者氏名：山田 太郎" line (see kpi-mgr's gmailScout
// history for the confirmed sample) — the subject only carries the recruiter's own name, not
// the candidate's, so this is the only place a candidate identity is available at all.
function extractRdsCandidateName(bodyText: string): string | null {
  const match = bodyText.match(/候補者氏名[:：]\s*([^\r\n]+)/);
  if (!match) return null;
  const name = match[1].trim();
  return name.length > 0 ? name : null;
}

async function fetchAndClassifyMessages(
  accessToken: string,
  q: string,
  onProgress?: GmailScanProgress
): Promise<{ matches: { dateISO: string; mediaId: string; messageId: string }[]; totalScanned: number }> {
  let messages: { id: string }[] = [];
  let pageToken: string | undefined;
  do {
    const params = new URLSearchParams({ q, maxResults: '100' });
    if (pageToken) params.set('pageToken', pageToken);
    const listRes = await gmailFetch(accessToken, `messages?${params.toString()}`);
    messages = messages.concat(listRes.messages || []);
    pageToken = listRes.nextPageToken;
  } while (pageToken);

  const details = await mapWithConcurrency(
    messages,
    8,
    (msg) => gmailFetch(accessToken, `messages/${msg.id}?format=metadata&metadataHeaders=Subject`),
    onProgress
  );

  // rds is deliberately excluded here — fetchRdsReplyMatches below handles it separately with
  // its own cross-day dedup logic instead of the plain per-message classification used here.
  const matches: { dateISO: string; mediaId: string; messageId: string }[] = [];
  details.forEach((detail, i) => {
    const subjectHeader = (detail.payload?.headers || []).find((h: any) => h.name === 'Subject');
    const subject: string = subjectHeader?.value || '';
    const matched = MEDIA_SUBJECT_MATCHERS.find((m) => m.mediaId !== 'rds' && m.test(subject));
    if (matched) {
      matches.push({ dateISO: toDateISO(Number(detail.internalDate)), mediaId: matched.mediaId, messageId: messages[i].id });
    }
  });

  return { matches, totalScanned: messages.length };
}

const RDS_SUBJECT_QUERY = 'subject:リクルートダイレクトスカウト subject:スカウト送付';
// RDS sends a fresh notification email for every message a candidate sends — not just their
// first-ever reply — so an ongoing exchange (the notification itself asks recruiters to respond
// within 5 business days, implying exchanges can span that long) produces one "返信が来ました"
// email per message, on whatever day each message happens to land. Counting every such email as
// a new reply overcounts vs. the actual number of candidates who replied. This scans a window
// widened RDS_HISTORY_LOOKBACK_DAYS before the requested range, extracts each message's
// candidate name from the body, and keeps only each candidate's chronologically first message —
// dropping every later message from the same candidate (same day or weeks later) and anything
// whose first-ever appearance predates the requested range. A body fetch failure or unrecognized
// format leaves that message ungrouped (counted on its own) rather than risking merging two
// different candidates together.
const RDS_HISTORY_LOOKBACK_DAYS = 30;

async function fetchRdsReplyMatches(
  accessToken: string,
  startDateISO: string,
  endDateISOInclusive: string,
  onProgress?: GmailScanProgress
): Promise<{ matches: { dateISO: string; mediaId: string; messageId: string }[]; totalScanned: number }> {
  const rangeStart = new Date(startDateISO + 'T00:00:00');
  const rangeEndExclusive = new Date(endDateISOInclusive + 'T00:00:00');
  rangeEndExclusive.setDate(rangeEndExclusive.getDate() + 1);
  const scanStart = new Date(rangeStart);
  scanStart.setDate(scanStart.getDate() - RDS_HISTORY_LOOKBACK_DAYS);

  const q = `after:${toGmailDate(scanStart)} before:${toGmailDate(rangeEndExclusive)} ${RDS_SUBJECT_QUERY}`;
  let messages: { id: string }[] = [];
  let pageToken: string | undefined;
  do {
    const params = new URLSearchParams({ q, maxResults: '100' });
    if (pageToken) params.set('pageToken', pageToken);
    const listRes = await gmailFetch(accessToken, `messages?${params.toString()}`);
    messages = messages.concat(listRes.messages || []);
    pageToken = listRes.nextPageToken;
  } while (pageToken);

  const details = await mapWithConcurrency(
    messages,
    8,
    (msg) => gmailFetch(accessToken, `messages/${msg.id}?format=full`),
    onProgress
  );

  const entries = details
    .map((detail, i) => ({
      messageId: messages[i].id,
      internalDate: Number(detail.internalDate),
      dateISO: toDateISO(Number(detail.internalDate)),
      candidateName: extractRdsCandidateName(extractMessageBody(detail.payload)),
    }))
    .sort((a, b) => a.internalDate - b.internalDate);

  const rangeStartMs = rangeStart.getTime();
  const seenCandidates = new Set<string>();
  const matches: { dateISO: string; mediaId: string; messageId: string }[] = [];
  entries.forEach((entry) => {
    if (entry.candidateName) {
      if (seenCandidates.has(entry.candidateName)) return;
      seenCandidates.add(entry.candidateName);
    }
    if (entry.internalDate < rangeStartMs) return; // first appearance predates the requested range
    matches.push({ dateISO: entry.dateISO, mediaId: 'rds', messageId: entry.messageId });
  });

  return { matches, totalScanned: messages.length };
}

/** Fetches per-media scout-reply-notification-email counts for a single calendar date (local time). */
export async function fetchScoutReplyCounts(
  accessToken: string,
  dateISO: string,
  onProgress?: GmailScanProgress
): Promise<ScoutReplyFetchResult> {
  const dayStart = new Date(dateISO + 'T00:00:00');
  const dayEnd = new Date(dayStart);
  dayEnd.setDate(dayEnd.getDate() + 1);
  const q = `after:${toGmailDate(dayStart)} before:${toGmailDate(dayEnd)} subject:スカウト`;

  const [general, rds] = await Promise.all([
    fetchAndClassifyMessages(accessToken, q, onProgress),
    fetchRdsReplyMatches(accessToken, dateISO, dateISO),
  ]);
  const matches = [...general.matches, ...rds.matches];
  const counts: Record<string, number> = {};
  matches.forEach((m) => { counts[m.mediaId] = (counts[m.mediaId] || 0) + 1; });
  return { counts, totalMatched: matches.length, totalScanned: general.totalScanned + rds.totalScanned };
}

export interface FullMessage {
  id: string;
  dateISO: string;
  subject: string;
  body: string;
}

/**
 * Fetches every message in [startDateISO, endDateISOInclusive] with full body content — no
 * subject-keyword narrowing, unlike fetchScoutReplyCounts above. Used by
 * pipelineAchievementMatch.ts, which can't rely on a fixed notification template the way
 * scout-reply detection does (these emails come from varied ATS systems or individual
 * client-company staff, with no consistent subject/sender pattern to filter on up front).
 */
export async function fetchFullMessagesInRange(
  accessToken: string,
  startDateISO: string,
  endDateISOInclusive: string,
  onProgress?: GmailScanProgress
): Promise<FullMessage[]> {
  const start = new Date(startDateISO + 'T00:00:00');
  const endExclusive = new Date(endDateISOInclusive + 'T00:00:00');
  endExclusive.setDate(endExclusive.getDate() + 1);
  const q = `after:${toGmailDate(start)} before:${toGmailDate(endExclusive)}`;

  let messages: { id: string }[] = [];
  let pageToken: string | undefined;
  do {
    const params = new URLSearchParams({ q, maxResults: '100' });
    if (pageToken) params.set('pageToken', pageToken);
    const listRes = await gmailFetch(accessToken, `messages?${params.toString()}`);
    messages = messages.concat(listRes.messages || []);
    pageToken = listRes.nextPageToken;
  } while (pageToken);

  const details = await mapWithConcurrency(
    messages,
    8,
    (msg) => gmailFetch(accessToken, `messages/${msg.id}?format=full`),
    onProgress
  );

  return details.map((detail, i) => {
    const subjectHeader = (detail.payload?.headers || []).find((h: any) => h.name === 'Subject');
    return {
      id: messages[i].id,
      dateISO: toDateISO(Number(detail.internalDate)),
      subject: subjectHeader?.value || '',
      body: extractMessageBody(detail.payload),
    };
  });
}

/** Fetches per-media scout-reply-notification-email counts for every date in [startDateISO, endDateISOInclusive]. */
export async function fetchScoutReplyCountsForRange(
  accessToken: string,
  startDateISO: string,
  endDateISOInclusive: string,
  onProgress?: GmailScanProgress
): Promise<ScoutReplyRangeResult> {
  const start = new Date(startDateISO + 'T00:00:00');
  const endExclusive = new Date(endDateISOInclusive + 'T00:00:00');
  endExclusive.setDate(endExclusive.getDate() + 1);
  const q = `after:${toGmailDate(start)} before:${toGmailDate(endExclusive)} subject:スカウト`;

  // Run sequentially (not in parallel) so onProgress reports one coherent 0→total sweep at a
  // time instead of two competing counters interleaving on the same callback.
  const general = await fetchAndClassifyMessages(accessToken, q, onProgress);
  const rds = await fetchRdsReplyMatches(accessToken, startDateISO, endDateISOInclusive, onProgress);
  const matches = [...general.matches, ...rds.matches];
  const countsByDate: Record<string, Record<string, number>> = {};
  matches.forEach((m) => {
    countsByDate[m.dateISO] = countsByDate[m.dateISO] || {};
    countsByDate[m.dateISO][m.mediaId] = (countsByDate[m.dateISO][m.mediaId] || 0) + 1;
  });
  return { countsByDate, totalMatched: matches.length, totalScanned: general.totalScanned + rds.totalScanned };
}
