// Best-effort auto-fill for the daily "スカウト返信数" fields: Gmail's search narrows
// candidates by date + a broad "スカウト" subject keyword (cheap, one API call), then each
// candidate's exact Subject header is fetched and matched client-side against the known
// notification-email templates per media. This two-step approach avoids depending on Gmail's
// imperfect Japanese tokenization for the actual classification.

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

async function fetchAndClassifyMessages(
  accessToken: string,
  q: string,
  onProgress?: GmailScanProgress
): Promise<{ matches: { dateISO: string; mediaId: string }[]; totalScanned: number }> {
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

  const matches: { dateISO: string; mediaId: string }[] = [];
  details.forEach((detail) => {
    const subjectHeader = (detail.payload?.headers || []).find((h: any) => h.name === 'Subject');
    const subject: string = subjectHeader?.value || '';
    const matched = MEDIA_SUBJECT_MATCHERS.find((m) => m.test(subject));
    if (matched) {
      matches.push({ dateISO: toDateISO(Number(detail.internalDate)), mediaId: matched.mediaId });
    }
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

  const { matches, totalScanned } = await fetchAndClassifyMessages(accessToken, q, onProgress);
  const counts: Record<string, number> = {};
  matches.forEach((m) => { counts[m.mediaId] = (counts[m.mediaId] || 0) + 1; });
  return { counts, totalMatched: matches.length, totalScanned };
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

  const { matches, totalScanned } = await fetchAndClassifyMessages(accessToken, q, onProgress);
  const countsByDate: Record<string, Record<string, number>> = {};
  matches.forEach((m) => {
    countsByDate[m.dateISO] = countsByDate[m.dateISO] || {};
    countsByDate[m.dateISO][m.mediaId] = (countsByDate[m.dateISO][m.mediaId] || 0) + 1;
  });
  return { countsByDate, totalMatched: matches.length, totalScanned };
}
