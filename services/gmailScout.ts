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

const MEDIA_SUBJECT_MATCHERS: { mediaId: string; test: (subject: string) => boolean }[] = [
  { mediaId: 'biz', test: (s) => s.includes('ビズリーチ') && s.includes('スカウト返信') },
  { mediaId: 'rds', test: (s) => s.includes('リクルートダイレクトスカウト') && s.includes('スカウト送付') },
  { mediaId: 'doda', test: (s) => s.includes('返信・質問') && s.includes('スカウト') && s.toLowerCase().includes('doda') },
];

function toGmailDate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}/${m}/${day}`;
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

/** Fetches per-media scout-reply-notification-email counts for a single calendar date (local time). */
export async function fetchScoutReplyCounts(accessToken: string, dateISO: string): Promise<ScoutReplyFetchResult> {
  const dayStart = new Date(dateISO + 'T00:00:00');
  const dayEnd = new Date(dayStart);
  dayEnd.setDate(dayEnd.getDate() + 1);
  const q = `after:${toGmailDate(dayStart)} before:${toGmailDate(dayEnd)} subject:スカウト`;

  let messages: { id: string }[] = [];
  let pageToken: string | undefined;
  do {
    const params = new URLSearchParams({ q, maxResults: '100' });
    if (pageToken) params.set('pageToken', pageToken);
    const listRes = await gmailFetch(accessToken, `messages?${params.toString()}`);
    messages = messages.concat(listRes.messages || []);
    pageToken = listRes.nextPageToken;
  } while (pageToken);

  const counts: Record<string, number> = {};
  let totalMatched = 0;

  for (const msg of messages) {
    const detail = await gmailFetch(accessToken, `messages/${msg.id}?format=metadata&metadataHeaders=Subject`);
    const subjectHeader = (detail.payload?.headers || []).find((h: any) => h.name === 'Subject');
    const subject: string = subjectHeader?.value || '';
    const matched = MEDIA_SUBJECT_MATCHERS.find((m) => m.test(subject));
    if (matched) {
      counts[matched.mediaId] = (counts[matched.mediaId] || 0) + 1;
      totalMatched++;
    }
  }

  return { counts, totalMatched, totalScanned: messages.length };
}
