// Gemini呼び出しの窓口。APIキーはブラウザに持たせず、サーバー関数 /api/gemini が付けて送る。
// 呼び出し側は @google/genai と同じ形（gemini.models.generateContent({ model, contents, config })）で
// 使え、戻り値も response.text / response.candidates を持つ。
//
// サーバー関数が受け取れるのは約4.5MBまでなので、それを超える大きさのファイル（音声・PDFなど）は
// 一時ファイルとして本人のドライブに置き、そのIDだけを送る（api/gemini.ts がドライブから読む）。
// 一時ファイルは呼び出しが終わったら消す。

import { getCurrentSession, refreshTokenSilently } from './googleAuth';

/** responseSchema の型名（@google/genai の Type と同じ値）。 */
export const Type = {
  OBJECT: 'OBJECT',
  STRING: 'STRING',
  ARRAY: 'ARRAY',
  NUMBER: 'NUMBER',
  INTEGER: 'INTEGER',
  BOOLEAN: 'BOOLEAN',
} as const;

export interface GeminiResponse {
  text: string;
  candidates?: any[];
}

// これを超えるリクエストは、大きいファイルをドライブ経由にする（上限4.5MBに余裕を持たせた値）。
const MAX_INLINE_REQUEST_BYTES = 3_500_000;
// ドライブ経由にするファイルの最小サイズ（base64の文字数）。小さい画像などはそのまま送る。
const MIN_OFFLOAD_PART_CHARS = 200_000;

async function getAccessToken(): Promise<string> {
  let session = getCurrentSession();
  if (!session) {
    try { await refreshTokenSilently(); } catch { /* 下で判定 */ }
    session = getCurrentSession();
  }
  if (!session) throw new Error('AIを使うにはGoogleでログインしてください。');
  return session.accessToken;
}

const normalizePart = (p: any) => (typeof p === 'string' ? { text: p } : p);

/** @google/genai が受け付ける contents の書き方を、REST API の Content[] にそろえる。 */
function toContents(contents: any): any[] {
  if (typeof contents === 'string') return [{ role: 'user', parts: [{ text: contents }] }];
  if (Array.isArray(contents)) {
    if (contents.length > 0 && contents.every(c => c && typeof c === 'object' && Array.isArray(c.parts))) {
      return contents.map(c => ({ role: c.role || 'user', parts: c.parts.map(normalizePart) }));
    }
    return [{ role: 'user', parts: contents.map(normalizePart) }];
  }
  if (contents && Array.isArray(contents.parts)) return [{ role: contents.role || 'user', parts: contents.parts.map(normalizePart) }];
  return [{ role: 'user', parts: [{ text: String(contents) }] }];
}

function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

async function uploadTempDriveFile(token: string, data: string, mimeType: string): Promise<string> {
  const start = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id', {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json; charset=UTF-8', 'X-Upload-Content-Type': mimeType },
    body: JSON.stringify({ name: `kpi-mgr-ai-temp-${Date.now()}`, mimeType, description: 'KPI管理くんがAIに渡すための一時ファイルです（自動で削除されます）' }),
  });
  const location = start.headers.get('Location');
  if (!start.ok || !location) throw new Error(`AIに渡すファイルの準備に失敗しました（ドライブ ${start.status}）。`);
  const put = await fetch(location, { method: 'PUT', headers: { 'Content-Type': mimeType }, body: base64ToBytes(data) });
  const file = await put.json().catch(() => null);
  if (!put.ok || !file?.id) throw new Error(`AIに渡すファイルの準備に失敗しました（ドライブ ${put.status}）。`);
  return file.id;
}

async function deleteTempDriveFile(token: string, id: string) {
  await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}`, {
    method: 'DELETE',
    headers: { Authorization: `Bearer ${token}` },
  }).catch(() => undefined);
}

/** サーバーやGeminiのエラーを、画面に出せる日本語にする。 */
function toJapaneseError(status: number, payload: any): Error {
  const message: string = payload?.error?.message || '';
  const geminiStatus: string = payload?.error?.status || '';
  let text: string;
  if (status === 429 || geminiStatus === 'RESOURCE_EXHAUSTED') {
    text = /credit|billing|prepay|balance/i.test(message)
      ? 'AIの利用クレジットが不足しています。管理者にクレジットの追加を依頼してください。'
      : 'AIの利用が一時的に上限に達しました。1分ほど待ってからもう一度お試しください。';
  } else if (status === 401) {
    text = message || 'ログインの確認ができませんでした。ログインし直してください。';
  } else if (status === 413) {
    text = 'AIに送るデータが大きすぎます。ファイルを小さくしてからお試しください。';
  } else if (status === 503 || geminiStatus === 'UNAVAILABLE') {
    text = 'AIが混み合っています。少し待ってからもう一度お試しください。';
  } else if (/[ぁ-んァ-ン一-龥]/.test(message)) {
    text = message;
  } else {
    text = `AIの呼び出しに失敗しました（${status}${geminiStatus ? ` ${geminiStatus}` : ''}）。`;
  }
  const error = new Error(text) as Error & { status?: number; detail?: string };
  error.status = status;
  error.detail = message;
  return error;
}

async function generateContent(params: { model: string; contents: any; config?: Record<string, any> }): Promise<GeminiResponse> {
  const token = await getAccessToken();
  const { tools, systemInstruction, safetySettings, abortSignal, httpOptions, ...generationConfig } = params.config || {};
  const contents = toContents(params.contents);
  const body: any = { model: params.model, contents };
  if (Object.keys(generationConfig).length > 0) body.generationConfig = generationConfig;
  if (tools) body.tools = tools;
  if (systemInstruction) body.systemInstruction = typeof systemInstruction === 'string' ? { parts: [{ text: systemInstruction }] } : systemInstruction;
  if (safetySettings) body.safetySettings = safetySettings;

  // 大きすぎるときは、大きいファイルから順にドライブ経由に切り替える。
  const tempFileIds: string[] = [];
  try {
    let size = JSON.stringify(body).length;
    if (size > MAX_INLINE_REQUEST_BYTES) {
      const inlineParts = contents
        .flatMap((c: any) => c.parts.map((p: any, i: number) => ({ c, i, p })))
        .filter(({ p }) => typeof p?.inlineData?.data === 'string' && p.inlineData.data.length >= MIN_OFFLOAD_PART_CHARS)
        .sort((a: any, b: any) => b.p.inlineData.data.length - a.p.inlineData.data.length);
      for (const { c, i, p } of inlineParts) {
        if (size <= MAX_INLINE_REQUEST_BYTES) break;
        const id = await uploadTempDriveFile(token, p.inlineData.data, p.inlineData.mimeType);
        tempFileIds.push(id);
        c.parts[i] = { driveFile: { id, mimeType: p.inlineData.mimeType } };
        size -= p.inlineData.data.length;
      }
    }

    const res = await fetch('/api/gemini', {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
      signal: abortSignal,
    });
    const payload = await res.json().catch(() => null);
    if (!res.ok) throw toJapaneseError(res.status, payload);
    const candidates: any[] = payload?.candidates || [];
    const text = (candidates[0]?.content?.parts || [])
      .filter((p: any) => typeof p?.text === 'string' && !p.thought)
      .map((p: any) => p.text)
      .join('');
    return { text, candidates };
  } finally {
    await Promise.all(tempFileIds.map(id => deleteTempDriveFile(token, id)));
  }
}

/** @google/genai の `new GoogleGenAI(...)` の代わりに使う。 */
export const gemini = { models: { generateContent } };
