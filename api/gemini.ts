// Gemini（generateContent）の中継。APIキーはこのサーバー関数だけが持ち、ブラウザには渡さない
// （以前は vite.config.ts の define でブラウザ用プログラムに埋め込まれていて、誰でも取り出せた）。
//
// 呼べるのは bloom-firm.com のGoogleアカウントでログインしている人だけ。ブラウザが持っている
// Googleのアクセストークンを Authorization ヘッダーで受け取り、tokeninfo で確かめる。
//
// Vercelのサーバー関数は受け取れるデータが約4.5MBまでなので、大きいファイル（音声・PDFなど）は
// ブラウザ側（services/gemini.ts）が一時ファイルとして本人のドライブに置き、{ driveFile: { id, mimeType } }
// の形で送ってくる。ここで本人のトークンを使ってドライブから読み、Gemini の Files API に上げてから
// 渡す。Gemini 側のファイルは使い終わったら消す（ドライブの一時ファイルはブラウザ側が消す）。

const GEMINI_BASE = 'https://generativelanguage.googleapis.com';
const ALLOWED_MODELS = new Set(['gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-2.5-flash-lite']);
const ALLOWED_DOMAIN = 'bloom-firm.com';

// 同じトークンで続けて呼ばれることが多いので、確認結果を少しの間だけ覚えておく。
const verifiedTokens = new Map<string, { email: string; until: number }>();

async function verifyUser(authHeader: unknown): Promise<string | null> {
  const token = typeof authHeader === 'string' ? authHeader.replace(/^Bearer\s+/i, '').trim() : '';
  if (!token) return null;
  const cached = verifiedTokens.get(token);
  if (cached && cached.until > Date.now()) return cached.email;
  const res = await fetch(`https://oauth2.googleapis.com/tokeninfo?access_token=${encodeURIComponent(token)}`);
  if (!res.ok) return null;
  const info: any = await res.json().catch(() => null);
  const email = String(info?.email || '').toLowerCase();
  const verified = info?.email_verified === true || info?.email_verified === 'true';
  if (!verified || !email.endsWith(`@${ALLOWED_DOMAIN}`)) return null;
  const clientId = process.env.GOOGLE_CLIENT_ID;
  if (clientId && info?.aud !== clientId && info?.azp !== clientId) return null;
  const expiresInMs = Math.max(0, Number(info?.expires_in || 0) * 1000);
  verifiedTokens.set(token, { email, until: Date.now() + Math.min(5 * 60 * 1000, expiresInMs) });
  if (verifiedTokens.size > 500) verifiedTokens.clear();
  return email;
}

/** ドライブの一時ファイルを読み、Gemini の Files API に上げて、処理が終わるまで待つ。 */
async function uploadDriveFileToGemini(fileId: string, mimeType: string, accessToken: string, apiKey: string): Promise<{ name: string; uri: string }> {
  const driveRes = await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?alt=media&supportsAllDrives=true`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!driveRes.ok) throw Object.assign(new Error(`ドライブの一時ファイルを読めませんでした（${driveRes.status}）`), { status: 502 });
  const bytes = Buffer.from(await driveRes.arrayBuffer());

  const start = await fetch(`${GEMINI_BASE}/upload/v1beta/files`, {
    method: 'POST',
    headers: {
      'x-goog-api-key': apiKey,
      'X-Goog-Upload-Protocol': 'resumable',
      'X-Goog-Upload-Command': 'start',
      'X-Goog-Upload-Header-Content-Length': String(bytes.length),
      'X-Goog-Upload-Header-Content-Type': mimeType,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ file: { display_name: `kpi-mgr-${Date.now()}` } }),
  });
  const uploadUrl = start.headers.get('x-goog-upload-url');
  if (!start.ok || !uploadUrl) throw Object.assign(new Error(`AIへのファイル送信を始められませんでした（${start.status}）`), { status: 502 });
  const done = await fetch(uploadUrl, {
    method: 'POST',
    headers: { 'X-Goog-Upload-Command': 'upload, finalize', 'X-Goog-Upload-Offset': '0' },
    body: bytes,
  });
  const uploaded: any = await done.json().catch(() => null);
  let file = uploaded?.file;
  if (!done.ok || !file?.name) throw Object.assign(new Error(`AIへのファイル送信に失敗しました（${done.status}）`), { status: 502 });

  // 音声・動画は処理が終わる（ACTIVE になる）まで使えない。
  const deadline = Date.now() + 120 * 1000;
  while (file.state === 'PROCESSING' && Date.now() < deadline) {
    await new Promise(r => setTimeout(r, 2000));
    const r = await fetch(`${GEMINI_BASE}/v1beta/${file.name}`, { headers: { 'x-goog-api-key': apiKey } });
    if (r.ok) file = await r.json();
  }
  if (file.state && file.state !== 'ACTIVE') throw Object.assign(new Error('AIがファイルを読み込めませんでした。'), { status: 502 });
  return { name: file.name, uri: file.uri };
}

export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: { message: 'Method not allowed' } });
    return;
  }
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: { message: 'サーバーにAIのAPIキーが設定されていません。' } });
    return;
  }
  const authHeader = req.headers?.authorization;
  const email = await verifyUser(authHeader).catch(() => null);
  if (!email) {
    res.status(401).json({ error: { message: 'ログインの確認ができませんでした。ログインし直してください。' } });
    return;
  }

  const { model, contents, generationConfig, tools, systemInstruction, safetySettings } = req.body || {};
  if (typeof model !== 'string' || !ALLOWED_MODELS.has(model) || !Array.isArray(contents)) {
    res.status(400).json({ error: { message: 'リクエストの形式が正しくありません。' } });
    return;
  }

  const accessToken = String(authHeader).replace(/^Bearer\s+/i, '').trim();
  const geminiFiles: string[] = [];
  try {
    // { driveFile } の部品を Gemini のファイル参照に置き換える
    for (const content of contents) {
      if (!Array.isArray(content?.parts)) continue;
      for (let i = 0; i < content.parts.length; i++) {
        const df = content.parts[i]?.driveFile;
        if (!df) continue;
        if (typeof df.id !== 'string' || typeof df.mimeType !== 'string') throw Object.assign(new Error('ファイルの指定が正しくありません。'), { status: 400 });
        const uploaded = await uploadDriveFileToGemini(df.id, df.mimeType, accessToken, apiKey);
        geminiFiles.push(uploaded.name);
        content.parts[i] = { fileData: { mimeType: df.mimeType, fileUri: uploaded.uri } };
      }
    }

    const body: any = { contents };
    if (generationConfig) body.generationConfig = generationConfig;
    if (tools) body.tools = tools;
    if (systemInstruction) body.systemInstruction = systemInstruction;
    if (safetySettings) body.safetySettings = safetySettings;
    const r = await fetch(`${GEMINI_BASE}/v1beta/models/${model}:generateContent`, {
      method: 'POST',
      headers: { 'x-goog-api-key': apiKey, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const data: any = await r.json().catch(() => ({}));
    if (!r.ok) {
      res.status(r.status).json({ error: { code: r.status, status: data?.error?.status, message: data?.error?.message || `Gemini error ${r.status}` } });
      return;
    }
    res.status(200).json(data);
  } catch (e: any) {
    res.status(e?.status || 500).json({ error: { message: e?.message || 'AIの呼び出しに失敗しました。' } });
  } finally {
    await Promise.all(geminiFiles.map(name =>
      fetch(`${GEMINI_BASE}/v1beta/${name}`, { method: 'DELETE', headers: { 'x-goog-api-key': apiKey } }).catch(() => undefined),
    ));
  }
}
