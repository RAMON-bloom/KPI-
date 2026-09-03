// Google Chatの受信Webhookへメッセージを転送するだけのプロキシ。
//
// ブラウザからchat.googleapis.comへ直接fetchするとCORSでブロックされるため、このサーバー
// レス関数を経由してPOSTする（フロントエンドの sendChatWebhookMessage / TeamChatReport
// Panel から呼ばれる — index.tsx参照）。転送先はchat.googleapis.comのURLのみに制限し、
// 任意の内部/外部ホストへ中継できてしまうのを防ぐ。
export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  const { webhookUrl, text } = req.body || {};
  if (typeof webhookUrl !== 'string' || typeof text !== 'string' || !text.trim()) {
    res.status(400).json({ error: 'webhookUrl and text are required' });
    return;
  }

  let parsed: URL;
  try {
    parsed = new URL(webhookUrl);
  } catch {
    res.status(400).json({ error: 'invalid webhookUrl' });
    return;
  }
  if (parsed.protocol !== 'https:' || parsed.hostname !== 'chat.googleapis.com') {
    res.status(400).json({ error: 'webhookUrl must be an https://chat.googleapis.com URL' });
    return;
  }

  try {
    const chatResponse = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=UTF-8' },
      body: JSON.stringify({ text }),
    });
    const bodyText = await chatResponse.text();
    if (!chatResponse.ok) {
      res.status(502).json({ error: `Google Chatへの送信に失敗しました（${chatResponse.status}）`, detail: bodyText });
      return;
    }
    res.status(200).json({ ok: true });
  } catch (err: any) {
    res.status(502).json({ error: err?.message || 'Google Chatへの送信に失敗しました' });
  }
}
