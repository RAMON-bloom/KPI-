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

  const { webhookUrl, text, threadKey } = req.body || {};
  if (typeof webhookUrl !== 'string' || typeof text !== 'string' || !text.trim()) {
    res.status(400).json({ error: 'webhookUrl and text are required' });
    return;
  }
  if (threadKey !== undefined && typeof threadKey !== 'string') {
    res.status(400).json({ error: 'threadKey must be a string' });
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

  // threadKeyを指定すると、同じキーの初回投稿でスレッドが作られ、以降の投稿はそのスレッドへの
  // 返信として投稿される（Chat REST APIのspaces.messages.create — 受信Webhookも同じエンド
  // ポイントなので同じ仕組みが使える）。messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD
  // で「既存スレッドがあれば返信、なければ新規スレッドを作る」動作にする。
  if (threadKey) {
    parsed.searchParams.set('messageReplyOption', 'REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD');
  }

  try {
    const chatResponse = await fetch(parsed.toString(), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=UTF-8' },
      body: JSON.stringify(threadKey ? { text, thread: { threadKey } } : { text }),
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
