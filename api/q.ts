// 面接振り返りアンケートの短縮リンク /q/<code> を、事前入力付きのGoogleフォームへ転送する。
// （vercel.json の rewrite で /q/:code → /api/q?code=:code）
// 転送先は常に固定のフォームURLで、コードから復元した受付番号・企業名・面接の段階を
// 事前入力として付けるだけ。何も保存しない。
import { decodeFeedbackCode, buildFeedbackFormUrl } from '../services/feedbackLink';

export default function handler(req: any, res: any) {
  const code = typeof req.query?.code === 'string' ? req.query.code : '';
  const decoded = code ? decodeFeedbackCode(code) : null;
  if (!decoded) {
    res.status(400).setHeader('Content-Type', 'text/html; charset=utf-8');
    res.end('<!doctype html><meta charset="utf-8"><title>リンクが正しくありません</title><p>リンクが正しくありません。お手数ですが、担当者にご連絡ください。</p>');
    return;
  }
  const url = buildFeedbackFormUrl({ id: decoded.applicationId, companyName: decoded.companyName }, decoded.interviewStage);
  res.statusCode = 302;
  res.setHeader('Location', url);
  res.setHeader('Cache-Control', 'no-store');
  res.end();
}
