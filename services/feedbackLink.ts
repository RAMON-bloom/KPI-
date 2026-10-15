// 面接振り返りアンケートの事前入力URLと、その短縮リンク（/q/<code>）の組み立て・復元。
// ブラウザ（index.tsx）とVercelのサーバー関数（api/q.ts）の両方から使うので、他のモジュールに
// 依存しない（googleAuth などブラウザ専用のコードを読み込まない）こと。
//
// 短縮リンクは外部の短縮サービスを使わず、必要な情報（受付番号＝応募ID・面接の段階・企業名）を
// コードそのものに詰めてある。サーバー側は復元してフォームへリダイレクトするだけで、データを
// どこにも保存しない。

export const FEEDBACK_FORM_RESPONDER_URL =
  'https://docs.google.com/forms/d/e/1FAIpQLSdj3HPCvNbh7p9COmZfmLUmpfhTJCmB8jbZWoKEODGl_hS7BA/viewform';
const ENTRY_APPLICATION_ID = '1841117153';
const ENTRY_COMPANY = '1783941001';
const ENTRY_PHASE = '1671602610';

// アンケートの対象にできる（＝候補者が実際に受ける）面接の段階。並び順は短縮コードの中で
// 番号として使うので、変更するときは末尾に追加するだけにすること（既に送ったリンクが壊れる）。
export const FEEDBACK_INTERVIEW_STAGES = ['カジュアル面談', '1次面接', '2次面接', '最終面接', 'オファー面談'] as const;
export const isFeedbackStage = (s: string | undefined): s is string =>
  !!s && (FEEDBACK_INTERVIEW_STAGES as readonly string[]).includes(s);

export const buildFeedbackFormUrl = (app: { id: string; companyName: string }, interviewStage: string): string => {
  const params = new URLSearchParams({ usp: 'pp_url' });
  params.set(`entry.${ENTRY_APPLICATION_ID}`, app.id);
  params.set(`entry.${ENTRY_COMPANY}`, app.companyName.trim());
  params.set(`entry.${ENTRY_PHASE}`, isFeedbackStage(interviewStage) ? interviewStage : '');
  return `${FEEDBACK_FORM_RESPONDER_URL}?${params.toString()}`;
};

// --- 短縮コード ---

const toBase64Url = (bytes: Uint8Array): string => {
  let bin = '';
  bytes.forEach(b => { bin += String.fromCharCode(b); });
  return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
};
const fromBase64Url = (s: string): Uint8Array => {
  const bin = atob(s.replace(/-/g, '+').replace(/_/g, '/'));
  return Uint8Array.from(bin, c => c.charCodeAt(0));
};

// 法人格の表記は短縮のために落とす（候補者には「FLUX」のように表示される）。
const shortenCompany = (name: string): string =>
  name.trim().replace(/株式会社|有限会社|合同会社|（株）|\(株\)|㈱/g, '').trim() || name.trim();

/**
 * 応募ID「app_<ミリ秒>_<英数字>」はミリ秒を36進数にして詰める。それ以外の形式のIDはそのまま入れる。
 * 中身: <段階番号1桁><IDトークン>~<企業名> を UTF-8 → base64url。
 */
export const encodeFeedbackCode = (applicationId: string, interviewStage: string, companyName: string): string | null => {
  const stageIndex = (FEEDBACK_INTERVIEW_STAGES as readonly string[]).indexOf(interviewStage);
  if (stageIndex < 0) return null;
  const m = applicationId.match(/^app_(\d+)_([a-z0-9]+)$/);
  const idToken = m ? `${Number(m[1]).toString(36)}.${m[2]}` : `!${applicationId}`;
  const payload = `${stageIndex}${idToken}~${shortenCompany(companyName)}`;
  return toBase64Url(new TextEncoder().encode(payload));
};

export const decodeFeedbackCode = (code: string): { applicationId: string; interviewStage: string; companyName: string } | null => {
  try {
    const payload = new TextDecoder().decode(fromBase64Url(code));
    const m = payload.match(/^(\d)([^~]+)~(.*)$/s);
    if (!m) return null;
    const interviewStage = FEEDBACK_INTERVIEW_STAGES[Number(m[1])];
    if (!interviewStage) return null;
    let applicationId: string;
    if (m[2].startsWith('!')) {
      applicationId = m[2].slice(1);
    } else {
      const [ms36, rand] = m[2].split('.');
      if (!ms36 || !rand) return null;
      applicationId = `app_${parseInt(ms36, 36)}_${rand}`;
    }
    return { applicationId, interviewStage, companyName: m[3] };
  } catch {
    return null;
  }
};

/** 候補者に送る短縮リンク（kpi-mgr自身の /q/<code> がフォームへ転送する）。 */
export const buildShortFeedbackUrl = (origin: string, app: { id: string; companyName: string }, interviewStage: string): string => {
  const code = encodeFeedbackCode(app.id, interviewStage, app.companyName);
  return code ? `${origin}/q/${code}` : buildFeedbackFormUrl(app, interviewStage);
};
