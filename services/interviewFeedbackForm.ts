// 候補者向けの「面接振り返りアンケート」（Googleフォーム）との連携。
//
// フォームと回答スプレッドシートは東原のbloom-firmアカウントで tools/create-interview-feedback-form.gs
// から作成した（2026-10-15）。担当者は選考ごとに事前入力付きリンク（受付番号＝応募ID・企業名・
// フェーズ入り）をメール／LINEで候補者に送り、回答は回答スプレッドシートに溜まる。kpi-mgrは
// それを読み、受付番号で選考に紐づけて候補者の面談ログに取り込む。取り込んだログは、企業ナレッジ
// 投稿の「AIで下書き」の材料にもなる。
//
// 回答スプレッドシートを他のメンバーが読むには閲覧権限が必要（無いと取り込み一覧が出ないだけ）。

import { googleFetch } from './interviewPrepData';

export const FEEDBACK_FORM_RESPONDER_URL =
  'https://docs.google.com/forms/d/e/1FAIpQLSdj3HPCvNbh7p9COmZfmLUmpfhTJCmB8jbZWoKEODGl_hS7BA/viewform';
export const FEEDBACK_RESPONSE_SPREADSHEET_ID = '1PVwq1Dqw1AkfWGqGcQk7VVVEZGgpkrJLBYQ48ZJ4CKM';
const ENTRY_APPLICATION_ID = '1841117153';
const ENTRY_COMPANY = '1783941001';
const ENTRY_PHASE = '1671602610';

// フォームの質問タイトル（回答スプレッドシートの見出しになる）。フォーム側で文言を変えたらここも直す。
const COL = {
  timestamp: 'タイムスタンプ',
  applicationId: '受付番号',
  company: '企業名',
  phase: '選考フェーズ',
  interviewDate: '面接日',
} as const;
// 面談ログに書き出す質問（この順で、回答があるものだけ）。
const DETAIL_COLUMNS = [
  '面接の形式', '所要時間', '面接官の人数・役職', '聞かれた質問', 'うまく答えられた点',
  '答えにくかった点・準備不足だと感じた点', '面接官の反応・雰囲気', '逆質問とその回答',
];
const OPTIONAL_COLUMNS = [
  '面接の手応え', 'この企業への志望度', '面接を受けて、志望度は変わりましたか',
  '他社の選考状況（内定・最終面接の予定など）', '担当者に伝えたいこと・不安な点',
];

export const buildFeedbackFormUrl = (app: { id: string; companyName: string; stage: string }): string => {
  const params = new URLSearchParams({ usp: 'pp_url' });
  params.set(`entry.${ENTRY_APPLICATION_ID}`, app.id);
  params.set(`entry.${ENTRY_COMPANY}`, app.companyName.trim());
  params.set(`entry.${ENTRY_PHASE}`, app.stage);
  return `${FEEDBACK_FORM_RESPONDER_URL}?${params.toString()}`;
};

/** 候補者に送る案内文。メールは件名付きの丁寧な文面、LINEは短め。 */
export const buildFeedbackInviteText = (
  channel: 'email' | 'line',
  params: { candidateName: string; companyName: string; stage: string; url: string },
): string => {
  const name = params.candidateName.trim() ? `${params.candidateName.trim()}様` : '';
  if (channel === 'line') {
    return [
      `${name ? `${name}\n` : ''}${params.companyName}の${params.stage}、お疲れさまでした！`,
      '今後の対策のため、面接で聞かれたことなどを5分ほどのアンケートで教えていただけますか？',
      params.url,
      '覚えている範囲で大丈夫です。よろしくお願いします。',
    ].join('\n');
  }
  return [
    `件名：【bloom】${params.companyName} ${params.stage}の振り返りアンケートのお願い`,
    '',
    name,
    '',
    'お世話になっております。bloomでございます。',
    `本日は${params.companyName}の${params.stage}、お疲れさまでした。`,
    '',
    '次の選考に向けた対策のため、面接で聞かれた質問や雰囲気などを、下記のアンケートで教えていただけますでしょうか。',
    '5分ほどで回答できます。覚えている範囲で構いません。',
    '',
    params.url,
    '',
    'ご回答いただいた内容は、今後の選考対策に活用させていただきます。',
    '何卒よろしくお願いいたします。',
  ].join('\n');
};

export interface FeedbackResponse {
  key: string; // 取り込み済み判定用（タイムスタンプ＋受付番号）
  applicationId: string;
  companyName: string;
  phase: string;
  submittedAt: string; // スプレッドシートのタイムスタンプ文字列のまま
  interviewDate: string;
  content: string; // 面談ログに入れる本文
}

const formatResponse = (row: Record<string, string>): string => {
  const lines: string[] = [];
  for (const col of DETAIL_COLUMNS) if (row[col]?.trim()) lines.push(`■${col}\n${row[col].trim()}`);
  const optional = OPTIONAL_COLUMNS.filter(col => row[col]?.trim());
  if (optional.length > 0) {
    lines.push('――― 任意回答（確度の参考）―――');
    for (const col of optional) lines.push(`■${col}\n${row[col].trim()}`);
  }
  return lines.join('\n\n');
};

/** 回答スプレッドシートの最初のシート（フォームの回答 1）を読む。 */
export const loadFeedbackResponses = async (): Promise<FeedbackResponse[]> => {
  const json = await googleFetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${FEEDBACK_RESPONSE_SPREADSHEET_ID}/values/${encodeURIComponent('A:Z')}`,
  );
  const values: string[][] = json.values || [];
  if (values.length < 2) return [];
  const header = values[0];
  // タイムスタンプ列の見出しはスプレッドシートの言語設定で変わる（「タイムスタンプ」「Timestamp」）
  // ため、見つからなければフォームの回答シートで常に先頭にある1列目を使う。
  const timestampCol = header.find(h => h === COL.timestamp || /^timestamp$/i.test(h)) ?? header[0];
  return values.slice(1).flatMap((v, rowIndex) => {
    const row = Object.fromEntries(header.map((h, i) => [h, v[i] ?? '']));
    const applicationId = (row[COL.applicationId] || '').trim();
    if (!applicationId) return [];
    const timestamp = row[timestampCol] || `row${rowIndex + 2}`;
    return [{
      key: `form:${timestamp}_${applicationId}`,
      applicationId,
      companyName: row[COL.company] || '',
      phase: row[COL.phase] || '',
      submittedAt: row[timestampCol] || '',
      interviewDate: row[COL.interviewDate] || '',
      content: formatResponse(row),
    }];
  });
};

/** 「2026/10/15 9:12:34」などのタイムスタンプ／面接日を yyyy-mm-dd にする（読めなければ空）。 */
export const toIsoDate = (s: string): string => {
  const m = s.match(/(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  return m ? `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}` : '';
};
