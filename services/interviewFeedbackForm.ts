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
import { isFeedbackStage } from './feedbackLink';

// フォームURL・事前入力・短縮リンクは feedbackLink.ts（サーバー関数 api/q.ts と共用）。
export { FEEDBACK_INTERVIEW_STAGES, buildFeedbackFormUrl, buildShortFeedbackUrl } from './feedbackLink';

export const FEEDBACK_RESPONSE_SPREADSHEET_ID = '1PVwq1Dqw1AkfWGqGcQk7VVVEZGgpkrJLBYQ48ZJ4CKM';

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

// フォームの「選考フェーズ」欄には面接の段階（FEEDBACK_INTERVIEW_STAGES）以外——お見送り・
// 選考辞退・内定など、kpi-mgr内部の選考ステータス——を絶対に入れない。候補者に見えてしまうため。

/** 選考が今どういう状態か（案内文の書き分け用）。 */
export type FeedbackOutcome = 'in_progress' | 'rejected' | 'withdrawn';
export const feedbackOutcomeOf = (stage: string): FeedbackOutcome =>
  stage === 'お見送り' ? 'rejected' : (stage === '選考辞退' || stage === '内定承諾後辞退') ? 'withdrawn' : 'in_progress';

/**
 * アンケートで聞く面接の段階の初期値。お見送り・辞退なら「どの段階で終わったか」（exitedFromStage）、
 * 面接中のステージならそのステージ、それ以外（次の段階の日程調整中・内定など）なら直近に受けた面接。
 * 面接を一度も受けていない選考（打診・書類選考のみ）は null。
 */
export const defaultFeedbackStage = (app: {
  stage: string;
  exitedFromStage?: string;
  stageHistory?: { stage: string }[];
}): string | null => {
  if (feedbackOutcomeOf(app.stage) !== 'in_progress' && isFeedbackStage(app.exitedFromStage)) return app.exitedFromStage;
  if (isFeedbackStage(app.stage)) return app.stage;
  const past = [...(app.stageHistory || [])].reverse().find(h => isFeedbackStage(h.stage));
  return past ? past.stage : null;
};

/**
 * 候補者に送る案内文。メールは件名付きの丁寧な文面、LINEは短め。お見送りになった選考には
 * 結果に触れつつ受けてもらったことへのお礼から入り、「次の選考」のような表現は使わない。
 */
export const buildFeedbackInviteText = (
  channel: 'email' | 'line',
  params: { candidateName: string; companyName: string; interviewStage: string; outcome: FeedbackOutcome; url: string },
): string => {
  const name = params.candidateName.trim() ? `${params.candidateName.trim()}様` : '';
  const target = `${params.companyName}の${params.interviewStage}`;
  const opening = {
    in_progress: { line: `${target}、お疲れさまでした！`, email: `先日は${target}、お疲れさまでした。` },
    rejected: {
      line: `${target}、お疲れさまでした。今回は残念な結果となりましたが、選考に臨んでいただきありがとうございました。`,
      email: `先日は${target}にご参加いただき、ありがとうございました。今回は残念な結果となりましたが、真摯に選考に臨んでいただいたことに改めて御礼申し上げます。`,
    },
    withdrawn: { line: `${target}、ありがとうございました。`, email: `先日は${target}にご参加いただき、ありがとうございました。` },
  }[params.outcome];
  const purpose = {
    in_progress: { line: '今後の対策のため、', email: '今後の選考に向けた対策のため、' },
    rejected: { line: '今後の転職活動のサポートに活かすため、', email: '今後の転職活動のサポートにより活かしていくため、' },
    withdrawn: { line: '今後のサポートの参考にするため、', email: '今後のサポートの参考にさせていただくため、' },
  }[params.outcome];
  if (channel === 'line') {
    return [
      `${name ? `${name}\n` : ''}${opening.line}`,
      `${purpose.line}面接で聞かれたことなどを5分ほどのアンケートで教えていただけますか？`,
      params.url,
      '覚えている範囲で大丈夫です。よろしくお願いします。',
    ].join('\n');
  }
  return [
    `件名：【bloom】${target}についてのアンケートのお願い`,
    '',
    name,
    '',
    'お世話になっております。bloomでございます。',
    opening.email,
    '',
    `${purpose.email}面接で聞かれた質問や雰囲気などを、下記のアンケートで教えていただけますでしょうか。`,
    '5分ほどで回答できます。覚えている範囲で構いません。',
    '',
    params.url,
    '',
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
