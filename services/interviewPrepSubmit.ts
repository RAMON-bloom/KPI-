// kpi-mgrから面接対策ジェネレーターへ、企業の選考ナレッジ（面接で聞かれたこと・面接官・
// 雰囲気など）と選考結果を投稿する。
//
// ジェネレーターの既存の表（cases / phases など）には書かない。ジェネレーターには管理者の
// 承認フロー（pending_approvals）と投稿者ごとの貢献記録（contributions）があり、直接書くと
// それを素通りしてしまうため。代わりに同じスプレッドシートの専用シート `kpi_submissions` に
// 1行ずつ追記し、ジェネレーター側の承認画面で取り込んでもらう（status=pending → 管理者が
// approved / rejected にして reviewed_* を埋める想定）。シートが無ければ初回投稿時に作る。
// 既存の行は一切変更しない（追記のみ）。
//
// 候補者の氏名・連絡先は送らない。AI下書きでも氏名を「候補者」に置き換える。

import { GoogleGenAI, Type } from '@google/genai';
import { googleFetch, INTERVIEW_PREP_SPREADSHEET_ID, companyNamesEqual, type InterviewPrepData } from './interviewPrepData';

export const SUBMISSIONS_SHEET = 'kpi_submissions';

export const SUBMISSION_COLUMNS = [
  'id', 'status', 'kind', 'submitted_at', 'submitted_by',
  'company_id', 'company_name', 'phase', 'position',
  'result', 'reason', 'interviewer', 'content', 'candidate_attr',
  'kpi_application_id', 'reviewed_at', 'reviewed_by', 'review_note',
] as const;

// ジェネレーター側のフェーズ・結果の表記（cases / mail_proposals と同じ）。
export const GENERATOR_PHASES = ['書類選考', '適性検査', '一次面接', '二次面接', '最終面接', 'オファー面談'] as const;
export const GENERATOR_RESULTS = ['通過', '不採用', '辞退', '内定', '保留'] as const;

export interface KnowledgeSubmission {
  companyName: string;
  phase: string;
  position: string;
  result: string; // '' = 結果なし（ナレッジのみ）
  reason: string;
  interviewer: string;
  content: string;
  candidateAttr: string;
  applicationId: string;
  submittedBy: string;
}

const SHEETS_BASE = `https://sheets.googleapis.com/v4/spreadsheets/${INTERVIEW_PREP_SPREADSHEET_ID}`;

const ensureSubmissionsSheet = async (): Promise<void> => {
  const meta = await googleFetch(`${SHEETS_BASE}?fields=sheets.properties.title`);
  const exists = (meta.sheets || []).some((sh: any) => sh.properties?.title === SUBMISSIONS_SHEET);
  if (exists) return;
  try {
    await googleFetch(`${SHEETS_BASE}:batchUpdate`, {
      method: 'POST',
      body: JSON.stringify({ requests: [{ addSheet: { properties: { title: SUBMISSIONS_SHEET } } }] }),
    });
  } catch (err: any) {
    // 同時に別の人が作った場合は「既に存在」で失敗するので、それは成功扱い。
    if (!/already exists/i.test(err?.body || '')) throw err;
    return;
  }
  await googleFetch(`${SHEETS_BASE}/values/${encodeURIComponent(`${SUBMISSIONS_SHEET}!A1`)}?valueInputOption=RAW`, {
    method: 'PUT',
    body: JSON.stringify({ values: [SUBMISSION_COLUMNS as unknown as string[]] }),
  });
};

const newSubmissionId = () => `ks_${Date.now()}_${Math.floor(Math.random() * 10000)}`;

/** 投稿を1行追記する。ジェネレーターの企業マスタに完全一致する企業があれば company_id も入れる。 */
export const submitKnowledge = async (sub: KnowledgeSubmission, data: InterviewPrepData | null): Promise<string> => {
  await ensureSubmissionsSheet();
  let companyId = '';
  if (data) {
    const ids: string[] = [];
    data.companies.forEach((name, id) => { if (companyNamesEqual(name, sub.companyName)) ids.push(id); });
    if (ids.length === 1) companyId = ids[0];
  }
  const id = newSubmissionId();
  const kind = sub.content.trim() ? (sub.result ? 'knowledge+result' : 'knowledge') : 'result';
  const row: Record<(typeof SUBMISSION_COLUMNS)[number], string> = {
    id, status: 'pending', kind, submitted_at: new Date().toISOString(), submitted_by: sub.submittedBy,
    company_id: companyId, company_name: sub.companyName.trim(), phase: sub.phase, position: sub.position.trim(),
    result: sub.result, reason: sub.reason.trim(), interviewer: sub.interviewer.trim(), content: sub.content.trim(),
    candidate_attr: sub.candidateAttr.trim(), kpi_application_id: sub.applicationId,
    reviewed_at: '', reviewed_by: '', review_note: '',
  };
  await googleFetch(
    `${SHEETS_BASE}/values/${encodeURIComponent(`${SUBMISSIONS_SHEET}!A1`)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
    { method: 'POST', body: JSON.stringify({ values: [SUBMISSION_COLUMNS.map(c => row[c])] }) },
  );
  return id;
};

// --- AI下書き ---

const MODEL = 'gemini-2.5-flash';

/** 氏名（フルネーム・姓・名）とメール・電話番号を伏せる。 */
export const maskPersonalInfo = (text: string, candidateName: string): string => {
  let out = text
    .replace(/[\w.+-]+@[\w-]+\.[\w.-]+/g, '（メール）')
    .replace(/0\d{1,4}-?\d{1,4}-?\d{3,4}/g, '（電話番号）');
  const tokens = candidateName.normalize('NFKC').trim().split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return out;
  const escape = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  // フルネームは姓名の間の空白（全角・半角・無し）を問わず1つの「候補者」にする。
  out = out.replace(new RegExp(tokens.map(escape).join('\\s*'), 'g'), '候補者');
  for (const part of tokens.filter(t => t.length >= 2).sort((a, b) => b.length - a.length)) {
    out = out.split(part).join('候補者');
  }
  return out;
};

export interface KnowledgeDraft {
  content: string;
  interviewer: string;
  candidateAttr: string;
  result: string;
  reason: string;
}

/**
 * 面接後の振り返り面談のログ（議事録の要約・メモ）から、企業ナレッジの下書きを作る。
 * ログは送信前に氏名・連絡先を伏せてからAIに渡す。
 */
export const draftKnowledgeFromMemo = async (params: {
  memoText: string;
  candidateName: string;
  companyName: string;
  phase: string;
}): Promise<KnowledgeDraft> => {
  const masked = maskPersonalInfo(params.memoText, params.candidateName);
  const instruction = [
    'あなたは人材紹介会社のキャリアアドバイザーです。次の「面談ログ」は、候補者が企業の面接を受けた後に担当者と行った振り返り面談などの記録です。',
    `ここから「${params.companyName}」の「${params.phase}」について、今後同じ企業を受ける別の候補者の面接対策に使える企業ナレッジを抜き出してください。`,
    'ルール:',
    '- 候補者個人を特定できる情報（氏名・現職の社名・年齢の具体値・連絡先）は書かない。候補者は「候補者」と呼ぶ。',
    '- ログに書かれていないことは推測で書かない。該当する情報が無い項目は空文字にする。',
    '- content は次の見出しで、箇条書き中心に書く（情報が無い見出しは省く）:',
    '  ■形式・所要時間 ■面接官 ■聞かれた質問 ■評価されたポイント・懸念された点 ■逆質問とその回答 ■雰囲気・その他',
    '- interviewer は面接官の役職・人数（例:「事業部長1名、人事1名」）。個人名がわかる場合も役職だけにする。',
    '- candidateAttr は候補者のタイプを個人が特定できない粒度で（例:「30代前半・事業会社の経理経験者」）。',
    `- result はログから選考結果が読み取れる場合のみ ${GENERATOR_RESULTS.join('／')} のいずれか。わからなければ空文字。reason はその理由（企業側の評価）。`,
    '',
    '■面談ログ',
    masked,
  ].join('\n');
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY as string });
  const response = await ai.models.generateContent({
    model: MODEL,
    contents: instruction,
    config: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          content: { type: Type.STRING },
          interviewer: { type: Type.STRING },
          candidateAttr: { type: Type.STRING },
          result: { type: Type.STRING },
          reason: { type: Type.STRING },
        },
        required: ['content', 'interviewer', 'candidateAttr', 'result', 'reason'],
      },
    },
  });
  const parsed = JSON.parse((response.text || '{}').trim());
  const clean = (v: unknown) => maskPersonalInfo(typeof v === 'string' ? v : '', params.candidateName);
  const result = (GENERATOR_RESULTS as readonly string[]).includes(parsed.result) ? parsed.result : '';
  return {
    content: clean(parsed.content),
    interviewer: clean(parsed.interviewer),
    candidateAttr: clean(parsed.candidateAttr),
    result,
    reason: clean(parsed.reason),
  };
};
