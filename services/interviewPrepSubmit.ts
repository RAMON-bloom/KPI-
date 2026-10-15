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
import { googleFetch, INTERVIEW_PREP_SPREADSHEET_ID, companyNamesEqual, companyNamesMatch, type InterviewPrepData } from './interviewPrepData';

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
  // 選考の企業名をマスタから選んでいれば、その企業ID（名前での照合より優先）。
  gasCompanyId?: string;
}

const SHEETS_BASE = `https://sheets.googleapis.com/v4/spreadsheets/${INTERVIEW_PREP_SPREADSHEET_ID}`;

const colLetter = (i: number): string => (i >= 26 ? colLetter(Math.floor(i / 26) - 1) : '') + String.fromCharCode(65 + (i % 26));

/**
 * KPI管理くん専用のシート（kpi_submissions / kpi_handoffs）が無ければ見出し付きで作る。
 * 既にある場合は、こちらで新しく使い始めた列の見出しが無ければ右端に足す（既存の列は動かさない）。
 */
const ensureSheet = async (title: string, columns: readonly string[]): Promise<void> => {
  const meta = await googleFetch(`${SHEETS_BASE}?fields=sheets.properties.title`);
  const exists = (meta.sheets || []).some((sh: any) => sh.properties?.title === title);
  if (exists) {
    const header: string[] = ((await googleFetch(`${SHEETS_BASE}/values/${encodeURIComponent(`${title}!1:1`)}`)).values || [[]])[0] || [];
    const missing = columns.filter(c => !header.includes(c));
    if (missing.length > 0) {
      const start = colLetter(header.length);
      await googleFetch(`${SHEETS_BASE}/values/${encodeURIComponent(`${title}!${start}1`)}?valueInputOption=RAW`, {
        method: 'PUT',
        body: JSON.stringify({ values: [missing] }),
      });
    }
    return;
  }
  try {
    await googleFetch(`${SHEETS_BASE}:batchUpdate`, {
      method: 'POST',
      body: JSON.stringify({ requests: [{ addSheet: { properties: { title } } }] }),
    });
  } catch (err: any) {
    // 同時に別の人が作った場合は「既に存在」で失敗するので、それは成功扱い。
    if (!/already exists/i.test(err?.body || '')) throw err;
    return;
  }
  await googleFetch(`${SHEETS_BASE}/values/${encodeURIComponent(`${title}!A1`)}?valueInputOption=RAW`, {
    method: 'PUT',
    body: JSON.stringify({ values: [[...columns]] }),
  });
};
const ensureSubmissionsSheet = () => ensureSheet(SUBMISSIONS_SHEET, SUBMISSION_COLUMNS);

/**
 * シートの見出し行に合わせて1行追記する。ジェネレーター側が列を足したり並べ替えたりしても
 * （2026-10-15時点で kpi_submissions に case_id / phase_id が追加されている）、値が別の列に
 * ずれないようにするため。見出しに無い項目は書かず、こちらが知らない列は空欄のままにする。
 */
const appendRowByHeader = async (sheet: string, fields: Record<string, string>): Promise<void> => {
  const header: string[] = ((await googleFetch(`${SHEETS_BASE}/values/${encodeURIComponent(`${sheet}!1:1`)}`)).values || [[]])[0] || [];
  const missing = Object.keys(fields).filter(k => !header.includes(k));
  if (missing.length > 0) console.warn(`${sheet} の見出しに無い項目は書き込みませんでした:`, missing);
  await googleFetch(
    `${SHEETS_BASE}/values/${encodeURIComponent(`${sheet}!A1`)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
    { method: 'POST', body: JSON.stringify({ values: [header.map(h => fields[h] ?? '')] }) },
  );
};

const newSubmissionId = () => `ks_${Date.now()}_${Math.floor(Math.random() * 10000)}`;

/** 投稿を1行追記する。ジェネレーターの企業マスタに完全一致する企業があれば company_id も入れる。 */
export const submitKnowledge = async (sub: KnowledgeSubmission, data: InterviewPrepData | null): Promise<string> => {
  await ensureSubmissionsSheet();
  let companyId = '';
  // 企業名を後から書き換えてマスタの名前と合わなくなったIDは使わない。
  if (data && sub.gasCompanyId && companyNamesMatch(data.companies.get(sub.gasCompanyId) || '', sub.companyName)) {
    companyId = sub.gasCompanyId;
  } else if (data) {
    const ids: string[] = [];
    data.companies.forEach((name, id) => { if (companyNamesEqual(name, sub.companyName)) ids.push(id); });
    if (ids.length === 1) companyId = ids[0];
  }
  const id = newSubmissionId();
  const kind = sub.content.trim() ? (sub.result ? 'knowledge+result' : 'knowledge') : 'result';
  const row: Record<(typeof SUBMISSION_COLUMNS)[number], string> = {
    id, status: 'pending', kind, submitted_at: new Date().toISOString(), submitted_by: sub.submittedBy,
    company_id: companyId, company_name: sub.companyName.trim(), phase: sub.phase, position: sub.position.trim(),
    result: sub.result, reason: sub.reason.trim(), interviewer: sub.interviewer.trim(), content: fitSheetCell(sub.content.trim()),
    candidate_attr: sub.candidateAttr.trim(), kpi_application_id: sub.applicationId,
    reviewed_at: '', reviewed_by: '', review_note: '',
  };
  await appendRowByHeader(SUBMISSIONS_SHEET, row);
  return id;
};

// --- 生成画面への受け渡し（自動生成用） ---
//
// 「面接対策ジェネレーターで作る」を押した時、面談メモ（氏名・連絡先は伏せる）を専用シート
// `kpi_handoffs` に一時的に置き、ジェネレーターのURLに受け渡し番号（handoff）を付けて開く。
// ジェネレーター側が対応すれば、番号でメモを読み込んで、企業・フェーズ・メモ入りで自動的に
// 生成を始める（開発者への依頼「お願い6」）。対応前はURLのパラメータが無視されるだけ。
// メモは個人を特定しにくい形にしてあるが、ジェネレーター側には読み込んだら memo を消して
// status=used にしてもらう（24時間で期限切れ）。

export const HANDOFFS_SHEET = 'kpi_handoffs';
export const HANDOFF_COLUMNS = [
  'token', 'status', 'created_at', 'expires_at', 'created_by',
  'kpi_candidate_id', 'kpi_application_id', 'company_id', 'company_name', 'phase', 'position',
  'interview_memo', 'used_at',
  // 候補者名（ジェネレーターの候補者名欄に入れる。ジェネレーターは画面上の「あなた」の置き換えにだけ
  // 使い、AIには送らない）。面談メモと同じく、ジェネレーターが読み込んだら消す。
  'candidate_name',
  // 候補者の登録情報（現職・職種・学歴・年齢・現年収・経歴概要）をレジュメ代わりのテキストにしたもの。
  // kpi-mgrにはレジュメのファイル本体が無いため。氏名・連絡先は含めない。読んだら消す。
  'resume_text',
] as const;
const HANDOFF_TTL_MS = 24 * 60 * 60 * 1000;

// Googleスプレッドシートの1セルは最大50,000文字。面談ログは詳しく残す方針なので、何件分も
// つなぐと超えることがある。超える場合は、時系列で後ろ（＝新しい面談）の方を残す。
const SHEET_CELL_LIMIT = 49000;
const fitSheetCell = (text: string): string =>
  text.length <= SHEET_CELL_LIMIT ? text : `（文字数の上限のため、古い面談ログの一部を省略しています）\n\n${text.slice(text.length - SHEET_CELL_LIMIT + 60)}`;

/** 推測されにくい受け渡し番号（URLに載る）。 */
export const newHandoffToken = (): string => {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  return 'kh_' + Array.from(bytes, b => b.toString(36).padStart(2, '0')).join('');
};

export const createHandoff = async (h: {
  token: string;
  createdBy: string;
  candidateId: string;
  candidateName: string;
  applicationId: string;
  gasCompanyId?: string;
  companyName: string;
  phase: string;
  memoText: string;
  resumeText?: string;
}): Promise<void> => {
  await ensureSheet(HANDOFFS_SHEET, HANDOFF_COLUMNS);
  const now = Date.now();
  const row: Record<(typeof HANDOFF_COLUMNS)[number], string> = {
    token: h.token, status: 'pending', created_at: new Date(now).toISOString(),
    expires_at: new Date(now + HANDOFF_TTL_MS).toISOString(), created_by: h.createdBy,
    kpi_candidate_id: h.candidateId, kpi_application_id: h.applicationId,
    company_id: h.gasCompanyId || '', company_name: h.companyName.trim(), phase: h.phase, position: '',
    // 面談メモは氏名・メール・電話番号を伏せてから置く（ジェネレーター側の自動マスクも通る）。
    interview_memo: fitSheetCell(maskPersonalInfo(h.memoText, h.candidateName)),
    used_at: '',
    candidate_name: h.candidateName.trim(),
    resume_text: fitSheetCell(maskPersonalInfo(h.resumeText || '', h.candidateName)),
  };
  await appendRowByHeader(HANDOFFS_SHEET, row);
};

// --- 企業の新規登録 ---

/** 新規登録する企業名の表記をそろえる（全角英数→半角、空白の連続を1つに）。 */
export const cleanCompanyName = (name: string): string =>
  name.normalize('NFKC').replace(/\s+/g, ' ').trim();

/**
 * 面接対策ジェネレーターの企業マスタ（companies シート）に企業を1件追加し、その企業IDを返す。
 * ジェネレーター自身も、求人票の貼り付け時に未登録企業を承認なしで自動登録しているので、
 * それと同じ形式（id = c_<ミリ秒>_<数字>、created_at / updated_at = ISO）で1行追記する。
 * 列は見出し行の名前に合わせて並べる（列の並び替え・追加に耐える）。既存の行は変更しない。
 * 追記の直前にマスタを読み直し、同じ企業（表記ゆれ込み）が既にあればそのIDを返す。
 */
export const registerCompany = async (rawName: string): Promise<{ id: string; name: string; created: boolean }> => {
  const name = cleanCompanyName(rawName);
  if (!name) throw new Error('企業名が空です。');
  const fullHeader: string[] = ((await googleFetch(`${SHEETS_BASE}/values/${encodeURIComponent('companies!1:1')}`)).values || [[]])[0];
  const idCol = fullHeader.indexOf('id');
  const nameCol = fullHeader.indexOf('name');
  if (idCol < 0 || nameCol < 0) throw new Error('ジェネレーターの企業マスタの形式が想定と違うため、登録できませんでした。');
  // id列とname列だけを読み直す（memo列などの長い本文は読まない）。
  const ranges = [idCol, nameCol].map(i => `ranges=${encodeURIComponent(`companies!${colLetter(i)}:${colLetter(i)}`)}`).join('&');
  const cols = (await googleFetch(`${SHEETS_BASE}/values:batchGet?${ranges}&majorDimension=COLUMNS`)).valueRanges || [];
  const ids: string[] = (cols[0]?.values?.[0] || []).slice(1);
  const names: string[] = (cols[1]?.values?.[0] || []).slice(1);
  const dupIndex = names.findIndex(n => n && companyNamesEqual(n, name));
  if (dupIndex >= 0) return { id: ids[dupIndex], name: names[dupIndex], created: false };

  const now = new Date().toISOString();
  const id = `c_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
  const fields: Record<string, string> = { id, name, created_at: now, updated_at: now };
  const row = fullHeader.map(h => fields[h] ?? '');
  await googleFetch(
    `${SHEETS_BASE}/values/${encodeURIComponent('companies!A1')}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
    { method: 'POST', body: JSON.stringify({ values: [row] }) },
  );
  return { id, name, created: true };
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
    '- 要約しすぎない。次に同じ企業を受ける候補者がそのまま準備に使えるよう、ログにある情報は省略せず具体的に書く。',
    '- 聞かれた質問は、ログに出たものをすべて、聞かれた言い回しにできるだけ近い形で列挙する。深掘りされた点や、候補者の回答に対する面接官の反応も書く。',
    '- 数字（所要時間・人数・選考の段階数など）や、企業側が重視していると分かる点は必ず残す。',
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

/**
 * kpi-mgrの登録情報から、レジュメ代わりのプロフィール文を作る（面接対策ジェネレーターのレジュメ欄用）。
 * 氏名・電話番号・メールアドレスは入れない。登録が無い項目は省く。何も無ければ空文字。
 */
export const buildCandidateProfileText = (c: {
  age?: number; currentCompany?: string; jobType?: string; education?: string;
  currentSalary?: number; summary?: string;
}): string => {
  const lines: string[] = [];
  if (c.age) lines.push(`・年齢：${c.age}歳`);
  if (c.currentCompany?.trim()) lines.push(`・現職：${c.currentCompany.trim()}`);
  if (c.jobType?.trim()) lines.push(`・職種：${c.jobType.trim()}`);
  if (c.education?.trim()) lines.push(`・最終学歴：${c.education.trim()}`);
  if (c.currentSalary) lines.push(`・現年収：${c.currentSalary}万円`);
  const summary = c.summary?.trim();
  if (lines.length === 0 && !summary) return '';
  return [
    '【候補者プロフィール（KPI管理くんの登録情報。レジュメ原本ではありません）】',
    ...lines,
    ...(summary ? ['', '【経歴概要】', summary] : []),
  ].join('\n');
};
