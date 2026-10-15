// 面接対策ジェネレーター（GAS）のデータをkpi-mgr側から「読むだけ」で活用する。
//
// GASはデータベースとしてスプレッドシート（bloom_interview_v3_clean）を使っており、社内の
// メンバーには閲覧権限が付いている。kpi-mgrはログイン時にdriveスコープを取っているので、
// Sheets APIでそのまま読める（GAS側の改修は不要）。読むのは連携に必要なシートと列だけ
// （companies / cases / phases / mail_proposals）で、users・mail_inbox・auditなどは読まない。
// 書き込みは一切しない — GAS側のデータの持ち方を壊さないため。
//
// あわせて、GASが生成した面接対策資料（ログインユーザー自身のドライブに
// 「面接対策_企業名_ポジション_フェーズ_日付.docx」の名前で保存される）をドライブ検索で拾う。
//
// シートの列名が変わったら、ここ（readSheet呼び出しの列名）を直す。列は見出し行の名前で
// 引くので、列の並び替え・追加には耐える。

import { getCurrentSession, refreshTokenSilently } from './googleAuth';

export const INTERVIEW_PREP_SPREADSHEET_ID = '1wLVLb9uGdrjJf5mkc9BNHeI3rg-rOvXvskgZaK6HOOE';

// kpi-mgrのステージ ↔ GASのフェーズ表記（データ照合用。対策資料ボタンの対象判定は
// interviewPrep.tsのtoGeneratorPhaseで別に持つ — 書類選考・適性検査は資料の対象外のため）。
const KPI_STAGE_BY_GAS_PHASE: Record<string, string> = {
  '書類選考': '書類選考',
  '適性検査': '適性検査',
  'カジュアル面談': 'カジュアル面談',
  '一次面接': '1次面接',
  '二次面接': '2次面接',
  '最終面接': '最終面接',
  'オファー面談': 'オファー面談',
};
const GAS_PHASE_BY_KPI_STAGE: Record<string, string> = Object.fromEntries(
  Object.entries(KPI_STAGE_BY_GAS_PHASE).map(([gas, kpi]) => [kpi, gas]),
);
export const gasPhaseForStage = (stage: string): string | null => GAS_PHASE_BY_KPI_STAGE[stage] ?? null;
export const kpiStageForGasPhase = (phase: string): string | null => KPI_STAGE_BY_GAS_PHASE[phase] ?? null;

// --- 名寄せ ---

const COMPANY_NOISE = /株式会社|（株）|\(株\)|㈱|有限会社|合同会社|ホールディングス|グループ|[\s・･\.,，、\-‐－_＿]/g;
export const normalizeCompanyName = (name: string): string =>
  name.normalize('NFKC').replace(COMPANY_NOISE, '').toLowerCase();

const companyNameVariants = (name: string): string[] =>
  name.split(/[\/／]/).map(normalizeCompanyName).filter(s => s.length >= 2);

/** 企業名の完全一致（正規化後）。GAS側は「A / B」のように複数表記を持つことがあるので分割して比べる。 */
export const companyNamesEqual = (a: string, b: string): boolean => {
  const bs = companyNameVariants(b);
  return companyNameVariants(a).some(x => bs.includes(x));
};

/** 企業名のゆるい一致（片方がもう片方を含む）。「ベイカレント」と「ベイカレント・コンサルティング」など。 */
export const companyNamesMatch = (a: string, b: string): boolean => {
  const bs = companyNameVariants(b);
  return companyNameVariants(a).some(x => bs.some(y => x === y || (x.length >= 3 && y.includes(x)) || (y.length >= 3 && x.includes(y))));
};

export const normalizePersonName = (name: string): string =>
  name.normalize('NFKC').replace(/[\s　]/g, '').replace(/(様|さん|氏)$/, '');

// --- 取得 ---

export class InterviewPrepAccessError extends Error {}

export async function googleFetch(url: string, init: RequestInit = {}): Promise<any> {
  let session = getCurrentSession();
  if (!session) {
    try { await refreshTokenSilently(); } catch { /* 下で判定 */ }
    session = getCurrentSession();
    if (!session) throw new Error('ログインが必要です。');
  }
  const res = await fetch(url, {
    ...init,
    cache: 'no-store',
    headers: { ...(init.headers || {}), Authorization: `Bearer ${session.accessToken}`, ...(init.body ? { 'Content-Type': 'application/json' } : {}) },
  });
  if (res.status === 403 || res.status === 404) {
    const body = await res.text().catch(() => '');
    // kpi-mgrのGoogle CloudプロジェクトでSheets APIが有効になっていない場合も403になる。
    if (/SERVICE_DISABLED|has not been used|is disabled/.test(body)) {
      throw new InterviewPrepAccessError('Google Sheets APIが有効になっていないため、面接対策ジェネレーターのデータを読めません（管理者による設定が必要です）。');
    }
    throw new InterviewPrepAccessError(init.method && init.method !== 'GET'
      ? '面接対策ジェネレーターのスプレッドシートへの書き込み権限がありません。'
      : '面接対策ジェネレーターのデータを閲覧する権限がありません。');
  }
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    const error = new Error(`Google APIエラー（${res.status}）`) as Error & { body?: string };
    error.body = body;
    throw error;
  }
  return res.json();
}

type Row = Record<string, string>;

/** 見出し行の名前で列を引いて、行をオブジェクトにする。 */
const toRows = (values: string[][] | undefined, columns: string[]): Row[] => {
  if (!values || values.length === 0) return [];
  const header = values[0];
  const idx = columns.map(c => header.indexOf(c));
  return values.slice(1).map(v => Object.fromEntries(columns.map((c, i) => [c, idx[i] >= 0 ? (v[idx[i]] ?? '') : ''])));
};

export interface PrepCase { companyId: string; phase: string; result: string; createdAt: string }
export interface PrepKnowledge { companyId: string; phase: string; position: string; content: string; updatedAt: string }
export interface MailProposal {
  id: string;
  companyId: string;
  companyHint: string;
  candidate: string;
  phase: string;
  result: string;
  reason: string;
  confidence: string;
  status: string;
  createdAt: string;
}
export interface PrepDoc { id: string; name: string; url: string; createdTime: string; company: string; position: string; phase: string; date: string }

export interface InterviewPrepData {
  companies: Map<string, string>; // company_id → 企業名
  cases: PrepCase[];
  knowledge: PrepKnowledge[];
  proposals: MailProposal[];
  docs: PrepDoc[];
  // スプレッドシートだけ読めなかった場合の理由（資料の検索は別に成功していることがある）。
  sheetsError?: string;
  loadedAt: number;
}

const loadSheets = async () => {
  // 見出しを含めて列全体を取る（A:Zで足りる。読むシートだけ指定）。
  const ranges = ['companies!A:Z', 'cases!A:Z', 'phases!A:Z', 'mail_proposals!A:Z'];
  const qs = ranges.map(r => `ranges=${encodeURIComponent(r)}`).join('&');
  const json = await googleFetch(`https://sheets.googleapis.com/v4/spreadsheets/${INTERVIEW_PREP_SPREADSHEET_ID}/values:batchGet?${qs}&majorDimension=ROWS`);
  const [companiesV, casesV, phasesV, proposalsV] = (json.valueRanges || []).map((r: any) => r.values as string[][]);
  const companies = new Map(toRows(companiesV, ['id', 'name']).map(r => [r.id, r.name]));
  const cases = toRows(casesV, ['company_id', 'phase', 'result', 'created_at'])
    .map(r => ({ companyId: r.company_id, phase: r.phase, result: r.result, createdAt: r.created_at }));
  const knowledge = toRows(phasesV, ['company_id', 'phase', 'position', 'content', 'updated_at'])
    .filter(r => r.content.trim() && r.content.trim() !== 'なし')
    .map(r => ({ companyId: r.company_id, phase: r.phase, position: r.position, content: r.content, updatedAt: r.updated_at }));
  const proposals = toRows(proposalsV, ['id', 'company_id', 'company_hint', 'candidate', 'phase', 'result', 'reason', 'confidence', 'status', 'created_at'])
    .filter(r => r.id && r.status !== 'dismissed')
    .map(r => ({
      id: r.id, companyId: r.company_id, companyHint: r.company_hint, candidate: r.candidate, phase: r.phase,
      result: r.result, reason: r.reason, confidence: r.confidence, status: r.status, createdAt: r.created_at,
    }));
  return { companies, cases, knowledge, proposals };
};

// 「面接対策_企業名_ポジション_フェーズ_YYYY-MM-DD.docx」を後ろから分解する（ポジションに
// 「_」が入っていても崩れないように）。
export const parseDocName = (name: string): Omit<PrepDoc, 'id' | 'name' | 'url' | 'createdTime'> | null => {
  const m = name.replace(/\.docx$/i, '').match(/^面接対策_(.+)$/);
  if (!m) return null;
  const parts = m[1].split('_');
  if (parts.length < 3) return null;
  const date = /^\d{4}-\d{2}-\d{2}$/.test(parts[parts.length - 1]) ? parts.pop()! : '';
  const phase = parts.pop()!;
  const company = parts.shift()!;
  return { company, position: parts.join('_'), phase, date };
};

const loadDocs = async (): Promise<PrepDoc[]> => {
  const q = "name contains '面接対策_' and 'me' in owners and trashed = false";
  const url = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&orderBy=createdTime desc&pageSize=300&fields=files(id,name,webViewLink,createdTime)`;
  const json = await googleFetch(url);
  return (json.files || []).flatMap((f: any) => {
    const parsed = parseDocName(f.name);
    return parsed ? [{ id: f.id, name: f.name, url: f.webViewLink, createdTime: f.createdTime, ...parsed }] : [];
  });
};

let cache: { promise: Promise<InterviewPrepData>; at: number } | null = null;
const CACHE_MS = 10 * 60 * 1000;

/** 10分キャッシュ付きで読み込む。force=trueで再読込。 */
export const loadInterviewPrepData = (force = false): Promise<InterviewPrepData> => {
  if (!force && cache && Date.now() - cache.at < CACHE_MS) return cache.promise;
  const promise = (async () => {
    // シートが読めなくても（権限なし等）資料の検索は自分のドライブなので独立して試す。
    const [sheets, docs] = await Promise.allSettled([loadSheets(), loadDocs()]);
    if (sheets.status === 'rejected' && docs.status === 'rejected') throw sheets.reason;
    const s = sheets.status === 'fulfilled' ? sheets.value : { companies: new Map(), cases: [], knowledge: [], proposals: [] };
    const sheetsError = sheets.status === 'rejected'
      ? (sheets.reason instanceof InterviewPrepAccessError ? sheets.reason.message : '面接対策ジェネレーターのデータを読み込めませんでした。')
      : undefined;
    return { ...s, docs: docs.status === 'fulfilled' ? docs.value : [], sheetsError, loadedAt: Date.now() };
  })();
  cache = { promise, at: Date.now() };
  promise.catch(() => { if (cache?.promise === promise) cache = null; });
  return promise;
};

// --- 集計・照合 ---

/**
 * 選考に保存された企業ID（企業名をマスタから選んだ場合）が有効ならそれを使う。企業名を後から
 * 書き換えてマスタの名前と合わなくなっていたら、古いIDとみなして無視する。
 */
const validGasCompanyId = (data: InterviewPrepData, companyName: string, gasCompanyId?: string): string | null => {
  if (!gasCompanyId) return null;
  const masterName = data.companies.get(gasCompanyId);
  return masterName && companyNamesMatch(masterName, companyName) ? gasCompanyId : null;
};

// 企業IDがあればそれだけを使う。無ければ名前で探し、完全一致する企業があればそれだけ、
// なければゆるい一致にする（「オリックス」に「オリックス銀行」の実績が混ざらないように）。
const companyIdsMatching = (data: InterviewPrepData, companyName: string, gasCompanyId?: string): Set<string> => {
  const byId = validGasCompanyId(data, companyName, gasCompanyId);
  if (byId) return new Set([byId]);
  const exact = new Set<string>();
  const loose = new Set<string>();
  data.companies.forEach((name, id) => {
    if (companyNamesEqual(name, companyName)) exact.add(id);
    else if (companyNamesMatch(name, companyName)) loose.add(id);
  });
  return exact.size > 0 ? exact : loose;
};

export interface PhaseStats { passed: number; rejected: number; other: number; rate: number | null }

/** 企業×フェーズの過去実績（cases）。通過率 = (通過+内定) / (通過+内定+不採用)。 */
export const getPhaseStats = (data: InterviewPrepData, companyName: string, stage: string, gasCompanyId?: string): PhaseStats | null => {
  const phase = gasPhaseForStage(stage);
  if (!phase) return null;
  const ids = companyIdsMatching(data, companyName, gasCompanyId);
  if (ids.size === 0) return null;
  let passed = 0, rejected = 0, other = 0;
  for (const c of data.cases) {
    if (!ids.has(c.companyId) || c.phase !== phase) continue;
    if (c.result === '通過' || c.result === '内定') passed++;
    else if (c.result === '不採用') rejected++;
    else other++;
  }
  const decided = passed + rejected;
  return { passed, rejected, other, rate: decided > 0 ? passed / decided : null };
};

export const getKnowledge = (data: InterviewPrepData, companyName: string, stage: string, gasCompanyId?: string): PrepKnowledge[] => {
  const phase = gasPhaseForStage(stage);
  if (!phase) return [];
  const ids = companyIdsMatching(data, companyName, gasCompanyId);
  return data.knowledge
    .filter(k => ids.has(k.companyId) && k.phase === phase)
    .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
};

/** この選考に該当しそうな資料（同じ企業の資料。フェーズ一致を先に）。 */
export const findDocsForApplication = (data: InterviewPrepData, companyName: string, stage: string): PrepDoc[] => {
  const phase = gasPhaseForStage(stage);
  const exact = data.docs.filter(d => companyNamesEqual(d.company, companyName));
  return (exact.length > 0 ? exact : data.docs.filter(d => companyNamesMatch(d.company, companyName)))
    .sort((a, b) => Number(b.phase === phase) - Number(a.phase === phase) || b.createdTime.localeCompare(a.createdTime));
};

export interface ResultSuggestion {
  proposal: MailProposal;
  candidateId: string;
  applicationId: string;
  candidateName: string;
  companyName: string;
  currentStage: string;
  suggestedStage: string;
}

const FORWARD = ['打診', '書類選考', '適性検査', 'カジュアル面談', '1次面接', '2次面接', '最終面接', 'オファー面談', '内定', '内定承諾'];

const suggestedStageFor = (currentStage: string, result: string): string | null => {
  if (result === '不採用') return 'お見送り';
  if (result === '辞退') return '選考辞退';
  if (result === '内定') return '内定';
  if (result === '通過') {
    const i = FORWARD.indexOf(currentStage);
    // 書類選考の次が適性検査とは限らないので、通過の次は「1次面接」を既定にする。
    if (currentStage === '書類選考' || currentStage === '適性検査') return '1次面接';
    return i >= 0 && i + 1 < FORWARD.length ? FORWARD[i + 1] : null;
  }
  return null; // 保留など
};

/**
 * メールから読み取られた選考結果（mail_proposals）のうち、kpi-mgr側の選考がまだその
 * フェーズで止まっているものを「反映の提案」にする。照合は候補者名（空白無視）＋企業名。
 * そのステージに入った日より前のメールは古い結果なので除外する。
 */
export const findResultSuggestions = (
  data: InterviewPrepData,
  candidates: { id: string; name: string; isHidden?: boolean; applications: { id: string; companyName: string; stage: string; gasCompanyId?: string; stageHistory?: { stage: string; date?: string }[] }[] }[],
): ResultSuggestion[] => {
  const byName = new Map<string, typeof candidates>();
  for (const c of candidates) {
    if (c.isHidden) continue;
    const key = normalizePersonName(c.name);
    if (!key) continue;
    byName.set(key, [...(byName.get(key) || []), c]);
  }
  const suggestions: ResultSuggestion[] = [];
  const seen = new Set<string>();
  // 新しいメールから見る（同じ選考に複数届いていたら最新の1件を採用）。
  const newestFirst = [...data.proposals].sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  for (const p of newestFirst) {
    const stage = kpiStageForGasPhase(p.phase);
    if (!stage) continue;
    const matchedCandidates = byName.get(normalizePersonName(p.candidate));
    if (!matchedCandidates) continue;
    const proposalCompany = data.companies.get(p.companyId) || p.companyHint;
    for (const c of matchedCandidates) {
      for (const app of c.applications) {
        if (app.stage !== stage) continue;
        const appCompanyId = validGasCompanyId(data, app.companyName, app.gasCompanyId);
        // 企業IDが一致すれば確実に同じ会社。IDが違っても（マスタの重複登録など）名前が合えば同じとみなす。
        const sameCompany = (!!appCompanyId && appCompanyId === p.companyId) || companyNamesMatch(app.companyName, proposalCompany);
        if (!sameCompany) continue;
        const enteredAt = [...(app.stageHistory || [])].reverse().find(h => h.stage === stage)?.date;
        if (enteredAt && p.createdAt && p.createdAt.slice(0, 10) < enteredAt) continue;
        const suggestedStage = suggestedStageFor(app.stage, p.result);
        if (!suggestedStage) continue;
        const key = `${app.id}:${p.phase}`;
        if (seen.has(key)) continue;
        seen.add(key);
        suggestions.push({
          proposal: p, candidateId: c.id, applicationId: app.id, candidateName: c.name,
          companyName: app.companyName, currentStage: app.stage, suggestedStage,
        });
      }
    }
  }
  return suggestions;
};
