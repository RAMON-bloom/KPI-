// 求人票ジェネレーター（GAS）がドライブに保存した求人票を、索引スプレッドシートから「読むだけ」で使う。
//
// 求人票ジェネレーターは解析した求人票を、共有ドライブ「3_BCA事業部」の
// 1_メインフォルダ/3_企業側/zzz_求人票ジェネレータストレージ/<企業>/<ポジション>/ に
// PDFとGoogleドキュメントで保存し、同じフォルダ直下の「求人票_索引」（シート jobs）に1行記録する。
// KPI管理くんはドライブのファイル名を探しに行かず、この索引だけを読む（2026-10-16 合意の設計）。
// 書き込みは一切しない。列は見出し名で引くので、列の追加・並び替えには耐える。
//
// 索引の position は求人票の解析結果（求人タイトル）なので、KPIの「選考ポジション」
// （面接対策ジェネレーターの企業ナレッジ・jds の名前）とは一致しないことがある。突き合わせはゆるめにする。

import { googleFetch, toRows, companyNamesEqual, companyNamesMatch, InterviewPrepAccessError } from './interviewPrepData';

export const JD_INDEX_SPREADSHEET_ID = '15hIOIOYYfWFc8Qlcq3gh-44xls8hzB3vS1cgWNxvSTg';
const JD_INDEX_SHEET = 'jobs';

const INDEX_COLUMNS = [
  'jd_id', 'kind', 'company', 'position', 'company_folder_id', 'folder_id', 'pdf_file_id', 'doc_file_id',
  'version', 'content_hash', 'data_json', 'created_at', 'created_by', 'updated_at', 'updated_by',
  'doc_saved_modified', 'status',
];

/** 求人票の構造化データ（求人票ジェネレーターの JD_DATA_KEYS）。 */
export type JdData = Record<string, string>;

export interface JdIndexEntry {
  jdId: string;
  kind: string; // master（正本）／ arrange（アレンジ版、今後）
  company: string;
  position: string;
  companyFolderId: string;
  folderId: string;
  pdfFileId: string;
  docFileId: string;
  version: number;
  data: JdData;
  createdAt: string;
  createdBy: string;
  updatedAt: string;
  updatedBy: string;
  docSavedModified: string;
  status: string;
}

export interface JdIndexData {
  entries: JdIndexEntry[];
  loadedAt: number;
}

const parseData = (json: string): JdData => {
  try {
    const v = JSON.parse(json || '{}');
    if (!v || typeof v !== 'object') return {};
    return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, x == null ? '' : String(x)]));
  } catch {
    return {};
  }
};

const loadIndex = async (): Promise<JdIndexData> => {
  let json: any;
  try {
    json = await googleFetch(`https://sheets.googleapis.com/v4/spreadsheets/${JD_INDEX_SPREADSHEET_ID}/values/${encodeURIComponent(`${JD_INDEX_SHEET}!A:Z`)}`);
  } catch (err) {
    // googleFetch の文言は面接対策ジェネレーター向けなので、索引向けに言い換える
    if (err instanceof InterviewPrepAccessError) {
      throw new InterviewPrepAccessError(/Sheets API/.test(err.message)
        ? 'Google Sheets APIが有効になっていないため、求人票の索引を読めません（管理者による設定が必要です）。'
        : '求人票の索引（求人票ジェネレーターの保存先）を閲覧する権限がありません。');
    }
    throw err;
  }
  const entries = toRows(json.values, INDEX_COLUMNS)
    .filter(r => r.jd_id && r.company)
    .map(r => ({
      jdId: r.jd_id, kind: r.kind || 'master', company: r.company, position: r.position,
      companyFolderId: r.company_folder_id, folderId: r.folder_id, pdfFileId: r.pdf_file_id, docFileId: r.doc_file_id,
      version: Number(r.version) || 1, data: parseData(r.data_json),
      createdAt: r.created_at, createdBy: r.created_by, updatedAt: r.updated_at, updatedBy: r.updated_by,
      docSavedModified: r.doc_saved_modified, status: r.status || 'active',
    }));
  return { entries, loadedAt: Date.now() };
};

let cache: { promise: Promise<JdIndexData>; at: number } | null = null;
const CACHE_MS = 10 * 60 * 1000;

/** 10分キャッシュ付きで読み込む。force=trueで再読込。 */
export const loadJdIndex = (force = false): Promise<JdIndexData> => {
  if (!force && cache && Date.now() - cache.at < CACHE_MS) return cache.promise;
  const promise = loadIndex();
  cache = { promise, at: Date.now() };
  promise.catch(() => { if (cache?.promise === promise) cache = null; });
  return promise;
};

// --- 照合 ---

const normalizePosition = (s: string): string =>
  (s || '').normalize('NFKC').replace(/[\s・･\/／\-‐－_＿()（）【】\[\]「」]/g, '').toLowerCase();

/** 企業名（表記ゆれを吸収）に合う正本（kind=master, status=active）。新しい順。完全一致が無ければゆるい一致。 */
export const listCompanyMasters = (data: JdIndexData, companyName: string): JdIndexEntry[] => {
  const name = companyName.trim();
  if (!name) return [];
  const masters = data.entries.filter(e => e.kind === 'master' && e.status === 'active');
  const exact = masters.filter(e => companyNamesEqual(e.company, name));
  return (exact.length > 0 ? exact : masters.filter(e => companyNamesMatch(e.company, name)))
    .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
};

/** 索引の正本のポジション名と同じか（空白・記号・全角半角の違いは無視）。 */
export const samePosition = (a: string, b: string): boolean => {
  const x = normalizePosition(a);
  return !!x && x === normalizePosition(b);
};

/** 求人票ジェネレーター（本番）。company / position のURLパラメータはジェネレーター側が今後対応する（今は無視される）。 */
export const JD_GENERATOR_URL = 'https://script.google.com/a/macros/bloom-firm.com/s/AKfycbzZ_D7gFoUTNZVASr30pCxoi_yFWxPSh8rUBUwsH5i6SehkLXF9hSxePuV6J_ELqbI4fw/exec';
export const buildJdGeneratorUrl = (companyName: string, position: string): string => {
  const params = new URLSearchParams();
  if (companyName.trim()) params.set('company', companyName.trim());
  if (position.trim()) params.set('position', position.trim());
  const qs = params.toString();
  return qs ? `${JD_GENERATOR_URL}?${qs}` : JD_GENERATOR_URL;
};

export interface MasterJdMatch {
  entry: JdIndexEntry;
  /** position: ポジションが一致 / only: その企業の正本が1つだけ（ポジション未指定または不一致）/ none: 一致なし（同じ企業の正本の一覧として出す） */
  how: 'position' | 'only' | 'none';
}

/**
 * 企業名（表記ゆれを吸収）とポジションに合う正本（kind=master, status=active）を探す。
 * 戻り値 best は求人票の初期値に使ってよいもの（ポジション一致、またはポジションを指定しておらず
 * その企業の正本が1つだけ）。others は同じ企業のほかの正本（リンクだけ出す）。
 */
export const findMasterJds = (
  data: JdIndexData, companyName: string, position: string,
): { best: MasterJdMatch | null; others: JdIndexEntry[] } => {
  const sameCompany = listCompanyMasters(data, companyName);
  if (sameCompany.length === 0) return { best: null, others: [] };

  const pos = normalizePosition(position);
  let best: MasterJdMatch | null = null;
  if (pos) {
    const eq = sameCompany.find(e => normalizePosition(e.position) === pos);
    const partial = eq ? null : sameCompany.filter(e => {
      const p = normalizePosition(e.position);
      return p.length >= 2 && (p.includes(pos) || pos.includes(p));
    });
    if (eq) best = { entry: eq, how: 'position' };
    else if (partial && partial.length === 1) best = { entry: partial[0], how: 'position' };
  } else if (sameCompany.length === 1) {
    best = { entry: sameCompany[0], how: 'only' };
  }
  return { best, others: sameCompany.filter(e => e !== best?.entry) };
};

export const jdPdfUrl = (e: JdIndexEntry) => e.pdfFileId ? `https://drive.google.com/file/d/${e.pdfFileId}/view` : '';
export const jdDocUrl = (e: JdIndexEntry) => e.docFileId ? `https://docs.google.com/document/d/${e.docFileId}/edit` : '';
export const jdFolderUrl = (e: JdIndexEntry) => e.folderId ? `https://drive.google.com/drive/folders/${e.folderId}` : '';

// 求人票ジェネレーターのドキュメント（buildDocImportHtml）と同じ並び・見出し
const JD_TEXT_SECTIONS: { title: string; rows: [string, string][] }[] = [
  { title: '基本情報', rows: [['会社名', 'company'], ['職種 / 募集ポジション', 'position'], ['雇用形態', 'employment'], ['契約期間', 'contract']] },
  { title: '給与・勤務条件', rows: [['給与', 'salary'], ['勤務地', 'location'], ['勤務時間', 'workhours'], ['休日・休暇', 'holiday']] },
  { title: '福利厚生・待遇', rows: [['福利厚生', 'benefits'], ['加入保険', 'insurance']] },
  { title: '業務内容', rows: [['主な業務内容', 'job_desc'], ['会社概要', 'company_overview'], ['募集背景', 'recruit_bg']] },
  { title: '応募要件', rows: [['必須要件', 'req_must'], ['歓迎要件', 'req_welcome'], ['求める人物像', 'req_person']] },
  { title: '会社情報', rows: [['代表者', 'ceo'], ['本社所在地', 'hq'], ['アクセス', 'access'], ['設立年月日', 'founded'], ['資本金', 'capital'], ['従業員数', 'employees'], ['事業内容', 'biz'], ['直近決算', 'financial']] },
];

/** 構造化データから、面接対策ジェネレーターに渡す求人票の本文を組み立てる。 */
export const buildJdText = (data: JdData): string =>
  JD_TEXT_SECTIONS.map(sec => {
    const lines = sec.rows
      .map(([label, key]) => [label, (data[key] || '').trim()] as const)
      .filter(([, v]) => v)
      .map(([label, v]) => (v.includes('\n') ? `${label}:\n${v}` : `${label}: ${v}`));
    return lines.length ? `【${sec.title}】\n${lines.join('\n')}` : '';
  }).filter(Boolean).join('\n\n');
