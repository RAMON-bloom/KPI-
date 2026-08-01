// 各チームがこれまで独自フォーマットで管理していた候補者パイプラインスプレッドシート（CSV書き出し）を
// アプリの候補者パイプライン（候補者＋選考企業ごとの応募）に変換するための一度きりの移行用パーサー。
// スプレッドシートの1行が「候補者1人」を表す形式（選考企業は高々1社分の列だけ持つ）にも、「候補者×
// 選考企業の組み合わせ1件」を表す形式（同じ候補者が複数行に分かれ、行ごとに別の選考企業を持つ）にも
// 同じロジックで対応する — 氏名でグルーピングし、同じ候補者の行が複数あれば選考企業ごとに別々の
// 応募として積み上げるだけなので、どちらの形式かをユーザーに事前に選んでもらう必要がない。
// どの列がどの項目に対応するか、各列の生の値（選考ステータス・経由媒体など）がどの選択肢に対応するかは
// AI（呼び出し側でGemini等に列名・サンプル値を渡して判定）に任せ、このモジュール自身は「対応付けが
// 決まった後の、実際の候補者データ組み立て」だけを担当する（KPI実績取込み: services/
// spreadsheetKpiImport.tsと同じ役割分担）。

import { parseFlexibleDate, type SpreadsheetGrid } from './spreadsheetKpiImport';

export type CandidateImportFieldKey =
  | 'name' | 'currentCompany' | 'currentSalary' | 'salary' | 'education' | 'jobType'
  | 'age' | 'phoneNumber' | 'email' | 'otherCompanyStatus' | 'desiredJoinTiming' | 'memo'
  | 'source' | 'companyName' | 'stage' | 'nextAction' | 'scheduledDate' | 'expectedDecisionDate';

export interface CandidateImportFieldSpec {
  key: CandidateImportFieldKey;
  label: string;
  required?: boolean;
}

// 選考企業（companyName）・選考ステータス（stage）・次のアクション・次回予定日・決定見込み日は
// 応募（CompanyApplication）側のフィールド、それ以外は候補者本体のフィールド。
export const CANDIDATE_IMPORT_FIELD_CATALOG: CandidateImportFieldSpec[] = [
  { key: 'name', label: '氏名', required: true },
  { key: 'currentCompany', label: '現職企業名' },
  { key: 'currentSalary', label: '現年収（万円）' },
  { key: 'salary', label: '希望年収（万円）' },
  { key: 'education', label: '学歴' },
  { key: 'jobType', label: '職種' },
  { key: 'age', label: '年齢' },
  { key: 'phoneNumber', label: '電話番号' },
  { key: 'email', label: 'メールアドレス' },
  { key: 'otherCompanyStatus', label: '他社選考状況' },
  { key: 'desiredJoinTiming', label: '入社希望時期' },
  { key: 'memo', label: '候補者メモ' },
  { key: 'source', label: '経由媒体' },
  { key: 'companyName', label: '選考企業名' },
  { key: 'stage', label: '選考ステータス' },
  { key: 'nextAction', label: '次のアクション' },
  { key: 'scheduledDate', label: '次回選考予定日' },
  { key: 'expectedDecisionDate', label: '決定見込み日' },
];

const CANDIDATE_LEVEL_FIELD_KEYS = new Set<CandidateImportFieldKey>([
  'currentCompany', 'currentSalary', 'salary', 'education', 'jobType', 'age',
  'phoneNumber', 'email', 'otherCompanyStatus', 'desiredJoinTiming', 'memo',
]);
const NUMERIC_FIELD_KEYS = new Set<CandidateImportFieldKey>(['currentSalary', 'salary', 'age']);

const FULL_WIDTH_DIGITS: Record<string, string> = {
  '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
  '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
};

function toHalfWidthDigits(value: string): string {
  return value.replace(/[０-９]/g, ch => FULL_WIDTH_DIGITS[ch] ?? ch);
}

function toNumber(value: string): number | undefined {
  const cleaned = toHalfWidthDigits(value.trim()).replace(/[,万円]/g, '');
  const n = Number(cleaned);
  return Number.isFinite(n) && n !== 0 ? n : undefined;
}

export interface ImportedApplicationDraft {
  companyName: string;
  stage: string;
  nextAction?: string;
  scheduledDate?: string; // yyyy-mm-dd
  expectedDecisionDate?: string; // yyyy-mm-dd
}

export interface ImportedCandidateDraft {
  name: string;
  currentCompany?: string;
  currentSalary?: number;
  salary?: number;
  education?: string;
  jobType?: string;
  age?: number;
  phoneNumber?: string;
  email?: string;
  otherCompanyStatus?: string;
  desiredJoinTiming?: string;
  memo?: string;
  source?: string; // 解決済み: MediaEntry.id または 'Other'
  applications: ImportedApplicationDraft[];
}

export interface BuildCandidateDraftsResult {
  // target（自分の場合は自分のメールアドレス、チームメンバー分の場合はそのメンバーのメール
  // アドレス）ごとの候補者ドラフト一覧。同じtarget内で氏名が一致する行は1人の候補者にまとめられ、
  // 選考企業列を持つ行は（重複する企業名+ステータスの組み合わせを除いて）それぞれ1件の応募として
  // 積み上げられる。
  draftsByTarget: Record<string, ImportedCandidateDraft[]>;
  rowsParsed: number;
  rowsSkippedNoName: number; // 氏名列が空だった行
  rowsUnassigned: number; // 担当者を対象メンバーに対応付けられなかった行（チーム取込み時のみ発生）
}

/**
 * 列→項目の対応付け（fieldByColumnIndex）と、選考ステータス・経由媒体それぞれの生の値→取込み先
 * キーの対応付け（stageValueMap/sourceValueMap、どちらも呼び出し側でAI判定→ユーザー確認済みの
 * もの）が確定した後、実際に全行を読んで候補者ドラフトへ変換する決定的なロジック。同じ氏名の行は
 * target単位でグルーピングし、複数の応募がある候補者は自然と複数件のapplicationsを持つ。
 */
export function buildCandidateDraftsFromGrid(
  grid: SpreadsheetGrid,
  fieldByColumnIndex: Map<number, CandidateImportFieldKey>,
  stageValueMap: Map<string, string>,
  sourceValueMap: Map<string, string>,
  fallbackYear: number,
  resolveTarget: (row: string[]) => string | null
): BuildCandidateDraftsResult {
  const colIdxByField = new Map<CandidateImportFieldKey, number>();
  fieldByColumnIndex.forEach((field, idx) => { if (!colIdxByField.has(field)) colIdxByField.set(field, idx); });

  const nameIdx = colIdxByField.get('name');
  const companyIdx = colIdxByField.get('companyName');
  const stageIdx = colIdxByField.get('stage');
  const sourceIdx = colIdxByField.get('source');
  const nextActionIdx = colIdxByField.get('nextAction');
  const scheduledDateIdx = colIdxByField.get('scheduledDate');
  const expectedDecisionDateIdx = colIdxByField.get('expectedDecisionDate');

  const draftsMapByTarget = new Map<string, Map<string, ImportedCandidateDraft>>();
  let rowsParsed = 0;
  let rowsSkippedNoName = 0;
  let rowsUnassigned = 0;

  grid.rows.forEach(row => {
    const target = resolveTarget(row);
    if (!target) { rowsUnassigned += 1; return; }

    const name = nameIdx !== undefined ? (row[nameIdx] || '').trim() : '';
    if (!name) { rowsSkippedNoName += 1; return; }

    let targetDrafts = draftsMapByTarget.get(target);
    if (!targetDrafts) { targetDrafts = new Map(); draftsMapByTarget.set(target, targetDrafts); }
    let draft = targetDrafts.get(name);
    if (!draft) {
      draft = { name, applications: [] };
      targetDrafts.set(name, draft);
    }

    colIdxByField.forEach((idx, field) => {
      if (!CANDIDATE_LEVEL_FIELD_KEYS.has(field)) return;
      if ((draft as any)[field] !== undefined) return; // 同じ候補者の複数行では最初に見つかった非空値を採用
      const raw = (row[idx] || '').trim();
      if (!raw) return;
      if (NUMERIC_FIELD_KEYS.has(field)) {
        const n = toNumber(raw);
        if (n !== undefined) (draft as any)[field] = n;
      } else {
        (draft as any)[field] = raw;
      }
    });

    if (draft.source === undefined && sourceIdx !== undefined) {
      const raw = (row[sourceIdx] || '').trim();
      if (raw) {
        const mapped = sourceValueMap.get(raw);
        if (mapped) draft.source = mapped;
      }
    }

    if (companyIdx !== undefined) {
      const companyName = (row[companyIdx] || '').trim();
      if (companyName) {
        // stage列が無ければ「打診」を初期ステータスとする。stage列はあるが、その行の生の値が
        // 対応表で「取り込まない」("")、または対応表に無い未知の値の場合はこの行の応募自体を
        // 作らない（候補者本体の情報はそのまま取り込む）。
        const stage = stageIdx === undefined
          ? '打診'
          : (stageValueMap.get((row[stageIdx] || '').trim()) || null);
        if (stage) {
          const nextAction = nextActionIdx !== undefined ? (row[nextActionIdx] || '').trim() : '';
          const scheduledDateRaw = scheduledDateIdx !== undefined ? (row[scheduledDateIdx] || '').trim() : '';
          const expectedDecisionDateRaw = expectedDecisionDateIdx !== undefined ? (row[expectedDecisionDateIdx] || '').trim() : '';
          const scheduledDate = scheduledDateRaw ? parseFlexibleDate(scheduledDateRaw, fallbackYear) ?? undefined : undefined;
          const expectedDecisionDate = expectedDecisionDateRaw ? parseFlexibleDate(expectedDecisionDateRaw, fallbackYear) ?? undefined : undefined;
          const alreadyExists = draft.applications.some(a => a.companyName === companyName && a.stage === stage);
          if (!alreadyExists) {
            draft.applications.push({
              companyName,
              stage,
              ...(nextAction ? { nextAction } : {}),
              ...(scheduledDate ? { scheduledDate } : {}),
              ...(expectedDecisionDate ? { expectedDecisionDate } : {}),
            });
          }
        }
      }
    }

    rowsParsed += 1;
  });

  const draftsByTarget: Record<string, ImportedCandidateDraft[]> = {};
  draftsMapByTarget.forEach((byName, target) => {
    draftsByTarget[target] = Array.from(byName.values());
  });

  return { draftsByTarget, rowsParsed, rowsSkippedNoName, rowsUnassigned };
}

export function extractDistinctColumnValues(grid: SpreadsheetGrid, colIdx: number): string[] {
  const seen = new Set<string>();
  const values: string[] = [];
  grid.rows.forEach(row => {
    const v = (row[colIdx] || '').trim();
    if (v && !seen.has(v)) { seen.add(v); values.push(v); }
  });
  return values;
}
