// 各チームがこれまで独自フォーマットで管理していたKPI実績スプレッドシート（CSV書き出し）を
// アプリのKPI実績（日付ごとのKpiKey別数値）に変換するための一度きりの移行用パーサー。
// どの列がどのKPI項目に対応するかはAI（呼び出し側でGemini等に列名・サンプル行を渡して判定）に
// 任せ、このモジュール自身は「対応付けが決まった後の、実際の数値集計」だけを担当する —
// 数値の合計・日付の正規化はAIの読み間違いに左右されない決定的なロジックにするため。

import { parseCsvRows } from './mediaCsvImport';

export interface SpreadsheetGrid {
  header: string[];
  rows: string[][];
}

/** Reads a File as text, auto-detecting UTF-8 vs Shift-JIS (same heuristic as mediaCsvImport). */
export async function decodeSpreadsheetCsvFile(file: File): Promise<string> {
  const buffer = await file.arrayBuffer();
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(buffer);
  } catch {
    text = new TextDecoder('shift-jis').decode(buffer);
  }
  return text.charCodeAt(0) === 0xFEFF ? text.slice(1) : text;
}

/** Splits a decoded CSV string into a header row + data rows (no interpretation of columns yet). */
export function parseSpreadsheetGrid(text: string): SpreadsheetGrid {
  const rows = parseCsvRows(text);
  const [header, ...dataRows] = rows;
  return { header: header || [], rows: dataRows };
}

const FULL_WIDTH_DIGITS: Record<string, string> = {
  '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
  '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
};

function toHalfWidthDigits(value: string): string {
  return value.replace(/[０-９]/g, ch => FULL_WIDTH_DIGITS[ch] ?? ch);
}

/** Excel serial date (days since 1899-12-30, matching Excel's leap-year-bug epoch). */
function fromExcelSerial(serial: number): string | null {
  if (!Number.isFinite(serial) || serial < 20000 || serial > 80000) return null; // ~1954–~2119
  const ms = Date.UTC(1899, 11, 30) + serial * 86400000;
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return null;
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
}

/**
 * Parses a single date cell into yyyy-MM-dd, tolerating the various formats old team
 * spreadsheets tend to use: yyyy/mm/dd, yyyy-mm-dd, yyyy年mm月dd日, mm/dd（年なし — fallbackYear
 * を補う）、Excelのシリアル値がそのままテキスト化されてしまっているケース。判定できなければ
 * nullを返す（呼び出し側でスキップしてカウントする）。
 */
export function parseFlexibleDate(raw: string, fallbackYear: number): string | null {
  const value = toHalfWidthDigits((raw || '').trim());
  if (!value) return null;

  let match = value.match(/^(\d{4})[/\-年](\d{1,2})[/\-月](\d{1,2})日?$/);
  if (match) {
    const [, y, mo, d] = match;
    return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }

  match = value.match(/^(\d{1,2})[/\-月](\d{1,2})日?$/);
  if (match) {
    const [, mo, d] = match;
    return `${fallbackYear}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }

  if (/^\d{4,6}$/.test(value)) {
    const serial = fromExcelSerial(Number(value));
    if (serial) return serial;
  }

  // Last resort: let the JS Date parser take a shot (covers e.g. "2026-08-01T00:00:00" ISO
  // timestamps some exports include), but only accept plausible years to avoid silently
  // misreading garbage text as some far-future/past date.
  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime()) && parsed.getFullYear() >= 2000 && parsed.getFullYear() <= 2100) {
    return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, '0')}-${String(parsed.getDate()).padStart(2, '0')}`;
  }

  return null;
}

function toNumber(value: string | undefined): number {
  const cleaned = toHalfWidthDigits((value ?? '').trim()).replace(/,/g, '');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

export interface KpiImportByTargetResult {
  // target(取り込み先を表す任意のキー — 自分の場合は自分のメールアドレス、チームメンバー分の
  // 場合はそのメンバーのメールアドレス) -> 日付 -> KPIキー -> 合計値。
  countsByTarget: Record<string, Record<string, Record<string, number>>>;
  rowsParsed: number;
  rowsSkipped: number; // 日付を判定できなかった行
  rowsUnassigned: number; // 担当者を対象メンバーに対応付けられなかった行（チーム取込み時のみ発生）
}

/**
 * Given the confirmed column→KPI項目 mapping (dateColumnIndex + one KPI key per other column,
 * empty string meaning "この列は取り込まない") and a per-row target resolver, sums every row's
 * values into per-target/per-date totals. Multiple source columns mapped to the same KPI key are
 * summed together (e.g. a sheet with separate 新規スカウト/継続スカウト columns that should both
 * count toward scoutsSent), and multiple rows sharing the same target+date are summed rather than
 * overwritten. `resolveTarget` returning null skips the row entirely (counted as
 * rowsUnassigned) — used when a row's 担当者 value couldn't be matched to any known member.
 * Self-only imports pass a resolver that always returns the same (自分の)key, so this single
 * function covers both the "自分のみ" and "チームメンバー分も" cases.
 */
export function computeKpiCountsByTarget(
  grid: SpreadsheetGrid,
  dateColumnIndex: number,
  kpiKeyByColumnIndex: Map<number, string>,
  fallbackYear: number,
  resolveTarget: (row: string[]) => string | null
): KpiImportByTargetResult {
  const countsByTarget: Record<string, Record<string, Record<string, number>>> = {};
  let rowsParsed = 0;
  let rowsSkipped = 0;
  let rowsUnassigned = 0;

  grid.rows.forEach(row => {
    const target = resolveTarget(row);
    if (!target) { rowsUnassigned += 1; return; }
    const dateISO = parseFlexibleDate(row[dateColumnIndex] || '', fallbackYear);
    if (!dateISO) { rowsSkipped += 1; return; }
    const targetBucket = countsByTarget[target] || (countsByTarget[target] = {});
    const dateBucket = targetBucket[dateISO] || (targetBucket[dateISO] = {});
    kpiKeyByColumnIndex.forEach((kpiKey, colIdx) => {
      if (!kpiKey) return;
      dateBucket[kpiKey] = (dateBucket[kpiKey] || 0) + toNumber(row[colIdx]);
    });
    rowsParsed += 1;
  });

  return { countsByTarget, rowsParsed, rowsSkipped, rowsUnassigned };
}
