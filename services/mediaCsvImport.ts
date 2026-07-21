// Parses the daily scout-performance CSV each media platform's own admin screen can export
// (BIZ/ビズリーチ, Doda) into per-date スカウト数/返信数 counts, for direct import instead of
// relying on the Gmail-notification-based heuristic in gmailScout.ts. Both exports are already
// aggregated to one row per day, so no date-bucketing is needed on this side — just column
// lookup by header name (robust to column reordering) and a date-format conversion.

export type ScoutCsvMediaId = 'biz' | 'doda';

export interface ScoutCsvDayCounts {
  scoutsSent: number;
  scoutReplies: number;
}

export interface ScoutCsvParseResult {
  countsByDate: Record<string, ScoutCsvDayCounts>;
  rowsParsed: number;
  rowsSkipped: number;
}

/** Reads a File as text, auto-detecting UTF-8 vs Shift-JIS (BIZ exports UTF-8 w/ BOM; Doda exports Shift-JIS). */
export async function decodeCsvFile(file: File): Promise<string> {
  const buffer = await file.arrayBuffer();
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(buffer);
  } catch {
    text = new TextDecoder('shift-jis').decode(buffer);
  }
  // BIZ's export is UTF-8 with a leading BOM (U+FEFF) — TextDecoder doesn't strip it, and left
  // in place it silently breaks header.indexOf('日付') (the BOM attaches to the first column
  // name, e.g. "﻿日付" !== "日付").
  return text.charCodeAt(0) === 0xFEFF ? text.slice(1) : text;
}

// Minimal RFC4180-style parser: handles double-quoted fields (with "" as an escaped quote) and
// both CRLF/LF line endings. Sufficient for these flat, single-line-per-row exports.
function parseCsvRows(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  const len = text.length;
  while (i < len) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
        inQuotes = false; i += 1; continue;
      }
      field += ch; i += 1; continue;
    }
    if (ch === '"') { inQuotes = true; i += 1; continue; }
    if (ch === ',') { row.push(field); field = ''; i += 1; continue; }
    if (ch === '\r') { i += 1; continue; }
    if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i += 1; continue; }
    field += ch; i += 1;
  }
  if (field.length > 0 || row.length > 0) { row.push(field); rows.push(row); }
  return rows.filter(r => r.some(c => c.trim() !== ''));
}

function toDateISOFromSlash(value: string): string | null {
  const match = value.trim().match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (!match) return null;
  const [, y, mo, d] = match;
  return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`;
}

function toNumber(value: string | undefined): number {
  const n = Number((value ?? '').trim());
  return Number.isFinite(n) ? n : 0;
}

function requireColumnIndex(header: string[], columnName: string): number {
  const idx = header.indexOf(columnName);
  if (idx === -1) {
    throw new Error(`CSVの列「${columnName}」が見つかりませんでした。フォーマットが変わっている可能性があります。`);
  }
  return idx;
}

/** scoutsSent/scoutReplies are each the sum of every column listed, so a media whose export
 * splits sent/reply counts across several categories can include just the ones that should
 * count (e.g. skipping a category that shouldn't be reflected). */
function parseByDateAndSummedColumns(
  text: string,
  dateColumn: string,
  sentColumns: string[],
  repliesColumns: string[]
): ScoutCsvParseResult {
  const rows = parseCsvRows(text);
  const header = rows[0] || [];
  const dateIdx = requireColumnIndex(header, dateColumn);
  const sentIdxs = sentColumns.map(c => requireColumnIndex(header, c));
  const repliesIdxs = repliesColumns.map(c => requireColumnIndex(header, c));

  const countsByDate: Record<string, ScoutCsvDayCounts> = {};
  let rowsParsed = 0;
  let rowsSkipped = 0;
  rows.slice(1).forEach(row => {
    const dateISO = toDateISOFromSlash(row[dateIdx] || '');
    if (!dateISO) { rowsSkipped += 1; return; }
    countsByDate[dateISO] = {
      scoutsSent: sentIdxs.reduce((sum, idx) => sum + toNumber(row[idx]), 0),
      scoutReplies: repliesIdxs.reduce((sum, idx) => sum + toNumber(row[idx]), 0),
    };
    rowsParsed += 1;
  });

  return { countsByDate, rowsParsed, rowsSkipped };
}

/**
 * BIZ's export: 日付 column, summing 通常スカウト + プラチナスカウト送信数/返信数 only —
 * deliberately excludes プラチナスカウト（一時未ログイン専用）送信数/返信数 (a separate
 * category the platform otherwise folds into its own スカウト合計送信数/返信数 columns, which
 * this does NOT use for that reason).
 */
export function parseBizScoutCsv(text: string): ScoutCsvParseResult {
  return parseByDateAndSummedColumns(
    text,
    '日付',
    ['通常スカウト送信数', 'プラチナスカウト送信数'],
    ['通常スカウト返信数', 'プラチナスカウト返信数']
  );
}

/**
 * Doda's export: despite the header name, 抽出対象期間 holds one specific date per row (集計日
 * is the constant report-generation date, not a KPI date, and is ignored). Uses only the
 * (ALL) columns — 送信数(ALL)/有効返信数(ALL) — never the 通常/プラチナ/ダイヤモンド/限定ダイ
 * ヤモンド(内訳) breakdown columns. There's no separate raw-reply count distinct from
 * "effective" in this export, so 有効返信数(ALL) is used as this media's 返信数.
 */
export function parseDodaScoutCsv(text: string): ScoutCsvParseResult {
  return parseByDateAndSummedColumns(text, '抽出対象期間', ['送信数(ALL)'], ['有効返信数(ALL)']);
}

export function parseScoutCsv(mediaId: ScoutCsvMediaId, text: string): ScoutCsvParseResult {
  return mediaId === 'biz' ? parseBizScoutCsv(text) : parseDodaScoutCsv(text);
}
