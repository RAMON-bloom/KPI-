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

function parseByColumns(
  text: string,
  dateColumn: string,
  sentColumn: string,
  repliesColumn: string
): ScoutCsvParseResult {
  const rows = parseCsvRows(text);
  const header = rows[0] || [];
  const dateIdx = header.indexOf(dateColumn);
  const sentIdx = header.indexOf(sentColumn);
  const repliesIdx = header.indexOf(repliesColumn);
  if (dateIdx === -1 || sentIdx === -1 || repliesIdx === -1) {
    throw new Error(`CSVの列（${dateColumn}/${sentColumn}/${repliesColumn}）が見つかりませんでした。フォーマットが変わっている可能性があります。`);
  }

  const countsByDate: Record<string, ScoutCsvDayCounts> = {};
  let rowsParsed = 0;
  let rowsSkipped = 0;
  rows.slice(1).forEach(row => {
    const dateISO = toDateISOFromSlash(row[dateIdx] || '');
    if (!dateISO) { rowsSkipped += 1; return; }
    countsByDate[dateISO] = {
      scoutsSent: toNumber(row[sentIdx]),
      scoutReplies: toNumber(row[repliesIdx]),
    };
    rowsParsed += 1;
  });

  return { countsByDate, rowsParsed, rowsSkipped };
}

/** BIZ's export: 日付 column + スカウト合計送信数/スカウト合計返信数 (already the sum of 通常/プラチナ/一時未ログイン). */
export function parseBizScoutCsv(text: string): ScoutCsvParseResult {
  return parseByColumns(text, '日付', 'スカウト合計送信数', 'スカウト合計返信数');
}

/**
 * Doda's export: despite the header name, 抽出対象期間 holds one specific date per row (集計日
 * is the constant report-generation date, not a KPI date, and is ignored). Doda's own totals
 * column is 有効返信数(ALL) — there's no separate raw-reply count distinct from "effective"
 * in this export, so that's used as this media's 返信数.
 */
export function parseDodaScoutCsv(text: string): ScoutCsvParseResult {
  return parseByColumns(text, '抽出対象期間', '送信数(ALL)', '有効返信数(ALL)');
}

export function parseScoutCsv(mediaId: ScoutCsvMediaId, text: string): ScoutCsvParseResult {
  return mediaId === 'biz' ? parseBizScoutCsv(text) : parseDodaScoutCsv(text);
}
