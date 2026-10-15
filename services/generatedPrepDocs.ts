// 面接対策ジェネレーターで作った資料を、担当者本人のGoogleドライブに保存する（2026-10-16）。
//
// ジェネレーターはKPI管理くんから開いて作った資料を、Word用に整えたHTMLのまま共有スプレッド
// シートの `kpi_generated_docs` に置く（1セル5万文字の上限があるので part に分けて複数行）。
// ジェネレーターは作成者の権限で動くため、そこでドライブに保存すると全員の資料が作成者の
// ドライブに溜まってしまう。そこで本人の権限で動くKPI管理くんが受け取り、本人のドライブの
// 「面接対策資料」フォルダにGoogleドキュメントとして保存してから、シート上の本文を消す。
//
// 同じ人が複数のタブでKPI管理くんを開いていても二重に保存しないよう、保存前に status へ
// 「claimed:<タブID>」を書いて読み直し、自分の印が残っている時だけ保存する。

import { getCurrentSession, refreshTokenSilently } from './googleAuth';
import { googleFetch, INTERVIEW_PREP_SPREADSHEET_ID } from './interviewPrepData';

export const GENERATED_DOCS_SHEET = 'kpi_generated_docs';
export const PREP_DOCS_FOLDER_NAME = '面接対策資料';
const SHEETS_BASE = `https://sheets.googleapis.com/v4/spreadsheets/${INTERVIEW_PREP_SPREADSHEET_ID}`;
const TAB_ID = `tab_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
// 他のタブが受け取りを始めたまま止まった（タブを閉じた等）とみなすまでの時間
const STALE_CLAIM_MS = 5 * 60 * 1000;

/** ジェネレーターから届いた1件の資料（1回の送信分。part をつないだ本文を持つ）。 */
export interface GeneratedDocDelivery {
  docKey: string;
  sentAt: string;
  applicationId: string;
  candidateId: string;
  companyName: string;
  phase: string;
  position: string;
  fileName: string;
  html: string;
  /** シート上の行番号（1始まり）。part の数だけある */
  rows: number[];
}

/** 候補者の選考（CompanyApplication.interviewPrepDocs）に残す、保存済みの資料。 */
export interface SavedPrepDoc {
  key: string;
  fileId: string;
  name: string;
  savedAt: string;
}

export const prepDocOpenUrl = (fileId: string) => `https://docs.google.com/document/d/${fileId}/edit`;
export const prepDocDownloadUrl = (fileId: string) => `https://docs.google.com/document/d/${fileId}/export?format=docx`;
/** 「資料のURL」欄に入っているのがGoogleドキュメントなら、そのファイルID */
export const googleDocIdFromUrl = (url: string): string | null => url.match(/docs\.google\.com\/document\/d\/([\w-]+)/)?.[1] ?? null;

const colLetter = (index: number): string => {
  let n = index + 1, s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
};

type SheetSnapshot = { header: string[]; rows: string[][] };

const readSheet = async (): Promise<SheetSnapshot | null> => {
  try {
    const json = await googleFetch(`${SHEETS_BASE}/values/${encodeURIComponent(`${GENERATED_DOCS_SHEET}!A:Z`)}`);
    const values: string[][] = json.values || [];
    if (values.length === 0) return null;
    return { header: values[0], rows: values.slice(1) };
  } catch (err: any) {
    // シートがまだ無い（ジェネレーターがまだ一度も送っていない）時は 400 になる
    if (/（400）/.test(String(err?.message))) return null;
    throw err;
  }
};

const isClaimable = (status: string) => {
  if (status === 'pending') return true;
  const m = status.match(/^claimed:[^:]+:(\d+)$/);
  return !!m && Date.now() - Number(m[1]) > STALE_CLAIM_MS;
};

/**
 * 受け取り待ちの資料のうち、applicationIds（自分の候補者の選考）に宛てたものを返す。
 * 同じ資料（doc_key）が何度か送られていれば最新の送信分だけを返し、古い送信分は
 * superseded に書き換えて本文を消す。
 */
export const loadPendingGeneratedDocs = async (applicationIds: Set<string>): Promise<GeneratedDocDelivery[]> => {
  const snap = await readSheet();
  if (!snap) return [];
  const col = (name: string) => snap.header.indexOf(name);
  const get = (row: string[], name: string) => (col(name) >= 0 ? row[col(name)] ?? '' : '');
  const groups = new Map<string, { row: string[]; rowNo: number }[]>();
  snap.rows.forEach((row, i) => {
    if (!isClaimable(get(row, 'status')) || !applicationIds.has(get(row, 'kpi_application_id'))) return;
    const id = `${get(row, 'doc_key')}\u0000${get(row, 'sent_at')}`;
    groups.set(id, [...(groups.get(id) || []), { row, rowNo: i + 2 }]);
  });
  const latestByKey = new Map<string, GeneratedDocDelivery>();
  const superseded: number[] = [];
  groups.forEach(parts => {
    const first = parts[0].row;
    const total = Number(get(first, 'parts')) || 1;
    if (parts.length < total) return;   // まだ全部の part が揃っていない（書き込み途中）
    parts.sort((a, b) => Number(get(a.row, 'part')) - Number(get(b.row, 'part')));
    const delivery: GeneratedDocDelivery = {
      docKey: get(first, 'doc_key'), sentAt: get(first, 'sent_at'),
      applicationId: get(first, 'kpi_application_id'), candidateId: get(first, 'kpi_candidate_id'),
      companyName: get(first, 'company_name'), phase: get(first, 'phase'), position: get(first, 'position'),
      fileName: get(first, 'file_name'), html: parts.map(p => get(p.row, 'html')).join(''), rows: parts.map(p => p.rowNo),
    };
    const prev = latestByKey.get(delivery.docKey);
    if (!prev || prev.sentAt < delivery.sentAt) {
      if (prev) superseded.push(...prev.rows);
      latestByKey.set(delivery.docKey, delivery);
    } else {
      superseded.push(...delivery.rows);
    }
  });
  if (superseded.length > 0) await markRows(snap.header, superseded, 'superseded').catch(() => { /* 次回また整理する */ });
  return Array.from(latestByKey.values());
};

const markRows = async (header: string[], rows: number[], status: string, clearHtml = true): Promise<void> => {
  const cStatus = header.indexOf('status'), cHtml = header.indexOf('html'), cSaved = header.indexOf('saved_at');
  const data: { range: string; values: string[][] }[] = [];
  rows.forEach(r => {
    data.push({ range: `${GENERATED_DOCS_SHEET}!${colLetter(cStatus)}${r}`, values: [[status]] });
    if (clearHtml && cHtml >= 0) data.push({ range: `${GENERATED_DOCS_SHEET}!${colLetter(cHtml)}${r}`, values: [['']] });
    if (status === 'saved' && cSaved >= 0) data.push({ range: `${GENERATED_DOCS_SHEET}!${colLetter(cSaved)}${r}`, values: [[new Date().toISOString()]] });
  });
  await googleFetch(`${SHEETS_BASE}/values:batchUpdate`, {
    method: 'POST',
    body: JSON.stringify({ valueInputOption: 'RAW', data }),
  });
};

/** 保存の前に、この資料の行を自分（このタブ）の受け取りとして押さえる。押さえられたら true。 */
const claim = async (delivery: GeneratedDocDelivery): Promise<boolean> => {
  const snap = await readSheet();
  if (!snap) return false;
  const mark = `claimed:${TAB_ID}:${Date.now()}`;
  await markRows(snap.header, delivery.rows, mark, false);
  await new Promise(r => setTimeout(r, 1500));
  const after = await readSheet();
  if (!after) return false;
  const c = after.header.indexOf('status');
  return delivery.rows.every(r => (after.rows[r - 2]?.[c] ?? '').startsWith(`claimed:${TAB_ID}:`));
};

// ---- Googleドライブ ----

const driveFetch = async (url: string, init: RequestInit): Promise<Response> => {
  let session = getCurrentSession();
  if (!session) {
    try { await refreshTokenSilently(); } catch { /* 下で判定 */ }
    session = getCurrentSession();
    if (!session) throw new Error('ログインが必要です。');
  }
  return fetch(url, { ...init, headers: { ...(init.headers || {}), Authorization: `Bearer ${session.accessToken}` } });
};

let folderIdCache: string | null = null;
const ensurePrepDocsFolder = async (): Promise<string> => {
  if (folderIdCache) return folderIdCache;
  const q = `name = '${PREP_DOCS_FOLDER_NAME}' and mimeType = 'application/vnd.google-apps.folder' and 'me' in owners and 'root' in parents and trashed = false`;
  const found = await googleFetch(`https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&fields=files(id)&pageSize=1`);
  if (found.files?.[0]?.id) return (folderIdCache = found.files[0].id as string);
  const created = await googleFetch('https://www.googleapis.com/drive/v3/files?fields=id', {
    method: 'POST',
    body: JSON.stringify({ name: PREP_DOCS_FOLDER_NAME, mimeType: 'application/vnd.google-apps.folder', parents: ['root'] }),
  });
  return (folderIdCache = created.id as string);
};

const docNameOf = (d: GeneratedDocDelivery) => (d.fileName || `面接対策_${d.companyName}_${d.phase}`).replace(/\.docx$/i, '');

/** HTMLをGoogleドキュメントとして新規作成する（Drive がHTMLを変換する）。 */
const createDoc = async (d: GeneratedDocDelivery): Promise<{ id: string; name: string }> => {
  const folderId = await ensurePrepDocsFolder();
  const boundary = `kpi_${Math.random().toString(36).slice(2)}`;
  const metadata = { name: docNameOf(d), mimeType: 'application/vnd.google-apps.document', parents: [folderId] };
  const body =
    `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n` +
    `--${boundary}\r\nContent-Type: text/html; charset=UTF-8\r\n\r\n${d.html}\r\n--${boundary}--`;
  const res = await driveFetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name', {
    method: 'POST', headers: { 'Content-Type': `multipart/related; boundary=${boundary}` }, body,
  });
  if (!res.ok) throw new Error(`ドライブへの保存に失敗しました（${res.status}）`);
  return res.json();
};

/**
 * 既に保存した資料を、送り直された内容（生成後に編集された本文）で上書きする。消されていた・
 * 上書きできなかった時は false（呼び出し側で新しく作り直す）。
 */
const overwriteDoc = async (fileId: string, d: GeneratedDocDelivery): Promise<boolean> => {
  const res = await driveFetch(`https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media&fields=id`, {
    method: 'PATCH', headers: { 'Content-Type': 'text/html; charset=UTF-8' }, body: d.html,
  });
  if (!res.ok) console.warn(`ドライブの資料を上書きできなかったため、新しく作ります（${res.status}）`);
  return res.ok;
};

/**
 * 1件の資料を本人のドライブに保存し、保存した資料の情報を返す（シート上の本文は消す）。
 * 他のタブが受け取り中なら null。existing はその選考に既に保存してある資料の一覧。
 */
export const saveGeneratedDocToDrive = async (d: GeneratedDocDelivery, existing: SavedPrepDoc[]): Promise<SavedPrepDoc | null> => {
  if (!(await claim(d))) return null;
  const header = (await readSheet())?.header || [];
  try {
    const prev = existing.find(x => x.key === d.docKey);
    let saved: SavedPrepDoc;
    if (prev && await overwriteDoc(prev.fileId, d)) {
      saved = { ...prev, savedAt: new Date().toISOString() };
    } else {
      const file = await createDoc(d);
      saved = { key: d.docKey, fileId: file.id, name: file.name, savedAt: new Date().toISOString() };
    }
    // ここで失敗しても保存自体は済んでいるので、資料は返す（行は5分後に再取得されうるが、
    // 次回は existing に同じ key があるので新規作成ではなく上書きになる）
    await markRows(header, d.rows, 'saved').catch(err => console.warn('受け取り済みの記録に失敗しました', err));
    return saved;
  } catch (err) {
    // 押さえた行を戻して、次の確認で再挑戦できるようにする
    await markRows(header, d.rows, 'pending', false).catch(() => { /* 期限切れで戻る */ });
    throw err;
  }
};
