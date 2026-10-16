// 求人票ジェネレーターがドライブに保存した求人票PDF（正本）を検索して複数選び、
// 選んだ順に結合して候補者に渡す1つのPDFにする。索引（services/jdIndex.ts）を読むだけで、
// ドライブには書き込まない。結合はブラウザ内で行う（pdf-lib）。

import { PDFDocument } from 'pdf-lib';
import { downloadFileBytes } from './googleDrive';
import type { JdIndexData, JdIndexEntry } from './jdIndex';

const normalize = (s: string): string =>
  (s || '').normalize('NFKC').toLowerCase().replace(/株式会社|\(株\)|㈱/g, '').replace(/\s+/g, '');

/** 結合に使える求人票（正本・有効・PDFあり）。 */
export const bundlableEntries = (data: JdIndexData): JdIndexEntry[] =>
  data.entries
    .filter(e => e.kind === 'master' && e.status === 'active' && e.pdfFileId)
    .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));

/**
 * 空白区切りのキーワードがすべて含まれる求人票を返す（AND検索）。企業名・ポジションのほか、
 * 勤務地・給与・業務内容など求人票の中身も対象。株式会社の有無・全角半角・大小文字は無視する。
 */
export const searchJdEntries = (entries: JdIndexEntry[], query: string): JdIndexEntry[] => {
  const words = query.split(/[\s　]+/).map(normalize).filter(Boolean);
  if (words.length === 0) return entries;
  return entries.filter(e => {
    const hay = normalize([e.company, e.position, ...Object.values(e.data)].join(' '));
    return words.every(w => hay.includes(w));
  });
};

/** 選んだ求人票のPDFを、並び順のまま1つのPDFに結合する。読めないPDFがあればどれかを示して止める。 */
export const mergeJdPdfs = async (
  entries: JdIndexEntry[],
  onProgress?: (done: number, total: number) => void,
): Promise<Uint8Array> => {
  const out = await PDFDocument.create();
  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    const label = `${e.company}・${e.position || '企業共通'}`;
    let src: PDFDocument;
    try {
      src = await PDFDocument.load(await downloadFileBytes(e.pdfFileId), { ignoreEncryption: true });
    } catch (err) {
      throw new Error(`「${label}」：${err instanceof Error ? err.message : 'PDFを読み込めませんでした。'}`);
    }
    const pages = await out.copyPages(src, src.getPageIndices());
    pages.forEach(p => out.addPage(p));
    onProgress?.(i + 1, entries.length);
  }
  out.setTitle('求人票');
  out.setProducer('KPI管理くん');
  return out.save();
};
