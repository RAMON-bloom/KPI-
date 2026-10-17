// 推薦文の下書き（選考情報 →「推薦文の下書き」）。
//
// その選考の求人票（ドライブの正本 → 面接対策ジェネレーターの jds → 手入力の順）と、候補者の経歴・
// 面談メモを Gemini に渡し、応募先企業に送る推薦文の下書きを作る。担当者が画面で編集・コピーする。
//
// 個人情報の扱いはブラインドレジュメ・類似経歴リサーチと同じ: 氏名・電話番号・メールアドレスは
// Gemini に渡さない（本文中に出てくるものも redactPersonalInfo で伏せる）。生成結果にも同じ処理を
// かける。年齢・学歴・在籍企業名・年収は伏せない。推薦文の中の候補者は「〇〇様」と書かせ、送る前に
// 担当者が名前を入れる。

import { gemini } from './gemini';
import { redactPersonalInfo, collectNotes, basicFacts, type BlindSourceCandidate } from './blindResume';

const MODEL = 'gemini-2.5-flash';

const truncate = (s: string, max: number) => (s.length > max ? `${s.slice(0, max)}…` : s);

export interface RecommendationInput {
  candidate: BlindSourceCandidate;
  companyName: string;
  position: string;
  jdText: string;
}

/** Gemini に渡す匿名の材料（氏名・連絡先なし）。画面の確認用にも使う。 */
export const buildRecommendationMaterial = ({ candidate: c, companyName, position, jdText }: RecommendationInput): string => {
  const summary = redactPersonalInfo((c.summary || '').trim(), c);
  const notes = redactPersonalInfo(collectNotes(c), c);
  return [
    `# 応募先`,
    `企業名: ${companyName.trim() || '（未入力）'}`,
    position.trim() ? `ポジション: ${position.trim()}` : '',
    '',
    '# 求人票',
    jdText.trim() ? truncate(jdText.trim(), 8000) : '（求人票なし。企業名・ポジションから一般的に求められそうな点を踏まえて書く）',
    '',
    '# 候補者（氏名・連絡先は伏せています）',
    ...basicFacts(c),
    summary ? `経歴概要:\n${truncate(summary, 3000)}` : '経歴概要: （未登録）',
    notes ? `面談メモ:\n${truncate(notes, 6000)}` : '',
  ].filter(s => s !== '').join('\n');
};

export const generateRecommendationDraft = async (input: RecommendationInput): Promise<string> => {
  const material = buildRecommendationMaterial(input);
  const instruction = `あなたは人材紹介会社のキャリアアドバイザーです。以下の材料から、応募先企業の採用担当者に送る「推薦文」の下書きを日本語で作成してください。

# 書き方
- 冒頭に、推薦する候補者を一言で表す見出し（1行）を置く
- 続けて「推薦理由」を3〜5点の箇条書きで。求人票の必須要件・歓迎要件・求める人物像と、候補者の経験・実績・強みを具体的に結び付ける（数字・固有の経験があれば使う）
- 最後に「人物面・志向」を2〜4文で（転職理由・企業選びの軸・人柄など、面談メモから読み取れること）
- 懸念がありそうな点は推薦文に書かない。材料に無い経歴・実績は作らない。誇張しない
- 候補者は「〇〇様」と書く。氏名・電話番号・メールアドレスは書かない
- 全体で400〜700字程度。社外に送る文書としての丁寧な文体（です・ます調）
- 見出しや記号はプレーンテキストで（Markdownの # や ** は使わない。箇条書きは「・」）

${material}`;
  const ai = gemini;
  const response = await ai.models.generateContent({ model: MODEL, contents: instruction });
  const text = (response.text || '').trim();
  if (!text) throw new Error('推薦文を作成できませんでした。時間をおいて再度お試しください。');
  return redactPersonalInfo(text, input.candidate);
};
