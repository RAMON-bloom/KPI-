// ブラインドレジュメ作成（候補者パイプライン →「ブラインドレジュメ作成」）のロジック。
//
// 流れ: ①対象範囲の候補者を匿名プロフィール化 → ②Geminiで推薦先の求人とのマッチ度を採点 →
// ③選んだ候補者ごとにGeminiで匿名レジュメ本文を生成 → ④表紙付きのHTMLを印刷ダイアログで
// 1つのPDFに保存。
//
// 個人情報の扱い: Geminiに送る段階で氏名・電話番号・メールアドレスは渡さない（面談メモ等の
// 本文中に出てくる氏名・連絡先も機械的に伏せ字にしてから送る）。生成結果にも念のため同じ伏せ字
// 処理をかける。年齢・学歴・在籍企業名・年収は伏せない（2026-10-14 ユーザー指示）。

import { GoogleGenAI, Type } from '@google/genai';

const MODEL = 'gemini-2.5-flash';

// index.tsxのCandidateのうち、この機能で使う項目だけ（構造的に互換）。
export interface BlindSourceCandidate {
  id: string;
  name: string;
  age?: number;
  jobType?: string;
  currentCompany?: string;
  education?: string;
  currentSalary?: number;
  salary?: number;
  desiredJoinTiming?: string;
  summary?: string;
  phoneNumber?: string;
  email?: string;
  memos?: { title?: string; content?: string }[];
  interviewSummary?: string;
  interviewMemo?: string;
  applications?: { companyName?: string; stage?: string }[];
}

export interface BlindPoolEntry {
  candidate: BlindSourceCandidate;
  ownerEmail: string;
  ownerLabel: string;
  status: string; // 表示用（選考中・非表示・掘り起し など）
}

export interface MatchResult {
  key: string; // `${ownerEmail}::${candidate.id}`
  score: number; // 0-100
  reason: string;
  concern: string;
  recommendedPosition: string; // 推薦先で最も合いそうなポジション（不明なら空文字）
}

export interface CompanyResearch {
  text: string;
  sources: { title: string; uri: string }[];
}

export interface BlindCareer {
  period: string;
  company: string;
  role: string;
  description: string;
}

export interface BlindResumeContent {
  recommendedPosition: string;
  headline: string;
  careerSummary: string;
  careers: BlindCareer[];
  skills: string[];
  matchPoints: string[];
  coverPoint: string;
}

export interface BlindDisplayOptions {
  showCurrentSalary: boolean;
  showDesiredSalary: boolean;
  showJoinTiming: boolean;
}

export type JobContentPart = { inlineData: { mimeType: string; data: string } } | { text: string };

export const poolKey = (entry: BlindPoolEntry) => `${entry.ownerEmail}::${entry.candidate.id}`;

// 候補者A, B, … Z, AA, AB …
export const blindLabel = (index: number): string => {
  let n = index;
  let s = '';
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
};

const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** 氏名（フルネーム・空白区切りの姓/名）、電話番号、メールアドレスを「〇〇」等に置き換える。 */
export const redactPersonalInfo = (text: string, c: Pick<BlindSourceCandidate, 'name' | 'phoneNumber' | 'email'>): string => {
  if (!text) return '';
  let out = text;
  const name = (c.name || '').trim();
  if (name) {
    const parts = name.split(/[\s　]+/).filter(Boolean);
    const tokens = new Set<string>();
    tokens.add(parts.join(''));
    if (parts.length > 1) {
      tokens.add(name);
      parts.forEach(p => { if (p.length >= 2) tokens.add(p); });
    }
    Array.from(tokens)
      .sort((a, b) => b.length - a.length)
      .forEach(token => {
        out = out.replace(new RegExp(escapeRegExp(token), 'g'), '〇〇');
      });
  }
  if (c.email?.trim()) out = out.split(c.email.trim()).join('（メール非開示）');
  const phoneDigits = (c.phoneNumber || '').replace(/\D/g, '');
  // 一般的な電話番号の書式（ハイフン・空白区切り含む）をすべて伏せる — 本人の番号以外が
  // 紛れていても出力先は社外なので、伏せて困ることはない。
  out = out.replace(/(?:\+81[-\s]?)?0\d{1,4}[-\s(]?\d{1,4}[-\s)]?\d{3,4}/g, m => {
    const digits = m.replace(/\D/g, '');
    return digits.length >= 10 ? '（電話番号非開示）' : m;
  });
  if (phoneDigits.length >= 10) out = out.split(phoneDigits).join('（電話番号非開示）');
  out = out.replace(/[\w.+-]+@[\w-]+\.[\w.-]+/g, '（メール非開示）');
  return out;
};

const truncate = (s: string, max: number) => (s.length > max ? `${s.slice(0, max)}…` : s);

export const collectNotes = (c: BlindSourceCandidate): string => {
  const parts: string[] = [];
  (c.memos || []).forEach(m => {
    const content = (m.content || '').trim();
    if (content) parts.push(`【${(m.title || 'メモ').trim()}】${content}`);
  });
  if (c.interviewSummary?.trim()) parts.push(`【面談ログ】${c.interviewSummary.trim()}`);
  if (c.interviewMemo?.trim()) parts.push(`【面談メモ】${c.interviewMemo.trim()}`);
  return parts.join('\n');
};

export const basicFacts = (c: BlindSourceCandidate): string[] => [
  c.age ? `年齢: ${c.age}歳` : '',
  c.jobType ? `職種: ${c.jobType}` : '',
  c.currentCompany ? `現職: ${c.currentCompany}` : '',
  c.education ? `学歴: ${c.education}` : '',
  c.currentSalary ? `現年収: ${c.currentSalary}万円` : '',
  c.salary ? `希望年収: ${c.salary}万円` : '',
  c.desiredJoinTiming ? `入社希望時期: ${c.desiredJoinTiming}` : '',
].filter(Boolean);

/** マッチング用の短い匿名プロフィール（氏名・連絡先なし）。 */
export const buildMatchingProfile = (entry: BlindPoolEntry, id: string, detailed: boolean): string => {
  const c = entry.candidate;
  const summary = redactPersonalInfo((c.summary || '').trim(), c);
  const notes = redactPersonalInfo(collectNotes(c), c);
  return [
    `[ID: ${id}] (ステータス: ${entry.status})`,
    ...basicFacts(c),
    summary ? `経歴概要: ${truncate(summary, detailed ? 1500 : 500)}` : '経歴概要: （未登録）',
    notes ? `面談メモ抜粋: ${truncate(notes, detailed ? 2500 : 300)}` : '',
  ].filter(Boolean).join('\n');
};

const newAi = () => new GoogleGenAI({ apiKey: process.env.API_KEY as string });

// ---- 所要時間の見積もり ----
// 画面に「推定時間・残り時間」を出すための目安。Geminiの応答時間は混み具合で大きく変わるので、
// 実際にかかった時間をこのブラウザのlocalStorageに記録し（指数移動平均）、次回以降の見積もりに
// 使う。初期値は gemini-2.5-flash での経験的な目安。
export const MATCH_CONCURRENCY = 4;
export const GENERATE_CONCURRENCY = 3;

export interface BlindTimings {
  researchSec: number; // Web検索での企業調査 1回
  matchSecPer60: number; // 採点 1グループ（60名換算）
  rerankSec: number; // 上位候補の再評価 1回
  generateSec: number; // レジュメ本文の生成 1名
}

const DEFAULT_TIMINGS: BlindTimings = { researchSec: 25, matchSecPer60: 30, rerankSec: 25, generateSec: 20 };
const TIMINGS_STORAGE_KEY = 'kpiBlindResumeTimings';

export const loadBlindTimings = (): BlindTimings => {
  try {
    const raw = localStorage.getItem(TIMINGS_STORAGE_KEY);
    if (!raw) return { ...DEFAULT_TIMINGS };
    const parsed = JSON.parse(raw);
    const pick = (k: keyof BlindTimings) => (typeof parsed[k] === 'number' && parsed[k] > 0 ? parsed[k] : DEFAULT_TIMINGS[k]);
    return { researchSec: pick('researchSec'), matchSecPer60: pick('matchSecPer60'), rerankSec: pick('rerankSec'), generateSec: pick('generateSec') };
  } catch {
    return { ...DEFAULT_TIMINGS };
  }
};

const recordBlindTiming = (kind: keyof BlindTimings, seconds: number) => {
  if (!(seconds > 0) || seconds > 600) return;
  try {
    const current = loadBlindTimings();
    current[kind] = Math.round((current[kind] * 0.7 + seconds * 0.3) * 10) / 10;
    localStorage.setItem(TIMINGS_STORAGE_KEY, JSON.stringify(current));
  } catch {
    // 記録できなくても見積もりが初期値のままになるだけ
  }
};

const MATCH_SIZE_FLOOR = 0.35; // 少人数でも1回の呼び出しには一定の時間がかかる
const matchChunkSeconds = (size: number, t: BlindTimings) => t.matchSecPer60 * Math.max(MATCH_SIZE_FLOOR, size / 60);

/** 候補者の照合（＋必要ならWeb検索）の見積もり。workUnits は進捗計算用の内訳。 */
export const estimateSearchPlan = (poolSize: number, needsResearch: boolean, t: BlindTimings = loadBlindTimings()) => {
  const chunkSizes: number[] = [];
  for (let i = 0; i < poolSize; i += MATCH_CHUNK_SIZE) chunkSizes.push(Math.min(MATCH_CHUNK_SIZE, poolSize - i));
  const rounds: number[][] = [];
  for (let i = 0; i < chunkSizes.length; i += MATCH_CONCURRENCY) rounds.push(chunkSizes.slice(i, i + MATCH_CONCURRENCY));
  const matchSec = rounds.reduce((sum, round) => sum + Math.max(...round.map(s => matchChunkSeconds(s, t))), 0);
  const researchSec = needsResearch ? t.researchSec : 0;
  const rerankSec = chunkSizes.length > 1 ? t.rerankSec : 0;
  return {
    totalSec: Math.max(1, Math.round(researchSec + matchSec + rerankSec)),
    researchSec,
    // 並列実行のため、1グループ完了ごとに照合全体の見積もりを均等に進める
    perChunkCreditSec: chunkSizes.length ? matchSec / chunkSizes.length : 0,
    rerankSec,
  };
};

/** レジュメ本文の生成（count名）の見積もり。 */
export const estimateGeneratePlan = (count: number, t: BlindTimings = loadBlindTimings()) => {
  const totalSec = Math.ceil(count / GENERATE_CONCURRENCY) * t.generateSec;
  return { totalSec: Math.max(1, Math.round(totalSec)), perItemCreditSec: count ? totalSec / count : 0 };
};

const runWithConcurrency = async <T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> => {
  const results: R[] = new Array(items.length);
  let next = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const i = next++;
      results[i] = await fn(items[i]);
    }
  });
  await Promise.all(workers);
  return results;
};

/**
 * 求人内容が無いときに、企業名（＋任意のポジション名）からWeb検索（GeminiのGoogle検索
 * グラウンディング）で企業情報と採用ニーズの見立てを作る。以降のマッチング・レジュメ生成では
 * これを求人情報の代わりに使う。検索ツールとJSON出力は同時に使えないため、ここはテキストで返す。
 */
export const researchCompany = async (companyName: string, positionTitle: string): Promise<CompanyResearch> => {
  const prompt = [
    `人材紹介会社のキャリアアドバイザーとして、「${companyName}」に候補者を推薦する準備のため、Web検索で次の点を調べて日本語で簡潔に整理してください。`,
    positionTitle ? `担当者が想定しているポジション: ${positionTitle}` : '',
    '1. 企業概要（事業内容・主要サービス・業界・規模・拠点）',
    '2. 現在募集中、または採用ニーズが高いと考えられる職種・ポジション（採用ページや求人媒体の掲載があればそれを優先し、無ければ事業内容から推測したものと明記する）',
    '3. 各ポジションで求められそうな経験・スキル・人物像、想定年収レンジ（分かる範囲で）',
    '4. 社風・カルチャーや選考で重視されそうな点',
    '同名の別企業と混同しないよう注意し、確認できなかった点は「不明」と書くこと。',
  ].filter(Boolean).join('\n');
  const startedAt = Date.now();
  const response = await newAi().models.generateContent({
    model: MODEL,
    contents: prompt,
    config: { tools: [{ googleSearch: {} }] },
  });
  recordBlindTiming('researchSec', (Date.now() - startedAt) / 1000);
  const text = (response.text || '').trim();
  if (!text) throw new Error('企業情報のWeb検索結果が空でした');
  const chunks = response.candidates?.[0]?.groundingMetadata?.groundingChunks || [];
  const seen = new Set<string>();
  const sources: { title: string; uri: string }[] = [];
  chunks.forEach(c => {
    const uri = c.web?.uri || '';
    if (!uri || seen.has(uri)) return;
    seen.add(uri);
    sources.push({ title: c.web?.title || uri, uri });
  });
  return { text, sources };
};

const MATCH_CHUNK_SIZE = 60;

const scoreChunk = async (
  jobParts: JobContentPart[],
  companyName: string,
  extraConditions: string,
  profiles: { id: string; text: string }[],
  topK: number,
  hasJobDescription: boolean,
): Promise<{ id: string; score: number; reason: string; concern: string; recommendedPosition?: string }[]> => {
  const instruction = [
    hasJobDescription
      ? `あなたは人材紹介会社のキャリアアドバイザーです。推薦先企業「${companyName || '（企業名未入力）'}」の求人情報（上記）に対して、以下の候補者リストの中からマッチしそうな候補者を選び、マッチ度を採点してください。`
      : `あなたは人材紹介会社のキャリアアドバイザーです。推薦先企業「${companyName}」の具体的な求人票は無く、上記はWeb検索で調べた企業情報と採用ニーズの見立てです。これをもとに、この企業で活躍が見込め、採用ニーズに合いそうな候補者を以下の候補者リストから選び、マッチ度を採点してください。`,
    extraConditions ? `担当者からの追加条件: ${extraConditions}` : '',
    hasJobDescription
      ? '採点基準: 求人の必須要件・歓迎要件と候補者の職種・経験・スキルの一致度を最重視し、年齢・年収レンジ（現年収・希望年収と求人の提示年収の整合）、入社時期も考慮する。0〜100の整数で、80以上は「すぐ推薦できる」、60台は「条件次第で推薦可」、40未満は「ほぼ合わない」の目安。'
      : '採点基準: 企業の事業内容と採用ニーズが高そうな職種に対し、候補者の職種・経験・スキルがどれだけ一致するかを最重視し、業界・企業規模への適性、年齢・年収レンジの現実性も考慮する。0〜100の整数で、80以上は「すぐ推薦できる」、60台は「条件次第で推薦可」、40未満は「ほぼ合わない」の目安。',
    `マッチ度の高い順に最大${topK}名を返してください（明らかに合わない候補者は含めなくてよい）。reasonは推薦理由を60字程度、concernは懸念点（無ければ空文字）を40字程度で、いずれも候補者の氏名には触れずに書くこと。`,
    'recommendedPositionには、その候補者がこの企業で最も活躍できそうなポジション名（例: 法人営業（製造業向け）、経理（連結決算））を20字程度で書く。求人が1つのポジションに限られる場合はそのポジション名でよい。判断できなければ空文字。',
    'ステータスが「見送り」「辞退」「非表示」「掘り起し」の候補者も、過去の候補者の掘り起こし対象として同じ基準で評価してよい。',
    '',
    '候補者リスト:',
    profiles.map(p => p.text).join('\n\n'),
  ].filter(s => s !== '').join('\n');
  const response = await newAi().models.generateContent({
    model: MODEL,
    contents: { parts: [...jobParts, { text: instruction }] },
    config: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            id: { type: Type.STRING, description: '候補者リストの[ID: ...]の値' },
            score: { type: Type.NUMBER, description: 'マッチ度（0〜100の整数）' },
            reason: { type: Type.STRING, description: '推薦理由' },
            concern: { type: Type.STRING, description: '懸念点（無ければ空文字）' },
            recommendedPosition: { type: Type.STRING, description: '推奨ポジション（判断できなければ空文字）' },
          },
          required: ['id', 'score', 'reason', 'concern', 'recommendedPosition'],
        },
      },
    },
  });
  const parsed = JSON.parse((response.text || '[]').trim());
  return Array.isArray(parsed) ? parsed : [];
};

/**
 * 候補者プールを推薦先の求人と照合し、マッチ度順に並べた結果を返す。人数が多い場合は
 * MATCH_CHUNK_SIZE人ずつに分けて並列に採点し、上位をより詳しいプロフィールで再採点する
 * （分割したグループ間で採点の甘辛がぶれるのをならすため）。
 */
export const matchCandidates = async (params: {
  pool: BlindPoolEntry[];
  jobParts: JobContentPart[];
  companyName: string;
  extraConditions: string;
  wantedCount: number;
  hasJobDescription: boolean;
  onProgress?: (message: string) => void;
  // 残り時間表示用: 採点1グループ／再評価が終わるたびに呼ばれる
  onChunkDone?: () => void;
  onRerankDone?: () => void;
}): Promise<MatchResult[]> => {
  const { pool, jobParts, companyName, extraConditions, wantedCount, hasJobDescription, onProgress, onChunkDone, onRerankDone } = params;
  const idToKey = new Map<string, string>();
  const keyToEntry = new Map<string, BlindPoolEntry>();
  const profiles = pool.map((entry, i) => {
    const id = `P${String(i + 1).padStart(4, '0')}`;
    idToKey.set(id, poolKey(entry));
    keyToEntry.set(poolKey(entry), entry);
    return { id, entry };
  });
  const shortlistSize = Math.min(pool.length, Math.max(wantedCount * 3, wantedCount + 5, 10), 40);

  const chunks: typeof profiles[] = [];
  for (let i = 0; i < profiles.length; i += MATCH_CHUNK_SIZE) chunks.push(profiles.slice(i, i + MATCH_CHUNK_SIZE));
  let done = 0;
  onProgress?.(`${pool.length}名の候補者を求人と照合しています…`);
  const chunkResults = await runWithConcurrency(chunks, MATCH_CONCURRENCY, async chunk => {
    const chunkStartedAt = Date.now();
    const res = await scoreChunk(
      jobParts, companyName, extraConditions,
      chunk.map(p => ({ id: p.id, text: buildMatchingProfile(p.entry, p.id, chunks.length === 1) })),
      chunks.length === 1 ? shortlistSize : Math.min(chunk.length, Math.max(wantedCount * 2, 8)),
      hasJobDescription,
    );
    recordBlindTiming('matchSecPer60', (Date.now() - chunkStartedAt) / 1000 / Math.max(MATCH_SIZE_FLOOR, chunk.length / 60));
    done++;
    onChunkDone?.();
    if (chunks.length > 1) onProgress?.(`候補者を照合中…（${done}/${chunks.length}グループ完了）`);
    return res;
  });

  let scored = chunkResults.flat().filter(r => idToKey.has(r.id));
  if (chunks.length > 1 && scored.length > 0) {
    scored.sort((a, b) => b.score - a.score);
    const finalists = scored.slice(0, shortlistSize);
    onProgress?.(`上位${finalists.length}名を詳しいプロフィールで再評価しています…`);
    const byId = new Map(profiles.map(p => [p.id, p]));
    const rerankStartedAt = Date.now();
    const rescored = await scoreChunk(
      jobParts, companyName, extraConditions,
      finalists.map(f => ({ id: f.id, text: buildMatchingProfile(byId.get(f.id)!.entry, f.id, true) })),
      finalists.length,
      hasJobDescription,
    );
    recordBlindTiming('rerankSec', (Date.now() - rerankStartedAt) / 1000);
    onRerankDone?.();
    const rescoredValid = rescored.filter(r => idToKey.has(r.id));
    if (rescoredValid.length > 0) scored = rescoredValid;
  }

  const seen = new Set<string>();
  return scored
    .sort((a, b) => b.score - a.score)
    .filter(r => {
      if (seen.has(r.id)) return false;
      seen.add(r.id);
      return true;
    })
    .map(r => {
      const key = idToKey.get(r.id)!;
      const c = keyToEntry.get(key)!.candidate;
      return {
        key,
        score: Math.max(0, Math.min(100, Math.round(Number(r.score) || 0))),
        reason: redactPersonalInfo(r.reason || '', c),
        concern: redactPersonalInfo(r.concern || '', c),
        recommendedPosition: redactPersonalInfo((r.recommendedPosition || '').trim(), c),
      };
    });
};

/** 1名分のブラインドレジュメ本文を生成する（氏名・連絡先は渡さない）。 */
export const generateBlindResume = async (params: {
  entry: BlindPoolEntry;
  jobParts: JobContentPart[];
  companyName: string;
  recommendedPosition: string;
}): Promise<BlindResumeContent> => {
  const { entry, jobParts, companyName, recommendedPosition } = params;
  const c = entry.candidate;
  const summary = redactPersonalInfo((c.summary || '').trim(), c);
  const notes = redactPersonalInfo(collectNotes(c), c);
  // 推薦先も求人内容も無い場合（候補者カードから1名分だけ作るとき）は、特定の企業に寄せない
  // 汎用のブラインドレジュメにする。
  const hasTarget = !!companyName || jobParts.length > 0;
  const instruction = [
    hasTarget
      ? `あなたは人材紹介会社のキャリアアドバイザーです。推薦先企業「${companyName || '（企業名未入力）'}」（求人情報または企業情報は上記）へ候補者を匿名で推薦するための「ブラインドレジュメ」の本文を作成してください。`
      : 'あなたは人材紹介会社のキャリアアドバイザーです。推薦先企業を特定せず、複数の企業へ候補者を匿名で打診するときに使える汎用の「ブラインドレジュメ」の本文を作成してください。',
    recommendedPosition ? `推薦するポジション: ${recommendedPosition}（このポジションへの適合が伝わるように書く）` : '',
    '',
    '■候補者の登録情報',
    ...basicFacts(c),
    `経歴概要: ${truncate(summary || '（未登録）', 4000)}`,
    notes ? `面談メモ・面談ログ:\n${truncate(notes, 8000)}` : '',
    '',
    '■作成ルール',
    '- 書かれている事実だけを使い、経歴・スキル・資格・数値を創作しない。情報が無い項目は空にする（periodが不明なら空文字）。',
    '- 氏名・ふりがな・住所・電話番号・メールアドレス・生年月日・顔写真・家族構成・健康状態など、個人を特定できる情報やプライベートな事情は書かない。',
    '- 面談で聞いた転職理由の私的・ネガティブな事情、他社の選考状況、年収交渉の内情、担当者の所感（懸念点）は書かない。企業に提出する推薦資料として適切な表現にする。',
    '- 在籍企業名・職種・学歴はそのまま書いてよい。',
    '- careersは新しい順。descriptionは担当業務・実績を2〜4文で具体的に。',
    hasTarget
      ? '- matchPointsは推薦先の求人要件と結び付けた適合ポイントを3つ（各60字程度）。'
      : '- matchPointsは企業にとっての候補者の強み・アピールポイントを3つ（各60字程度）。',
    '- headlineは候補者を一言で表すキャッチ（30字程度）、coverPointは表紙の一覧に載せる適合ポイントの要約（50字程度）。',
  ].filter(s => s !== '').join('\n');
  const generateStartedAt = Date.now();
  const response = await newAi().models.generateContent({
    model: MODEL,
    contents: { parts: [...jobParts, { text: instruction }] },
    config: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          headline: { type: Type.STRING },
          careerSummary: { type: Type.STRING, description: '職務要約（200〜300字）' },
          careers: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                period: { type: Type.STRING, description: '在籍期間（例: 2019年4月〜現在）。不明なら空文字' },
                company: { type: Type.STRING },
                role: { type: Type.STRING, description: '部署・役職・職種' },
                description: { type: Type.STRING },
              },
              required: ['period', 'company', 'role', 'description'],
            },
          },
          skills: { type: Type.ARRAY, items: { type: Type.STRING }, description: '活かせる経験・スキル・資格' },
          matchPoints: { type: Type.ARRAY, items: { type: Type.STRING } },
          coverPoint: { type: Type.STRING },
        },
        required: ['headline', 'careerSummary', 'careers', 'skills', 'matchPoints', 'coverPoint'],
      },
    },
  });
  recordBlindTiming('generateSec', (Date.now() - generateStartedAt) / 1000);
  const parsed = JSON.parse((response.text || '{}').trim());
  const r = (s: unknown) => redactPersonalInfo(typeof s === 'string' ? s.trim() : '', c);
  return {
    recommendedPosition,
    headline: r(parsed.headline),
    careerSummary: r(parsed.careerSummary),
    careers: (Array.isArray(parsed.careers) ? parsed.careers : []).map((x: any) => ({
      period: r(x?.period), company: r(x?.company), role: r(x?.role), description: r(x?.description),
    })),
    skills: (Array.isArray(parsed.skills) ? parsed.skills : []).map(r).filter(Boolean),
    matchPoints: (Array.isArray(parsed.matchPoints) ? parsed.matchPoints : []).map(r).filter(Boolean),
    coverPoint: r(parsed.coverPoint),
  };
};

/** 生成結果（編集後を含む）に、元データの氏名・電話・メールが残っていないか確認する。 */
export const findLeakedPersonalInfo = (content: BlindResumeContent, c: BlindSourceCandidate): string[] => {
  const all = [
    content.recommendedPosition, content.headline, content.careerSummary, content.coverPoint,
    ...content.careers.flatMap(x => [x.period, x.company, x.role, x.description]),
    ...content.skills, ...content.matchPoints,
  ].join('\n');
  const leaks: string[] = [];
  const name = (c.name || '').trim();
  if (name) {
    const parts = name.split(/[\s　]+/).filter(Boolean);
    const tokens = new Set([parts.join(''), ...(parts.length > 1 ? parts.filter(p => p.length >= 2) : [])]);
    tokens.forEach(t => { if (t && all.includes(t)) leaks.push(`氏名（${t}）`); });
  }
  if (c.email?.trim() && all.includes(c.email.trim())) leaks.push('メールアドレス');
  const phoneDigits = (c.phoneNumber || '').replace(/\D/g, '');
  if (phoneDigits.length >= 10 && all.replace(/\D/g, '').includes(phoneDigits)) leaks.push('電話番号');
  if (/[\w.+-]+@[\w-]+\.[\w.-]+/.test(all) && !leaks.includes('メールアドレス')) leaks.push('メールアドレスらしき文字列');
  return leaks;
};

const esc = (s: string) => (s || '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const formatJoinTiming = (v?: string) => {
  if (!v) return '';
  const m = /^(\d{4})-(\d{2})$/.exec(v);
  return m ? `${m[1]}年${Number(m[2])}月` : v;
};

export interface BlindResumeDocItem {
  label: string; // 'A'
  candidate: BlindSourceCandidate;
  content: BlindResumeContent;
}

/** 表紙＋候補者ごとのページを、印刷（PDF保存）用の1つのHTML文書にする。 */
export const renderBlindResumeHtml = (params: {
  companyName: string;
  positionTitle: string;
  createdDate: string; // yyyy-mm-dd
  items: BlindResumeDocItem[];
  options: BlindDisplayOptions;
  documentTitle: string;
  // false: 表紙を付けず候補者のページだけにする（候補者カードから1名分だけ出力するとき）。
  includeCover?: boolean;
}): string => {
  const { companyName, positionTitle, createdDate, items, options, documentTitle, includeCover = true } = params;
  const [y, mo, d] = createdDate.split('-');
  const dateLabel = `${y}年${Number(mo)}月${Number(d)}日`;
  const salaryCell = (v?: number) => (v ? `${v}万円` : '—');
  const showPosition = items.some(i => i.content.recommendedPosition);

  const coverRows = items.map(({ label, candidate: c, content }) => `
      <tr>
        <td class="nowrap"><strong>候補者${esc(label)}</strong></td>
        ${showPosition ? `<td>${esc(content.recommendedPosition || '—')}</td>` : ''}
        <td class="nowrap">${c.age ? `${c.age}歳` : '—'}</td>
        <td>${esc(c.currentCompany || '—')}${c.jobType ? `<div class="sub">${esc(c.jobType)}</div>` : ''}</td>
        ${options.showCurrentSalary ? `<td class="nowrap">${salaryCell(c.currentSalary)}</td>` : ''}
        ${options.showDesiredSalary ? `<td class="nowrap">${salaryCell(c.salary)}</td>` : ''}
        <td>${esc(content.coverPoint)}</td>
      </tr>`).join('');

  const pages = items.map(({ label, candidate: c, content }) => {
    const facts: [string, string][] = [
      ...(content.recommendedPosition ? [['推奨ポジション', content.recommendedPosition] as [string, string]] : []),
      ['年齢', c.age ? `${c.age}歳` : '—'],
      ['最終学歴', c.education || '—'],
      ['現職', c.currentCompany || '—'],
      ['職種', c.jobType || '—'],
    ];
    if (options.showCurrentSalary) facts.push(['現年収', salaryCell(c.currentSalary)]);
    if (options.showDesiredSalary) facts.push(['希望年収', salaryCell(c.salary)]);
    if (options.showJoinTiming) facts.push(['入社希望時期', formatJoinTiming(c.desiredJoinTiming) || '—']);
    const factCells = facts.map(([k, v]) => `<div class="fact"><span class="k">${esc(k)}</span><span class="v">${esc(v)}</span></div>`).join('');
    const careers = content.careers.length
      ? content.careers.map(x => `
          <div class="career">
            <div class="career-head"><span class="company">${esc(x.company)}</span>${x.period ? `<span class="period">${esc(x.period)}</span>` : ''}</div>
            ${x.role ? `<div class="role">${esc(x.role)}</div>` : ''}
            ${x.description ? `<div class="desc">${esc(x.description)}</div>` : ''}
          </div>`).join('')
      : '<p class="muted">（登録情報なし）</p>';
    const list = (arr: string[]) => (arr.length ? `<ul>${arr.map(s => `<li>${esc(s)}</li>`).join('')}</ul>` : '<p class="muted">（登録情報なし）</p>');
    const title = includeCover ? `候補者${esc(label)}` : 'ブラインドレジュメ';
    const destination = !includeCover && companyName
      ? `<p class="to-inline">${esc(companyName)} 御中${positionTitle ? `　／　対象ポジション: ${esc(positionTitle)}` : ''}</p>`
      : '';
    return `
    <section class="page">
      <div class="page-tag">${includeCover ? `候補者${esc(label)}（ブラインド）` : `作成日: ${dateLabel}`}</div>
      <h2>${title}</h2>
      ${destination}
      ${content.headline ? `<p class="headline">${esc(content.headline)}</p>` : ''}
      <div class="facts">${factCells}</div>
      <h3>推薦ポイント</h3>
      ${list(content.matchPoints)}
      <h3>職務要約</h3>
      <p class="pre">${esc(content.careerSummary || '（登録情報なし）')}</p>
      <h3>職務経歴</h3>
      ${careers}
      <h3>活かせる経験・スキル</h3>
      ${list(content.skills)}
      ${includeCover ? '' : '<p class="note">※ 個人情報保護のため、氏名・連絡先等の個人を特定できる情報は記載しておりません。</p>'}
    </section>`;
  }).join('');

  return `<!doctype html>
<html lang="ja"><head><meta charset="utf-8"><title>${esc(documentTitle)}</title>
<style>
  @page { size: A4; margin: 14mm 14mm 16mm; }
  * { box-sizing: border-box; }
  body { font-family: "Hiragino Kaku Gothic ProN", "Yu Gothic", "Meiryo", "Noto Sans JP", sans-serif; color: #1f2937; font-size: 10.5pt; line-height: 1.6; margin: 0; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .page { page-break-after: always; break-after: page; position: relative; padding-top: 6mm; }
  .page:last-child { page-break-after: auto; break-after: auto; }
  .page-tag { position: absolute; top: 0; right: 0; font-size: 9pt; color: #6b7280; border: 1px solid #d1d5db; padding: 1px 8px; border-radius: 3px; }
  .cover h1 { font-size: 20pt; margin: 18mm 0 4mm; letter-spacing: 0.05em; }
  .cover .to { font-size: 13pt; margin: 0 0 2mm; }
  .cover .meta { color: #6b7280; margin-bottom: 8mm; }
  .note { font-size: 9pt; color: #6b7280; margin-top: 6mm; }
  .to-inline { margin: 0 0 2mm; color: #374151; }
  table { width: 100%; border-collapse: collapse; font-size: 9.5pt; }
  th, td { border: 1px solid #d1d5db; padding: 5px 7px; vertical-align: top; text-align: left; }
  th { background: #eef2f7; white-space: nowrap; }
  td.nowrap { white-space: nowrap; }
  .sub { color: #6b7280; font-size: 8.5pt; }
  h2 { font-size: 16pt; margin: 2mm 0 1mm; border-bottom: 2px solid #1e3a5f; padding-bottom: 1mm; color: #1e3a5f; }
  .headline { font-size: 11.5pt; font-weight: bold; margin: 2mm 0 3mm; }
  h3 { font-size: 11pt; margin: 5mm 0 1.5mm; padding-left: 6px; border-left: 4px solid #1e3a5f; break-after: avoid; page-break-after: avoid; }
  .facts { display: grid; grid-template-columns: 1fr 1fr; border: 1px solid #d1d5db; }
  .fact { display: flex; border-bottom: 1px solid #e5e7eb; }
  .fact:nth-child(odd) { border-right: 1px solid #e5e7eb; }
  .fact .k { width: 28mm; flex-shrink: 0; background: #eef2f7; padding: 3px 7px; font-weight: bold; font-size: 9.5pt; }
  .fact .v { padding: 3px 7px; }
  ul { margin: 0; padding-left: 1.3em; }
  li { margin: 0.5mm 0; }
  .pre { white-space: pre-wrap; margin: 0; }
  .career { border: 1px solid #e5e7eb; border-radius: 3px; padding: 2.5mm 3mm; margin-bottom: 2mm; break-inside: avoid; page-break-inside: avoid; }
  .career-head { display: flex; justify-content: space-between; gap: 8px; font-weight: bold; }
  .career .period { font-weight: normal; color: #4b5563; white-space: nowrap; }
  .career .role { color: #374151; font-size: 9.5pt; }
  .career .desc { margin-top: 1mm; white-space: pre-wrap; }
  .muted { color: #9ca3af; margin: 0; }
</style></head>
<body>
  ${includeCover ? `<section class="page cover">
    <h1>候補者のご推薦（ブラインドレジュメ）</h1>
    <p class="to">${esc(companyName || '')}${companyName ? ' 御中' : ''}</p>
    <p class="meta">${positionTitle ? `対象ポジション: ${esc(positionTitle)}　／　` : ''}作成日: ${dateLabel}　／　ご推薦人数: ${items.length}名</p>
    <table>
      <thead><tr><th>候補者</th>${showPosition ? '<th>推奨ポジション</th>' : ''}<th>年齢</th><th>現職</th>${options.showCurrentSalary ? '<th>現年収</th>' : ''}${options.showDesiredSalary ? '<th>希望年収</th>' : ''}<th>適合ポイント</th></tr></thead>
      <tbody>${coverRows}</tbody>
    </table>
    <p class="note">※ 個人情報保護のため、氏名・連絡先等の個人を特定できる情報は記載しておりません。ご興味をお持ちいただけた候補者がいらっしゃいましたら、候補者記号（例: 候補者A）でお知らせください。</p>
  </section>` : ''}
  ${pages}
</body></html>`;
};
