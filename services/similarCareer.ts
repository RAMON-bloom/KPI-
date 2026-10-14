// 類似経歴リサーチ（候補者カードの詳細／新規登録直後 →「類似経歴の過去事例をリサーチ」）のロジック。
//
// 流れ: ①対象範囲（全ユーザー or 自分）の候補者のうち、選考企業の記録がある人をプール化 →
// ②Geminiで対象候補者と経歴が似ている人を採点して上位を選ぶ → ③似ている人たちの応募企業と
// 選考結果（どのフェーズまで進んだか・見送り/辞退/内定）をアプリ内データからそのまま集計 →
// ④Geminiでその実績から傾向と、応募をおすすめできる企業・注意が要る企業を整理する。
//
// 個人情報の扱い: ブラインドレジュメと同じく、Geminiには氏名・電話番号・メールアドレスを渡さない
// （本文中に出てくる氏名・連絡先も redactPersonalInfo で伏せ字にしてから送る）。

import { GoogleGenAI, Type } from '@google/genai';
import { redactPersonalInfo } from './blindResume';

const MODEL = 'gemini-2.5-flash';
const CHUNK_SIZE = 60;
const CONCURRENCY = 4;

// index.tsxのCandidate / CompanyApplicationのうち、この機能で使う項目だけ（構造的に互換）。
export interface SimilarSourceApplication {
  id: string;
  companyName: string;
  stage: string;
  exitedFromStage?: string;
  exitedAt?: string;
  memo?: string;
  isHidden?: boolean;
  stageHistory?: { stage: string; date?: string }[];
}

export interface SimilarSourceCandidate {
  id: string;
  name: string;
  age?: number;
  jobType?: string;
  currentCompany?: string;
  education?: string;
  currentSalary?: number;
  salary?: number;
  summary?: string;
  phoneNumber?: string;
  email?: string;
  createdAt?: string;
  memos?: { title?: string; content?: string }[];
  interviewSummary?: string;
  interviewMemo?: string;
  applications: SimilarSourceApplication[];
}

export interface SimilarPoolEntry {
  candidate: SimilarSourceCandidate;
  ownerEmail: string;
  ownerLabel: string;
}

export interface SimilarMatch {
  key: string; // `${ownerEmail}::${candidate.id}`
  score: number; // 0-100
  reason: string;
}

export type OutcomeKind = 'offer' | 'declined' | 'withdrawn' | 'in_progress' | 'not_started';

export interface ApplicationOutcome {
  kind: OutcomeKind;
  label: string; // 表示用（例: 「1次面接で見送り」）
  reachedStage: string; // 到達した一番先のフェーズ
}

export interface CompanyStat {
  companyName: string;
  applicants: number;
  passedDocument: number; // 書類選考を通過（面接以降へ進んだ）
  offers: number;
  declined: number;
  withdrawn: number;
  inProgress: number;
  candidateKeys: string[];
}

export interface SimilarAnalysis {
  overview: string;
  passingFactors: string[];
  riskFactors: string[];
  recommendedCompanies: { companyName: string; reason: string }[];
  cautionCompanies: { companyName: string; reason: string }[];
  advice: string;
}

export const similarKey = (e: SimilarPoolEntry) => `${e.ownerEmail}::${e.candidate.id}`;

// 選考が前に進んだ順。お見送り・選考辞退は exitedFromStage で「どこで止まったか」を見る。
const FORWARD_ORDER = ['打診', '書類選考', '適性検査', 'カジュアル面談', '1次面接', '2次面接', '最終面接', 'オファー面談', '内定', '内定承諾'];
const stageRank = (s?: string) => (s ? FORWARD_ORDER.indexOf(s) : -1);
const OFFER_STAGES = new Set(['内定', '内定承諾', '内定承諾後辞退']);

/** 1件の選考の結果を、到達フェーズと「内定・見送り・辞退・選考中」に分類する。 */
export const classifyApplication = (app: SimilarSourceApplication): ApplicationOutcome => {
  // 到達した一番先のフェーズ — 履歴があればそこから、無ければ現在／離脱直前のフェーズ。
  let reached = '';
  const consider = (s?: string) => { if (s && stageRank(s) > stageRank(reached)) reached = s; };
  (app.stageHistory || []).forEach(h => consider(h.stage));
  consider(app.stage);
  consider(app.exitedFromStage);
  if (app.stage === '内定承諾後辞退') reached = reached || '内定';
  if (OFFER_STAGES.has(app.stage)) {
    const label = app.stage === '内定承諾' ? '内定承諾' : app.stage === '内定承諾後辞退' ? '内定（承諾後辞退）' : '内定';
    return { kind: 'offer', label, reachedStage: reached || '内定' };
  }
  if (app.stage === 'お見送り') {
    const at = app.exitedFromStage || (reached && reached !== '打診' ? reached : '');
    return { kind: 'declined', label: at ? `${at}で見送り` : '見送り', reachedStage: at || reached };
  }
  if (app.stage === '選考辞退') {
    const at = app.exitedFromStage || (reached && reached !== '打診' ? reached : '');
    return { kind: 'withdrawn', label: at ? `${at}で辞退` : '選考辞退', reachedStage: at || reached };
  }
  if (app.stage === '打診') return { kind: 'not_started', label: '打診（応募前）', reachedStage: '打診' };
  return { kind: 'in_progress', label: `${app.stage}（選考中）`, reachedStage: reached || app.stage };
};

/** 書類選考を通過したか（面接・適性検査以降に進んだ、または内定）。 */
const passedDocumentScreening = (o: ApplicationOutcome) =>
  o.kind === 'offer' || stageRank(o.reachedStage) > stageRank('書類選考');

/** プールに入れる選考（企業名があり、非表示でないもの）。 */
export const usableApplications = (c: SimilarSourceCandidate) =>
  (c.applications || []).filter(a => !a.isHidden && (a.companyName || '').trim());

const truncate = (s: string, max: number) => (s.length > max ? `${s.slice(0, max)}…` : s);

const collectNotes = (c: SimilarSourceCandidate): string => {
  const parts: string[] = [];
  (c.memos || []).forEach(m => {
    const content = (m.content || '').trim();
    if (content) parts.push(`【${(m.title || 'メモ').trim()}】${content}`);
  });
  if (c.interviewSummary?.trim()) parts.push(`【面談ログ】${c.interviewSummary.trim()}`);
  if (c.interviewMemo?.trim()) parts.push(`【面談メモ】${c.interviewMemo.trim()}`);
  return parts.join('\n');
};

const basicFacts = (c: SimilarSourceCandidate): string[] => [
  c.age ? `年齢: ${c.age}歳` : '',
  c.jobType ? `職種: ${c.jobType}` : '',
  c.currentCompany ? `現職: ${c.currentCompany}` : '',
  c.education ? `学歴: ${c.education}` : '',
  c.currentSalary ? `現年収: ${c.currentSalary}万円` : '',
  c.salary ? `希望年収: ${c.salary}万円` : '',
].filter(Boolean);

/** 経歴の類似度判定用の匿名プロフィール（氏名・連絡先なし、選考結果は含めない）。 */
const buildProfile = (c: SimilarSourceCandidate, id: string, detailed: boolean): string => {
  const summary = redactPersonalInfo((c.summary || '').trim(), c);
  const notes = detailed ? redactPersonalInfo(collectNotes(c), c) : '';
  return [
    `[ID: ${id}]`,
    ...basicFacts(c),
    summary ? `経歴概要: ${truncate(summary, detailed ? 1500 : 450)}` : '経歴概要: （未登録）',
    notes ? `面談メモ抜粋: ${truncate(notes, 1200)}` : '',
  ].filter(Boolean).join('\n');
};

/** 対象候補者の経歴が、類似度判定に足る程度に登録されているか。 */
export const hasEnoughProfile = (c: SimilarSourceCandidate) =>
  !!((c.summary || '').trim() || (c.jobType || '').trim() || (c.currentCompany || '').trim() || collectNotes(c).trim());

const newAi = () => new GoogleGenAI({ apiKey: process.env.API_KEY as string });

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

const scoreChunk = async (
  targetProfile: string,
  profiles: { id: string; text: string }[],
  topK: number,
): Promise<{ id: string; score: number; reason: string }[]> => {
  const instruction = [
    'あなたは人材紹介会社のキャリアアドバイザーです。次の「対象候補者」と経歴が似ている人を「過去の候補者リスト」から選び、類似度を採点してください。',
    '目的: 似た経歴の人がどの企業に応募し、どんな選考結果になったかを参考にするため。',
    '採点基準: 職種・担当業務・業界・経験年数（年齢）・スキルの近さを最重視し、現職の企業規模・学歴・年収レンジの近さも考慮する。0〜100の整数で、80以上は「ほぼ同じタイプの経歴」、60台は「職種や業界が近い」、40未満は「あまり似ていない」の目安。',
    `類似度の高い順に最大${topK}名を返してください（明らかに似ていない人は含めなくてよい）。reasonは似ている点を40字程度で、氏名には触れずに書くこと。`,
    '',
    '■対象候補者',
    targetProfile,
    '',
    '■過去の候補者リスト',
    profiles.map(p => p.text).join('\n\n'),
  ].join('\n');
  const response = await newAi().models.generateContent({
    model: MODEL,
    contents: instruction,
    config: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            id: { type: Type.STRING, description: '候補者リストの[ID: ...]の値' },
            score: { type: Type.NUMBER, description: '類似度（0〜100の整数）' },
            reason: { type: Type.STRING, description: '似ている点' },
          },
          required: ['id', 'score', 'reason'],
        },
      },
    },
  });
  const parsed = JSON.parse((response.text || '[]').trim());
  return Array.isArray(parsed) ? parsed : [];
};

/**
 * プールの中から対象候補者と経歴が似ている人を選ぶ。人数が多い場合は CHUNK_SIZE 人ずつ並列に
 * 採点し、上位を詳しいプロフィールで再採点する（グループ間の採点の甘辛をならすため）。
 */
export const findSimilarCandidates = async (params: {
  target: SimilarSourceCandidate;
  pool: SimilarPoolEntry[];
  wantedCount: number;
  onProgress?: (message: string) => void;
}): Promise<SimilarMatch[]> => {
  const { target, pool, wantedCount, onProgress } = params;
  const targetProfile = buildProfile(target, 'TARGET', true).replace('[ID: TARGET]\n', '');
  const idToKey = new Map<string, string>();
  const keyToEntry = new Map<string, SimilarPoolEntry>();
  const profiles = pool.map((entry, i) => {
    const id = `S${String(i + 1).padStart(4, '0')}`;
    idToKey.set(id, similarKey(entry));
    keyToEntry.set(similarKey(entry), entry);
    return { id, entry };
  });
  const shortlist = Math.min(pool.length, Math.max(wantedCount * 2, 20), 40);
  const chunks: typeof profiles[] = [];
  for (let i = 0; i < profiles.length; i += CHUNK_SIZE) chunks.push(profiles.slice(i, i + CHUNK_SIZE));
  let done = 0;
  onProgress?.(`${pool.length}名の過去の候補者と経歴を照合しています…`);
  const chunkResults = await runWithConcurrency(chunks, CONCURRENCY, async chunk => {
    const res = await scoreChunk(
      targetProfile,
      chunk.map(p => ({ id: p.id, text: buildProfile(p.entry.candidate, p.id, chunks.length === 1) })),
      chunks.length === 1 ? shortlist : Math.min(chunk.length, Math.max(wantedCount, 8)),
    );
    done++;
    if (chunks.length > 1) onProgress?.(`経歴を照合中…（${done}/${chunks.length}グループ完了）`);
    return res;
  });

  let scored = chunkResults.flat().filter(r => idToKey.has(r.id));
  if (chunks.length > 1 && scored.length > 0) {
    scored.sort((a, b) => b.score - a.score);
    const finalists = scored.slice(0, shortlist);
    onProgress?.(`上位${finalists.length}名を詳しいプロフィールで再評価しています…`);
    const byId = new Map(profiles.map(p => [p.id, p]));
    const rescored = await scoreChunk(
      targetProfile,
      finalists.map(f => ({ id: f.id, text: buildProfile(byId.get(f.id)!.entry.candidate, f.id, true) })),
      finalists.length,
    );
    const valid = rescored.filter(r => idToKey.has(r.id));
    if (valid.length > 0) scored = valid;
  }

  const seen = new Set<string>();
  return scored
    .sort((a, b) => b.score - a.score)
    .filter(r => (seen.has(r.id) ? false : (seen.add(r.id), true)))
    .slice(0, wantedCount)
    .map(r => {
      const key = idToKey.get(r.id)!;
      const c = keyToEntry.get(key)!.candidate;
      return {
        key,
        score: Math.max(0, Math.min(100, Math.round(Number(r.score) || 0))),
        reason: redactPersonalInfo(r.reason || '', c),
      };
    });
};

// ---- クイックリサーチ（Geminiを使わない類似度判定） ----
// 職種・経歴概要の文字の重なり（2文字ずつの組の一致率）と、年齢・現年収の近さ、現職企業の一致を
// 重み付けして0〜100点にする。AI判定ほど文脈は読めないが、数百人でも一瞬で終わる。

const bigrams = (text: string): Set<string> => {
  const s = (text || '').toLowerCase().replace(/[\s　、。・,.()（）「」【】\-ー/／:：]+/g, '');
  const set = new Set<string>();
  for (let i = 0; i < s.length - 1; i++) set.add(s.slice(i, i + 2));
  return set;
};

// 一致率: 共通の組の数 ÷ 小さい方の組の数（長さの違う文章どうしでも、短い方がどれだけ含まれるか）
const overlap = (a: Set<string>, b: Set<string>): number => {
  if (a.size === 0 || b.size === 0) return 0;
  const [small, large] = a.size <= b.size ? [a, b] : [b, a];
  let hit = 0;
  small.forEach(x => { if (large.has(x)) hit++; });
  return hit / small.size;
};

const careerText = (c: SimilarSourceCandidate) => [c.summary || '', collectNotes(c).slice(0, 1500)].join(' ');

export const findSimilarCandidatesQuick = (params: {
  target: SimilarSourceCandidate;
  pool: SimilarPoolEntry[];
  wantedCount: number;
  normalizeCompanyName: (name: string) => string;
}): SimilarMatch[] => {
  const { target, pool, wantedCount, normalizeCompanyName } = params;
  const tJob = bigrams(target.jobType || '');
  const tCareer = bigrams(careerText(target));
  const tCompany = normalizeCompanyName(target.currentCompany || '');
  const scored = pool.map(entry => {
    const c = entry.candidate;
    const parts: { weight: number; value: number }[] = [];
    const reasons: string[] = [];
    const cJob = bigrams(c.jobType || '');
    if (tJob.size && cJob.size) {
      const v = overlap(tJob, cJob);
      parts.push({ weight: 35, value: v });
      if (v >= 0.5) reasons.push(`職種が近い（${c.jobType}）`);
    }
    const cCareer = bigrams(careerText(c));
    if (tCareer.size && cCareer.size) {
      // 文章どうしは偶然の一致も多いので、一致率0.6でほぼ満点になるよう伸ばす
      const v = Math.min(1, overlap(tCareer, cCareer) / 0.6);
      parts.push({ weight: 35, value: v });
      if (v >= 0.6) reasons.push('経歴の内容が近い');
    }
    if (target.age && c.age) {
      const diff = Math.abs(target.age - c.age);
      parts.push({ weight: 15, value: Math.max(0, 1 - diff / 10) });
      if (diff <= 3) reasons.push(`年齢が近い（${c.age}歳）`);
    }
    if (target.currentSalary && c.currentSalary) {
      const diff = Math.abs(target.currentSalary - c.currentSalary);
      parts.push({ weight: 10, value: Math.max(0, 1 - diff / 300) });
      if (diff <= 100) reasons.push(`年収帯が近い（${c.currentSalary}万円）`);
    }
    if (tCompany && c.currentCompany) {
      const same = normalizeCompanyName(c.currentCompany) === tCompany;
      parts.push({ weight: 5, value: same ? 1 : 0 });
      if (same) reasons.push('現職が同じ企業');
    }
    const totalWeight = parts.reduce((s, p) => s + p.weight, 0);
    // 比べられる項目が少ないほど偶然の一致になりやすいので、その分を割り引く
    const coverage = Math.min(1, totalWeight / 70);
    const score = totalWeight ? Math.round((parts.reduce((s, p) => s + p.weight * p.value, 0) / totalWeight) * 100 * coverage) : 0;
    return { key: similarKey(entry), score, reason: reasons.join('・') || '一部の項目が近い' };
  });
  return scored
    .filter(s => s.score >= 30)
    .sort((a, b) => b.score - a.score)
    .slice(0, wantedCount);
};

/** 似ている人たちの選考実績を企業ごとに集計する（Geminiを使わず、登録データそのまま）。 */
export const aggregateByCompany = (
  entries: SimilarPoolEntry[],
  normalizeCompanyName: (name: string) => string,
): CompanyStat[] => {
  const byKey = new Map<string, CompanyStat & { nameCounts: Map<string, number> }>();
  entries.forEach(entry => {
    const key = similarKey(entry);
    // 同じ人が同じ企業に複数回応募していても、企業の集計では1名として数える（一番先まで進んだ結果）。
    const bestPerCompany = new Map<string, { name: string; outcome: ApplicationOutcome }>();
    usableApplications(entry.candidate).forEach(app => {
      const norm = normalizeCompanyName(app.companyName) || app.companyName.trim();
      const outcome = classifyApplication(app);
      const prev = bestPerCompany.get(norm);
      const better = !prev
        || (outcome.kind === 'offer' && prev.outcome.kind !== 'offer')
        || stageRank(outcome.reachedStage) > stageRank(prev.outcome.reachedStage);
      if (better) bestPerCompany.set(norm, { name: app.companyName.trim(), outcome });
    });
    bestPerCompany.forEach(({ name, outcome }, norm) => {
      let stat = byKey.get(norm);
      if (!stat) {
        stat = { companyName: name, applicants: 0, passedDocument: 0, offers: 0, declined: 0, withdrawn: 0, inProgress: 0, candidateKeys: [], nameCounts: new Map() };
        byKey.set(norm, stat);
      }
      stat.nameCounts.set(name, (stat.nameCounts.get(name) || 0) + 1);
      stat.applicants++;
      stat.candidateKeys.push(key);
      if (passedDocumentScreening(outcome)) stat.passedDocument++;
      if (outcome.kind === 'offer') stat.offers++;
      else if (outcome.kind === 'declined') stat.declined++;
      else if (outcome.kind === 'withdrawn') stat.withdrawn++;
      else stat.inProgress++;
    });
  });
  return Array.from(byKey.values())
    .map(({ nameCounts, ...stat }) => ({
      ...stat,
      // 表記ゆれがある場合は一番多く使われている表記を代表名にする
      companyName: Array.from(nameCounts.entries()).sort((a, b) => b[1] - a[1])[0][0],
    }))
    .sort((a, b) => b.applicants - a.applicants || b.offers - a.offers || b.passedDocument - a.passedDocument);
};

/** 似ている人たちの選考実績から、傾向とおすすめ／注意企業をGeminiに整理させる。 */
export const analyzeSimilarOutcomes = async (params: {
  target: SimilarSourceCandidate;
  similar: { entry: SimilarPoolEntry; match: SimilarMatch }[];
  companyStats: CompanyStat[];
}): Promise<SimilarAnalysis> => {
  const { target, similar, companyStats } = params;
  const targetProfile = buildProfile(target, 'TARGET', true).replace('[ID: TARGET]\n', '');
  const people = similar.map(({ entry, match }, i) => {
    const c = entry.candidate;
    const apps = usableApplications(c).map(app => {
      const o = classifyApplication(app);
      const memo = redactPersonalInfo((app.memo || '').trim(), c);
      return `  - ${app.companyName.trim()}: ${o.label}${memo ? `（メモ: ${truncate(memo.replace(/\s+/g, ' '), 150)}）` : ''}`;
    });
    return [
      `■類似者${i + 1}（類似度${match.score}: ${match.reason}）`,
      ...basicFacts(c).map(f => `  ${f}`),
      c.summary?.trim() ? `  経歴概要: ${truncate(redactPersonalInfo(c.summary.trim(), c).replace(/\s+/g, ' '), 300)}` : '',
      '  応募企業と結果:',
      ...apps,
    ].filter(Boolean).join('\n');
  });
  const statLines = companyStats.slice(0, 40).map(s =>
    `- ${s.companyName}: 応募${s.applicants}名 / 書類通過${s.passedDocument} / 内定${s.offers} / 見送り${s.declined} / 辞退${s.withdrawn} / 選考中・応募前${s.inProgress}`
  );
  const instruction = [
    'あなたは人材紹介会社のキャリアアドバイザーです。担当中の「対象候補者」の応募先を検討するため、社内の過去データから経歴が似ている人たちの応募先と選考結果を集めました。',
    'これをもとに、対象候補者の応募戦略に役立つ分析を日本語でまとめてください。',
    '',
    '■対象候補者',
    targetProfile,
    '',
    '■経歴が似ている過去の候補者と、その応募企業・選考結果',
    people.join('\n\n'),
    '',
    '■企業別の集計（似ている人たちの中での件数）',
    statLines.join('\n'),
    '',
    '■ルール',
    '- 上記のデータに書かれている事実だけを根拠にする。企業の一般的な評判や推測で企業名を挙げない。',
    '- 件数が少ない（1〜2名）場合はその旨が伝わる書き方にし、断定しすぎない。',
    '- overviewは全体の傾向（どんな企業・業界・ポジションに応募し、どのフェーズで通過／見送りが多いか）を200〜300字で。',
    '- passingFactorsは選考を通過・内定した人に共通する要因、riskFactorsは見送り・辞退につながった要因や注意点を、それぞれ最大4つ（各50字程度）。',
    '- recommendedCompaniesは似た経歴の人が内定・書類通過している等、対象候補者にも応募をおすすめできる企業を最大5社。reasonは根拠となる実績（例: 類似者3名中2名が書類通過）を含めて60字程度で。',
    '- cautionCompaniesは似た経歴の人が見送り・辞退になっている等、応募時に注意が要る企業を最大5社。reasonは何フェーズで止まったか等の根拠を含めて60字程度で。該当が無ければ空配列。',
    '- adviceは対象候補者の応募戦略・面接対策についての具体的な提案を150字程度で。',
    '- 氏名は書かず、「類似者1」のように番号で参照してよい。',
  ].join('\n');
  const response = await newAi().models.generateContent({
    model: MODEL,
    contents: instruction,
    config: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          overview: { type: Type.STRING },
          passingFactors: { type: Type.ARRAY, items: { type: Type.STRING } },
          riskFactors: { type: Type.ARRAY, items: { type: Type.STRING } },
          recommendedCompanies: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: { companyName: { type: Type.STRING }, reason: { type: Type.STRING } },
              required: ['companyName', 'reason'],
            },
          },
          cautionCompanies: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: { companyName: { type: Type.STRING }, reason: { type: Type.STRING } },
              required: ['companyName', 'reason'],
            },
          },
          advice: { type: Type.STRING },
        },
        required: ['overview', 'passingFactors', 'riskFactors', 'recommendedCompanies', 'cautionCompanies', 'advice'],
      },
    },
  });
  const parsed = JSON.parse((response.text || '{}').trim());
  const s = (v: unknown) => (typeof v === 'string' ? v.trim() : '');
  const list = (v: unknown) => (Array.isArray(v) ? v.map(s).filter(Boolean) : []);
  const companies = (v: unknown) => (Array.isArray(v) ? v : [])
    .map((x: any) => ({ companyName: s(x?.companyName), reason: s(x?.reason) }))
    .filter(x => x.companyName);
  return {
    overview: s(parsed.overview),
    passingFactors: list(parsed.passingFactors),
    riskFactors: list(parsed.riskFactors),
    recommendedCompanies: companies(parsed.recommendedCompanies),
    cautionCompanies: companies(parsed.cautionCompanies),
    advice: s(parsed.advice),
  };
};
