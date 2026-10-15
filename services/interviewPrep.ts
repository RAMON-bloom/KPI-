// 面接対策ジェネレーター（bloom社内のGoogle Apps Script Webアプリ）との連携。
//
// 2026-10時点でGAS側はまだ連携用の受け口を持っていないため、ここにあるのは「kpi-mgr側だけで
// 先に用意できる部分」:
// - 生成画面を開くURL（企業名・フェーズ・候補者ID・応募IDをクエリに載せる。GAS側が読まなくても
//   今まで通り普通に開くだけなので無害。氏名やメモなど個人情報はURLに載せない）
// - 選考結果イベント（通過／お見送り／辞退）の組み立てと送信。INTERVIEW_PREP_RESULT_ENDPOINT
//   が空の間は何も送らない — GAS側に受け口（doPost）ができたらそのURLを入れるだけで動く。

export const INTERVIEW_PREP_GENERATOR_URL =
  'https://script.google.com/a/macros/bloom-firm.com/s/AKfycbwnKXAiiHJXLZ9jWOg6lbtN4saG0IqV-4c9b6oEew6fFul393TJhjUgFQlijiUZQwZQhg/exec';

// GAS側の結果受け口（doPost）のURL。開発者との打ち合わせで決まるまでは空のまま＝送信しない。
export const INTERVIEW_PREP_RESULT_ENDPOINT = '';

// kpi-mgrの選考ステージ → 面接対策ジェネレーター側のフェーズ表記。対策資料の対象にならない
// ステージ（打診・内定など）はnull。
const GAS_PHASE_BY_STAGE: Record<string, string> = {
  'カジュアル面談': 'カジュアル面談',
  '1次面接': '一次面接',
  '2次面接': '二次面接',
  '最終面接': '最終面接',
  'オファー面談': 'オファー面談',
};

export const toGeneratorPhase = (stage: string): string | null => GAS_PHASE_BY_STAGE[stage] ?? null;

/** 面接対策資料を作れる（＝面接系の）ステージかどうか。 */
export const isInterviewPrepStage = (stage: string): boolean => toGeneratorPhase(stage) !== null;

/**
 * opts.handoff を付けると、kpi_handoffs シートに置いた面談メモの受け渡し番号と auto=1 を付ける
 * （ジェネレーター側が対応すれば、メモを読み込んで自動で生成を始める）。
 */
export const buildGeneratorUrl = (
  candidateId: string,
  app: { id: string; companyName: string; stage: string; gasCompanyId?: string },
  opts?: { handoff?: string },
): string => {
  const params = new URLSearchParams({ from: 'kpi-mgr', candidateId, applicationId: app.id });
  if (app.companyName.trim()) params.set('company', app.companyName.trim());
  if (app.gasCompanyId) params.set('companyId', app.gasCompanyId);
  const phase = toGeneratorPhase(app.stage);
  if (phase) params.set('phase', phase);
  if (opts?.handoff) {
    params.set('handoff', opts.handoff);
    params.set('auto', '1');
  }
  return `${INTERVIEW_PREP_GENERATOR_URL}?${params.toString()}`;
};

/** 生成画面の「面談メモ」欄に貼り付ける用のテキスト（取り込み済みの面談ログ・手書きメモを時系列で連結）。 */
export const buildInterviewMemoText = (memos: { title: string; content: string; logDate?: string; updatedAt: string }[]): string =>
  memos
    .filter(m => m.content.trim())
    .map(m => `【${m.title || 'メモ'}】${m.logDate ? `（${m.logDate}）` : ''}\n${m.content.trim()}`)
    .join('\n\n');

export type InterviewPrepResult = '通過' | 'お見送り' | '辞退';

export interface InterviewPrepResultEvent {
  source: 'kpi-mgr';
  candidateId: string;
  applicationId: string;
  companyName: string;
  phase: string; // 面接対策ジェネレーター側のフェーズ表記（一次面接など）
  kpiStage: string; // 結果が出たkpi-mgr側のステージ（1次面接など）
  nextKpiStage: string; // 変更後のステージ
  result: InterviewPrepResult;
  decidedAt: string; // yyyy-mm-dd
  ownerEmail: string; // 担当者（kpi-mgrにログイン中のユーザー）
}

const FORWARD_STAGES = ['打診', '書類選考', '適性検査', 'カジュアル面談', '1次面接', '2次面接', '最終面接', 'オファー面談', '内定', '内定承諾'];

const classifyTransition = (from: string, to: string): InterviewPrepResult | null => {
  if (to === 'お見送り') return 'お見送り';
  if (to === '選考辞退' || to === '内定承諾後辞退') return '辞退';
  const fromIdx = FORWARD_STAGES.indexOf(from);
  const toIdx = FORWARD_STAGES.indexOf(to);
  if (fromIdx >= 0 && toIdx > fromIdx) return '通過';
  return null;
};

/**
 * 保存前後の候補者を比べて、面接系ステージの結果が確定した応募ごとにイベントを作る。
 * 例: 1次面接→2次面接 = 一次面接「通過」、1次面接→お見送り = 一次面接「お見送り」。
 * 変更前が面接系ステージでない応募（打診→書類選考、内定承諾→内定承諾後辞退など）は対象外。
 */
export const computeResultEvents = (
  prev: { id: string; applications: { id: string; companyName: string; stage: string }[] } | undefined,
  next: { id: string; applications: { id: string; companyName: string; stage: string }[] },
  ownerEmail: string,
  todayStr: string,
): InterviewPrepResultEvent[] => {
  if (!prev) return [];
  const prevById = new Map(prev.applications.map(a => [a.id, a]));
  const events: InterviewPrepResultEvent[] = [];
  for (const app of next.applications) {
    const before = prevById.get(app.id);
    if (!before || before.stage === app.stage) continue;
    const phase = toGeneratorPhase(before.stage);
    if (!phase) continue;
    const result = classifyTransition(before.stage, app.stage);
    if (!result) continue;
    events.push({
      source: 'kpi-mgr',
      candidateId: next.id,
      applicationId: app.id,
      companyName: app.companyName,
      phase,
      kpiStage: before.stage,
      nextKpiStage: app.stage,
      result,
      decidedAt: todayStr,
      ownerEmail,
    });
  }
  return events;
};

/**
 * 結果イベントをGASへ送る。受け口URLが未設定なら何もしない。失敗しても保存処理は止めない
 * （コンソールに残すだけ）。Content-Typeをtext/plainにしているのは、GASのWebアプリが
 * CORSのプリフライト（OPTIONS）に応答できないため。
 */
export const sendResultEvents = async (events: InterviewPrepResultEvent[]): Promise<void> => {
  if (!INTERVIEW_PREP_RESULT_ENDPOINT || events.length === 0) return;
  try {
    await fetch(INTERVIEW_PREP_RESULT_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ type: 'interviewResults', events }),
    });
  } catch (err) {
    console.warn('面接対策ジェネレーターへの選考結果の送信に失敗しました', err);
  }
};
