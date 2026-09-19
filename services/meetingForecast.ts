// BP会（チームリーダーが集まる週次の会議）向けの「会議用レポート」の計算・テキスト生成ロジック。
//
// 候補者パイプラインの「チーム」タブで、各メンバーが自分のパイプラインに入力している意思決定時期・
// 確度（内定確度/入社確度）とは別に、会議用の「決定見込み時期（月＋週）」と「確度（S/A/B/C/D）」を
// 上書き設定できるようにし、その値で確度加重後の売上・粗利を算出して会議資料に貼れるテキストに
// する。メンバー入力側の確度（A/B+/B/B-/C）とは定義が異なる（会議側は決定済=S=100%など、
// 歩留まりを掛け合わせる前提のグレード）ため、メンバー入力からの自動変換は行わない。
//
// index.tsxのCandidate/CompanyApplication型に依存しないよう、入力は構造的な型（MeetingForecastRowInput）
// に落としてから受け取る純粋関数だけにしている（UI側でCandidate等から組み立てる）。

export const MEETING_CONFIDENCE_GRADES = ['S', 'A', 'B', 'C', 'D'] as const;
export type MeetingConfidence = typeof MEETING_CONFIDENCE_GRADES[number];

// 会議資料の定義（S=決定済/100%、A=70%、B=40%、C=20%）。Dは資料に定義は無いが、実際の資料で
// 「ほぼ決まらない」案件に使われているため「見込み外」（0%固定・加重前の合計にも含めない）として
// 用意している。A/B/Cの歩留まりは設定で変更できる（S=100%・D=0%は固定）。
export const DEFAULT_MEETING_CONFIDENCE_RATES: Record<MeetingConfidence, number> = {
  S: 100,
  A: 70,
  B: 40,
  C: 20,
  D: 0,
};

/** 決定見込み時期の上書き値のうち、「どの月の会議資料にも載せない」を表す特別な月キー。 */
export const MEETING_MONTH_NONE = 'none';

export interface MeetingTiming {
  month: string; // yyyy-MM
  week: number | null; // その月の第N週（1始まり）。null = 週は未定
}

export interface MeetingForecastOverride {
  // yyyy-MM、またはMEETING_MONTH_NONE。未設定 = メンバー入力の意思決定時期に従う。
  decisionMonth?: string;
  // 1〜6。0/未設定 = 週未定（decisionMonthが月の場合のみ意味を持つ）。
  decisionWeek?: number;
  // 未設定 = 内定承諾済みならS、それ以外は未設定のまま（集計対象外）。
  confidence?: MeetingConfidence;
}

export interface MeetingForecastSettings {
  // キーは buildMeetingRowKey の結果（登録者メール::候補者ID::選考ID）。
  overrides?: Record<string, MeetingForecastOverride>;
  // A/B/Cの歩留まり（%）。未設定の値はDEFAULT_MEETING_CONFIDENCE_RATESを使う。
  rates?: Partial<Record<MeetingConfidence, number>>;
  // 「当月の目標本数」。キーは `${teamId}:${yyyy-MM}`。
  targetCounts?: Record<string, number>;
}

export const buildMeetingRowKey = (ownerEmail: string, candidateId: string, applicationId: string): string =>
  `${ownerEmail.trim().toLowerCase()}::${candidateId}::${applicationId}`;

export const buildMeetingTargetKey = (teamId: string, month: string): string => `${teamId}:${month}`;

export const resolveMeetingRates = (
  custom: MeetingForecastSettings['rates'] | undefined
): Record<MeetingConfidence, number> => {
  const resolved = { ...DEFAULT_MEETING_CONFIDENCE_RATES };
  MEETING_CONFIDENCE_GRADES.forEach(grade => {
    const value = custom?.[grade];
    if ((grade === 'A' || grade === 'B' || grade === 'C') && typeof value === 'number' && Number.isFinite(value)) {
      resolved[grade] = Math.min(100, Math.max(0, value));
    }
  });
  return resolved;
};

// --- 週の算出 ---------------------------------------------------------------------------------
// 「9月3週」= 9月1日を含む週を第1週として数える（週の始まりはアプリの「週の始まり」設定に従う）。
// 月をまたぐ週は日付ごとにその日の属する月へ振り分ける（例: 10/1が含まれる週は9月の週ではなく
// 10月第1週として扱う）。DSTの影響を避けるため日数計算はUTC基準で行う。

const DAY_MS = 24 * 60 * 60 * 1000;

const parseISODate = (iso: string): { y: number; m: number; d: number } => {
  const [y, m, d] = iso.split('-').map(Number);
  return { y, m, d };
};

const startOfWeekUTC = (utcMs: number, weekStartsOn: 0 | 6): number => {
  const dayOfWeek = new Date(utcMs).getUTCDay();
  return utcMs - ((dayOfWeek - weekStartsOn + 7) % 7) * DAY_MS;
};

/** yyyy-mm-dd → その日が属する月と、月内の第N週。 */
export function getWeekOfMonth(dateISO: string, weekStartsOn: 0 | 6 = 0): MeetingTiming {
  const { y, m, d } = parseISODate(dateISO);
  const target = Date.UTC(y, m - 1, d);
  const firstOfMonth = Date.UTC(y, m - 1, 1);
  const week = Math.round((startOfWeekUTC(target, weekStartsOn) - startOfWeekUTC(firstOfMonth, weekStartsOn)) / (7 * DAY_MS)) + 1;
  return { month: `${y}-${String(m).padStart(2, '0')}`, week };
}

/** yyyy-MM の月が、週の始まり設定のもとで何週まであるか（5〜6）。 */
export function getWeekCountOfMonth(month: string, weekStartsOn: 0 | 6 = 0): number {
  const [y, m] = month.split('-').map(Number);
  const lastDay = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return getWeekOfMonth(`${month}-${String(lastDay).padStart(2, '0')}`, weekStartsOn).week!;
}

/** 第N週の日付範囲ラベル（月の範囲内に丸める）。例: "9/28〜9/30"。 */
export function getWeekRangeLabel(month: string, week: number, weekStartsOn: 0 | 6 = 0): string {
  const [y, m] = month.split('-').map(Number);
  const firstOfMonth = Date.UTC(y, m - 1, 1);
  const lastOfMonth = Date.UTC(y, m, 0);
  const weekStart = startOfWeekUTC(firstOfMonth, weekStartsOn) + (week - 1) * 7 * DAY_MS;
  const start = Math.max(weekStart, firstOfMonth);
  const end = Math.min(weekStart + 6 * DAY_MS, lastOfMonth);
  const fmt = (ms: number) => `${new Date(ms).getUTCMonth() + 1}/${new Date(ms).getUTCDate()}`;
  return `${fmt(start)}〜${fmt(end)}`;
}

export const formatMonthLabel = (month: string): string => `${Number(month.split('-')[1])}月`;

export const formatTimingLabel = (timing: MeetingTiming): string =>
  timing.week ? `${formatMonthLabel(timing.month)}${timing.week}週` : `${formatMonthLabel(timing.month)}（週未定）`;

// --- 行の解決 ---------------------------------------------------------------------------------

export interface MeetingForecastRowInput {
  key: string;
  // 同じ求職者の複数選考が同時に載っていないか（二重計上）の判定用。
  candidateKey: string;
  candidateName: string;
  companyName: string;
  caLabel: string;
  // メンバーが入力している意思決定時期から導いた月・週。どちらも未入力ならnull。
  memberTiming: MeetingTiming | null;
  // この選考が内定承諾まで進んでいるか（決定済 = 確度Sの既定値）。
  isAccepted: boolean;
  // 想定紹介料・想定粗利（万円）。算出に必要な入力が足りない場合はnull。
  revenue: number | null;
  profit: number | null;
}

export interface MeetingForecastRow extends MeetingForecastRowInput {
  timing: MeetingTiming | null;
  isTimingOverridden: boolean;
  confidence: MeetingConfidence | null;
  isConfidenceOverridden: boolean;
  weightedRevenue: number | null;
  weightedProfit: number | null;
}

export function resolveMeetingRow(
  input: MeetingForecastRowInput,
  override: MeetingForecastOverride | undefined,
  rates: Record<MeetingConfidence, number>
): MeetingForecastRow {
  let timing = input.memberTiming;
  let isTimingOverridden = false;
  if (override?.decisionMonth) {
    isTimingOverridden = true;
    timing = override.decisionMonth === MEETING_MONTH_NONE
      ? null
      : { month: override.decisionMonth, week: override.decisionWeek || null };
  }
  const confidence: MeetingConfidence | null = override?.confidence ?? (input.isAccepted ? 'S' : null);
  const rate = confidence ? rates[confidence] : null;
  return {
    ...input,
    timing,
    isTimingOverridden,
    confidence,
    isConfidenceOverridden: !!override?.confidence,
    weightedRevenue: rate !== null && input.revenue !== null ? input.revenue * rate / 100 : null,
    weightedProfit: rate !== null && input.profit !== null ? input.profit * rate / 100 : null,
  };
}

export interface MeetingForecastSummary {
  rowCount: number;
  decidedCount: number; // 確度S
  pendingCount: number; // 確度A/B/C
  excludedCount: number; // 確度D（見込み外。一覧には載せるが本数・金額の集計には含めない）
  unratedCount: number; // 確度未設定（集計に含まれない）
  unestimableCount: number; // 確度は設定済みだが売上・粗利が算出できない（集計に含まれない）
  totalRevenue: number; // 確度加重前
  totalProfit: number;
  weightedRevenue: number; // 確度加重後
  weightedProfit: number;
}

export const isRowInMonth = (row: MeetingForecastRow, month: string): boolean => row.timing?.month === month;

export function summarizeMeetingRows(rows: MeetingForecastRow[]): MeetingForecastSummary {
  const summary: MeetingForecastSummary = {
    rowCount: rows.length,
    decidedCount: 0,
    pendingCount: 0,
    excludedCount: 0,
    unratedCount: 0,
    unestimableCount: 0,
    totalRevenue: 0,
    totalProfit: 0,
    weightedRevenue: 0,
    weightedProfit: 0,
  };
  rows.forEach(row => {
    if (!row.confidence) {
      summary.unratedCount++;
      return;
    }
    if (row.confidence === 'D') {
      summary.excludedCount++;
      return;
    }
    if (row.confidence === 'S') summary.decidedCount++;
    else summary.pendingCount++;
    if (row.revenue === null || row.profit === null) {
      summary.unestimableCount++;
      return;
    }
    summary.totalRevenue += row.revenue;
    summary.totalProfit += row.profit;
    summary.weightedRevenue += row.weightedRevenue ?? 0;
    summary.weightedProfit += row.weightedProfit ?? 0;
  });
  return summary;
}

/** 同じ求職者の複数の選考が、対象行に重複して載っているものの候補者キー一覧（二重計上の注意喚起用）。 */
export function findMultiApplicationCandidateKeys(rows: MeetingForecastRow[]): Set<string> {
  const counts = new Map<string, number>();
  rows.forEach(row => counts.set(row.candidateKey, (counts.get(row.candidateKey) || 0) + 1));
  return new Set(Array.from(counts.entries()).filter(([, n]) => n > 1).map(([key]) => key));
}

// --- テキスト生成 -----------------------------------------------------------------------------

export interface MeetingForecastTextOptions {
  month: string;
  rates: Record<MeetingConfidence, number>;
  targetCount?: number | null;
  surnameOnly: boolean;
  includeAmounts: boolean;
}

const formatManYen = (n: number): string => `${Math.round(n).toLocaleString('ja-JP')}万円`;

const toSurname = (name: string): string => name.trim().split(/[\s　]+/)[0] || name;

const CONFIDENCE_ORDER: Record<MeetingConfidence, number> = { S: 0, A: 1, B: 2, C: 3, D: 4 };

/**
 * 会議資料（BP会のGoogleドキュメント）の「＜数字＞」ブロックと同じ体裁のテキストを作る。
 * rowsには対象月に決める行（isRowInMonthで絞り込み済み）だけを渡す。
 */
export function buildMeetingForecastText(rows: MeetingForecastRow[], options: MeetingForecastTextOptions): string {
  const { month, rates, targetCount, surnameOnly, includeAmounts } = options;
  const summary = summarizeMeetingRows(rows);
  const lines: string[] = [];

  lines.push('＜数字＞');
  lines.push('＜サマリ＞');
  const countDetail = `決定済：${summary.decidedCount}本 ／ 残見込み：${summary.pendingCount}本`;
  lines.push(
    targetCount !== null && targetCount !== undefined
      ? `・当月の目標本数：${targetCount}本（${countDetail}）`
      : `・当月の本数：${countDetail}`
  );
  lines.push(`・当月の売上予測：${formatManYen(summary.totalRevenue)}`);
  lines.push(`・当月の粗利予測：${formatManYen(summary.totalProfit)}`);
  const rateLegend = (['A', 'B', 'C'] as const).map(grade => `${grade}=${rates[grade]}%`).join(' / ');
  lines.push(`＜確度加重後＞（${rateLegend}）`);
  lines.push(`・総売上：${formatManYen(summary.weightedRevenue)}`);
  lines.push(`・総粗利：${formatManYen(summary.weightedProfit)}`);
  lines.push('');
  lines.push('＜詳細＞');

  // 週ごと（週未定は最後）にまとめ、週の中は確度の高い順→名前順。
  const groups = new Map<number, MeetingForecastRow[]>();
  rows.forEach(row => {
    const week = row.timing?.week ?? 0;
    const list = groups.get(week) || [];
    list.push(row);
    groups.set(week, list);
  });
  const orderedWeeks = Array.from(groups.keys()).sort((a, b) => (a === 0 ? 1 : b === 0 ? -1 : a - b));
  orderedWeeks.forEach((week, index) => {
    if (index > 0) lines.push('');
    lines.push(week === 0 ? `・${formatMonthLabel(month)}中（週未定）に決める人` : `・${formatMonthLabel(month)}${week}週に決める人`);
    const sorted = [...groups.get(week)!].sort((a, b) => {
      const rankA = a.confidence ? CONFIDENCE_ORDER[a.confidence] : 99;
      const rankB = b.confidence ? CONFIDENCE_ORDER[b.confidence] : 99;
      return rankA - rankB || a.candidateName.localeCompare(b.candidateName, 'ja');
    });
    sorted.forEach(row => {
      const name = surnameOnly ? toSurname(row.candidateName) : row.candidateName;
      const ca = surnameOnly ? toSurname(row.caLabel) : row.caLabel;
      let line = `　・${name}様／確度：${row.confidence ?? '未設定'}（決定先：${row.companyName || '未入力'}／CA：${ca}）`;
      if (includeAmounts) {
        line += row.revenue !== null && row.profit !== null
          ? `／売上：${formatManYen(row.revenue)}／粗利：${formatManYen(row.profit)}`
          : '／売上・粗利：未算出';
      }
      lines.push(line);
    });
  });
  if (orderedWeeks.length === 0) lines.push('（この月に決める人はいません）');

  return lines.join('\n');
}
