
import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import ReactDOM from 'react-dom/client';
import { GoogleGenAI, Type } from "@google/genai";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  BarElement,
  Title,
  Tooltip,
  Legend,
  Filler, // Import Filler for area charts
} from 'chart.js';
import { Line, Bar } from 'react-chartjs-2';
import { signIn, signOut, getCurrentSession, getLastKnownEmail, reauthorizeWithConsent, GoogleIdentity } from './services/googleAuth';
import { loadOwnData, saveOwnDataDebounced, flushPendingSave, forceSyncNow, hasPendingSync, retryPendingSyncIfNeeded, onSyncStatusChange, getLastSyncedAt, readLegacyAppData, loadAllTeammatesData, loadTeamsConfig, saveTeamsConfig, readLocalCache, loadMediaConfig, saveMediaConfig, readMediaConfigCache } from './services/dataSync';
import { searchInterviewLogsByName, exportGoogleDocAsText, InterviewLogFile } from './services/googleDrive';
import { fetchScoutReplyCounts, fetchScoutReplyCountsForRange, GmailPermissionError, ScoutReplyRangeResult } from './services/gmailScout';
import { decodeCsvFile, parseScoutCsv, ScoutCsvMediaId, ScoutCsvDayCounts, ScoutCsvParseResult } from './services/mediaCsvImport';
import { createPipelineTask, updatePipelineTask, deletePipelineTask, GoogleTasksPermissionError } from './services/googleTasks';

ChartJS.register(
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  BarElement,
  Title,
  Tooltip,
  Legend,
  Filler // Register Filler
);

// Media (scouting source) management is restricted to this single account, regardless of who
// happens to own the underlying Drive file — everyone else gets read-only access.
const MEDIA_ADMIN_EMAIL = 'gou.higashibara@bloom-firm.com';
// The one account that can always create/edit teams and grant that ability to others (see
// TeamsModal's permission section). Others gain edit access only via teamsAuthorizedEditors,
// a list this admin manages, persisted alongside the teams themselves.
const TEAMS_ADMIN_EMAIL = 'gou.higashibara@bloom-firm.com';

const GENERAL_KPIS = {
  candidatesSubmitted: { label: '候補者推薦数', target: 30 },
  documentScreeningPassed: { label: '書類選考通過数', target: 25 },
  firstInterviewPassed: { label: '1次面接通過数', target: 15 },
  secondInterviewPassed: { label: '2次面接通過数', target: 10 },
  finalInterviewPassed: { label: '最終面接合格数', target: 7 },
  offersExtended: { label: '内定数', target: 5 },
  placements: { label: '内定承諾数', target: 3 },
};

// A scouting media source (e.g. "RDS", "Doda"). User-editable and persisted to Google
// Drive (see services/dataSync.ts loadMediaConfig/saveMediaConfig). `id` is generated once
// and never changes, so renaming a media never breaks its historical KPI data; `name` is the
// editable display label.
interface MediaEntry {
  id: string;
  name: string;
  isArchived: boolean;
  createdAt: string;
  // Handling fee this media source charges (as a % of the placed candidate's expected annual
  // salary) when a candidate sourced through it gets placed — subtracted from the client fee
  // to compute expected gross profit.
  feeRate?: number;
}

interface MediaConfig {
  schemaVersion: number;
  media: MediaEntry[];
}

// Seeded once into the shared Drive media-config file the first time it's created, so
// existing installs keep working with zero migration: these ids match the exact lowercase
// prefixes ("rds_scoutsSent" etc.) that historical KPI data already uses.
const SEED_MEDIA: MediaEntry[] = [
  { id: 'rds', name: 'RDS', isArchived: false, createdAt: '2024-01-01T00:00:00.000Z' },
  { id: 'doda', name: 'Doda', isArchived: false, createdAt: '2024-01-01T00:00:00.000Z' },
  { id: 'liiga', name: 'Liiga', isArchived: false, createdAt: '2024-01-01T00:00:00.000Z' },
  { id: 'biz', name: 'BIZ', isArchived: false, createdAt: '2024-01-01T00:00:00.000Z' },
  { id: 'linkedin', name: 'Linkedin', isArchived: false, createdAt: '2024-01-01T00:00:00.000Z' },
  { id: 'ambi', name: 'AMBI', isArchived: false, createdAt: '2024-01-01T00:00:00.000Z' },
  { id: 'green', name: 'Green', isArchived: false, createdAt: '2024-01-01T00:00:00.000Z' },
];

const MEDIA_KPI_SUFFIXES = [
  'scoutsSent', 'scoutReplies', 'effectiveReplies',
  'documentsCollected', 'effectiveDocumentsCollected',
  'initialInterviews', 'effectiveInitialInterviews',
] as const;

type KpiKey = string;

/** Every KPI key that exists for a given media list: general KPIs plus 7 per media source. */
const buildAllKpiKeys = (media: MediaEntry[]): KpiKey[] => [
  ...Object.keys(GENERAL_KPIS),
  ...media.flatMap(m => MEDIA_KPI_SUFFIXES.map(suffix => `${m.id}_${suffix}`)),
];

const buildDefaultKpiTargets = (media: MediaEntry[]): Record<KpiKey, number> => {
    return buildAllKpiKeys(media).reduce((acc, key) => {
        if (key in GENERAL_KPIS) {
            acc[key] = GENERAL_KPIS[key as keyof typeof GENERAL_KPIS].target;
        } else if (key.endsWith('_scoutsSent')) {
            acc[key] = 200;
        } else if (key.endsWith('_scoutReplies')) {
            acc[key] = 20;
        } else if (key.endsWith('_effectiveReplies')) {
            acc[key] = 5;
        } else if (key.endsWith('_documentsCollected')) {
            acc[key] = 6;
        } else if (key.endsWith('_effectiveDocumentsCollected')) {
            acc[key] = 4;
        } else if (key.endsWith('_initialInterviews')) {
            acc[key] = 8;
        } else if (key.endsWith('_effectiveInitialInterviews')) {
            acc[key] = 7;
        }
        return acc;
    }, {} as Record<KpiKey, number>);
};


interface KpiEntry {
  id: number;
  date: string;
  values: { [key in KpiKey]: number };
}

type KpiTotals = { [key in KpiKey]: number };

const calculateMonthlyTotals = (entries: KpiEntry[], allMedia: MediaEntry[]): KpiTotals => {
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    const allKeys = buildAllKpiKeys(allMedia);

    const totals = allKeys.reduce((acc, key) => {
      acc[key] = 0;
      return acc;
    }, {} as KpiTotals);

    entries.forEach(entry => {
      const entryDate = new Date(entry.date);
      if (entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear) {
        allKeys.forEach(key => {
          totals[key] += entry.values[key] || 0;
        });
      }
    });

    return totals;
};

/** Same shape as calculateMonthlyTotals, but for an arbitrary inclusive [startDate, endDate] range. */
const calculateTotalsForRange = (entries: KpiEntry[], allMedia: MediaEntry[], startDate: Date, endDate: Date): KpiTotals => {
    const allKeys = buildAllKpiKeys(allMedia);
    const totals = allKeys.reduce((acc, key) => {
      acc[key] = 0;
      return acc;
    }, {} as KpiTotals);

    const start = startDate.getTime();
    const end = endDate.getTime();

    entries.forEach(entry => {
      const entryTime = new Date(entry.date).getTime();
      if (entryTime >= start && entryTime <= end) {
        allKeys.forEach(key => {
          totals[key] += entry.values[key] || 0;
        });
      }
    });

    return totals;
};


// Types for Weekly Summary
interface WeeklyMediaStats {
  source: string; // display name
  id: string; // stable media id, used to look up KPI keys
  scoutsSent: number;
  scoutReplies: number;
  effectiveReplies: number;
  documentsCollected: number;
  effectiveDocumentsCollected: number;
  initialInterviews: number;
  effectiveInitialInterviews: number;
}

interface WeeklyData {
  mediaStats: WeeklyMediaStats[];
  totalCandidatesSubmitted: number;
  totalInitialInterviews: number;
}


// --- Candidate Pipeline Types ---
const PIPELINE_STAGES = ['打診', '書類選考', '適性検査', 'カジュアル面談', '1次面接', '2次面接', '最終面接', '内定', '内定承諾', 'お見送り', '選考辞退'] as const;
type PipelineStage = typeof PIPELINE_STAGES[number];

const CONFIDENCE_GRADES = ['A', 'B+', 'B', 'B-', 'C'] as const;
type ConfidenceGrade = typeof CONFIDENCE_GRADES[number];
// Lower is better; used to rank applications when picking the single most-likely-to-close one
// per candidate. Missing ratings are treated as worse than any explicit grade (see
// pickBestApplicationPerCandidate) rather than defaulting to a specific grade.
const CONFIDENCE_RANK: Record<ConfidenceGrade, number> = { 'A': 0, 'B+': 1, 'B': 2, 'B-': 3, 'C': 4 };

const STAGE_COLOR_MAP: Record<PipelineStage, string> = {
    '打診': 'grey',
    '書類選考': 'cadetblue',
    '適性検査': 'mediumpurple',
    'カジュアル面談': 'lightblue',
    '1次面接': 'dodgerblue',
    '2次面接': 'royalblue',
    '最終面接': 'mediumblue',
    '内定': 'orange',
    '内定承諾': 'limegreen',
    'お見送り': 'crimson',
    '選考辞退': 'salmon',
};

/** Short labels for the pipeline calendar, where space is tight — e.g. "1次面接" → "1次". */
const STAGE_SHORT_LABELS: Record<PipelineStage, string> = {
    '打診': '打診',
    '書類選考': '書類',
    '適性検査': '適性',
    'カジュアル面談': 'カジュアル',
    '1次面接': '1次',
    '2次面接': '2次',
    '最終面接': '最終',
    '内定': 'オファー',
    '内定承諾': '承諾',
    'お見送り': '見送り',
    '選考辞退': '辞退',
};

const COMPANY_NAME_DESIGNATORS = [
    '株式会社', '有限会社', '合同会社', '合名会社', '合資会社',
    '一般社団法人', '一般財団法人', '公益社団法人', '公益財団法人',
    '特定非営利活動法人', 'NPO法人',
];

/**
 * Collapses common company-name spelling variance (株式会社 prefix vs. suffix vs. omitted,
 * full-width/half-width characters, extra whitespace, "(株)" abbreviations, English Inc./Co.,
 * Ltd. suffixes, case) into a single grouping key, so the company-pipeline view doesn't
 * silently split one company into several just because two people typed its name slightly
 * differently — without requiring anyone to standardize their input.
 */
function normalizeCompanyName(name: string): string {
    if (!name) return '';
    let s = name.trim();
    // Full-width alphanumerics/symbols -> half-width
    s = s.replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
    s = s.replace(/　/g, ' ');
    COMPANY_NAME_DESIGNATORS.forEach(d => { s = s.split(d).join(''); });
    s = s.replace(/[（(]\s*(株|有)\s*[）)]/g, '');
    s = s.toLowerCase();
    s = s.replace(/\b(inc|ltd|co|corp|corporation|company|llc|kk)\b\.?/g, '');
    s = s.replace(/[\s.,、。・\-ー_/／]/g, '');
    return s;
}

interface CompanyApplication {
  id: string;
  companyName: string;
  stage: PipelineStage;
  nextAction: string;
  // Manually-set date of the next scheduled interview/action for this application, shown on
  // the pipeline calendar. Free-form text notes about the current status, separate from
  // nextAction (which is more of a short "what's next" label).
  scheduledDate?: string; // ISO yyyy-mm-dd
  // Optional start time for scheduledDate, e.g. "13:30" — shown as a "13:30 " prefix on the
  // Google Tasks title (see buildPipelineTaskContent) since Tasks itself has no time-of-day
  // field, only a due date.
  scheduledTime?: string; // HH:mm, 24-hour
  // When a final decision (内定/内定承諾, occasionally お見送り) is expected to be reached for
  // this application — distinct from scheduledDate, which is the next scheduled
  // interview/action, not the eventual outcome date.
  expectedDecisionDate?: string; // ISO yyyy-mm-dd
  memo?: string;
  isHidden?: boolean;
  // Referral fee rate charged to this client company (% of the candidate's expected annual
  // salary) — position-specific, since different companies/positions can negotiate different
  // rates for the same candidate.
  feeRate?: number;
  // Likelihood ratings for this specific application: offerConfidence = chance of receiving
  // an offer at all, acceptanceConfidence = chance the candidate accepts it if offered. Used
  // to pick the single most-likely-to-close application per candidate for the gross-profit
  // dashboard, so a candidate interviewing at several companies isn't counted multiple times.
  offerConfidence?: ConfidenceGrade;
  acceptanceConfidence?: ConfidenceGrade;
}

interface Candidate {
  id: string;
  name: string;
  salary: number; // in JPY万
  currentSalary: number; // in JPY万
  currentCompany: string;
  education: string;
  source: string; // a MediaEntry.id, 'Other', or ''
  usingOtherAgents: boolean;
  applications: CompanyApplication[];
  summary: string;
  resumeFile?: { name: string; }; // for backward compatibility
  resumeFiles?: { name: string; }[];
  interviewAudioFile?: { name: string; } | null;
  interviewSummary?: string;
  createdAt: string; // ISO string
  isHidden?: boolean;
  // 掘り起しリスト: presence means this candidate is parked for future re-engagement rather
  // than actively pursued right now. Adding a candidate here also sets isHidden — it's a
  // candidate-level concept (not tied to any one CompanyApplication), since the point is
  // reconsidering the person as a whole, not a specific past application. nextActionDate drives
  // a reminder event on the pipeline calendar (see PipelineCalendarView).
  revival?: { nextAction: string; nextActionDate: string /* ISO yyyy-mm-dd */ };
  // Expected annual salary (万円) this candidate is likely to be placed at — distinct from
  // `salary` (their stated desired salary) — used with each application's feeRate to project
  // gross profit.
  expectedAnnualSalary?: number;
  // Only set when viewing the all-users/team-aggregated pipeline (never persisted) —
  // identifies whose data this candidate belongs to, for display and edit-permission checks.
  ownerEmail?: string;
  ownerLabel?: string;
}


// --- Multi-user data structures ---
interface UserData {
  entries: KpiEntry[];
  kpiTargets: Record<KpiKey, number>;
  weeklyKpiTargets: Record<KpiKey, number>;
  dailyKpiTargets: Record<KpiKey, number>;
  candidates: Candidate[];
  displayName?: string;
  // Automatically mirrors this user's own パイプラインカレンダー entries (applications'
  // scheduledDate) into their own Google Tasks default list on every save. Tracks the Google-
  // assigned task ID per CompanyApplication.id so edits update the same task instead of
  // creating duplicates — kept as a separate map (not a field on CompanyApplication itself) so
  // it survives untouched even through form code that rebuilds application objects field-by-
  // field and would otherwise silently drop an unrecognized property.
  googleTaskIdsByApplicationId?: Record<string, string>;
}

const normalizeEmail = (email: string): string => email.trim().toLowerCase();

/**
 * Case-insensitively resolves a Team's free-typed memberEmails entry against allUsersData's
 * keys (which are always exactly the casing Google returned at sign-in). A member email typed
 * or pasted with different casing than the account's actual sign-in email would otherwise be an
 * exact-match miss on every lookup keyed straight off that string — silently dropping that
 * member's data everywhere (team-scoped pipeline candidates, team KPI dashboards, etc.) even
 * though the member is genuinely in the team. Returns the canonical key so callers can tag/
 * store that instead of the original, possibly-miscased, string.
 */
const resolveUserDataEntry = (allData: Record<string, UserData>, email: string): [string, UserData] | null => {
  if (allData[email]) return [email, allData[email]];
  const target = normalizeEmail(email);
  const foundKey = Object.keys(allData).find(k => normalizeEmail(k) === target);
  return foundKey ? [foundKey, allData[foundKey]] : null;
};

interface Team {
  id: string;
  name: string;
  memberEmails: string[];
  createdBy: string;
  createdAt: string;
}

// BCA事業部 is split into two departments — every member belongs to one of these two (or is
// unassigned); "BCA" itself isn't a real assignment, just the header switcher's "show both
// combined" option.
type Department = 'F+' | 'AC';

interface TeamsConfig {
  schemaVersion: number;
  teams: Team[];
  authorizedEditorEmails?: string[];
  memberDepartments?: Record<string, Department>;
}

const getStartOfWeek = (date: Date): Date => {
  const d = new Date(date);
  const day = d.getDay(); // Sunday - 0, Monday - 1, ...
  const diff = d.getDate() - day; // Adjust date to Sunday
  d.setDate(diff);
  d.setHours(0, 0, 0, 0);
  return d;
};


const chartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    interaction: {
        mode: 'index' as const,
        intersect: false,
    },
    stacked: false,
    plugins: {
        title: {
            display: false,
        },
        legend: {
            position: 'top' as const,
        },
    },
    scales: {
        'y-axis-count': {
            type: 'linear' as const,
            display: true,
            position: 'left' as const,
            title: {
                display: true,
                text: '件数',
                font: { size: 14 }
            },
            grid: {
                drawOnChartArea: true,
            },
            beginAtZero: true,
        },
        'y-axis-rate': {
            type: 'linear' as const,
            display: true,
            position: 'right' as const,
            title: {
                display: true,
                text: '返信率 (%)',
                font: { size: 14 }
            },
            grid: {
                drawOnChartArea: false, // only show grid for left axis
            },
            beginAtZero: true,
            suggestedMax: 20,
        },
        x: {
           grid: {
              display: false
           }
        }
    },
};

const CurrentMonthPerformanceChart: React.FC<{ data: any }> = ({ data }) => {
    return (
        <div className="chart-container">
            <Line options={chartOptions} data={data} />
        </div>
    );
};

interface TrendMetricDef {
  key: string;
  label: string;
  unit: 'count' | 'percent';
  getValue: (totals: KpiTotals, allMedia: MediaEntry[]) => number;
}

const MONTHLY_TREND_COLORS = [
  '#e83e8c', '#0d6efd', '#20c997', '#fd7e14', '#6610f2', '#198754', '#dc3545',
  '#0dcaf0', '#6c757d', '#ffc107', '#3c8abe', '#a9def9', '#28a745', '#adb5bd', '#d63384',
];

const buildTrendMetrics = (): TrendMetricDef[] => [
  { key: 'scoutsSent', label: 'スカウト数', unit: 'count', getValue: (t, m) => getTotalFromLump(t, '_scoutsSent', m) },
  { key: 'scoutReplies', label: 'スカウト返信数', unit: 'count', getValue: (t, m) => getTotalFromLump(t, '_scoutReplies', m) },
  { key: 'replyRate', label: 'スカウト返信率(%)', unit: 'percent', getValue: (t, m) => {
      const sent = getTotalFromLump(t, '_scoutsSent', m);
      const replies = getTotalFromLump(t, '_scoutReplies', m);
      return sent > 0 ? (replies / sent) * 100 : 0;
  } },
  { key: 'effectiveReplies', label: '有効返信数', unit: 'count', getValue: (t, m) => getTotalFromLump(t, '_effectiveReplies', m) },
  { key: 'effectiveReplyRate', label: '有効返信率(%)', unit: 'percent', getValue: (t, m) => {
      const replies = getTotalFromLump(t, '_scoutReplies', m);
      const effective = getTotalFromLump(t, '_effectiveReplies', m);
      return replies > 0 ? (effective / replies) * 100 : 0;
  } },
  { key: 'documentsCollected', label: '書類回収数', unit: 'count', getValue: (t, m) => getTotalFromLump(t, '_documentsCollected', m) },
  { key: 'effectiveDocumentsCollected', label: '有効書類回収数', unit: 'count', getValue: (t, m) => getTotalFromLump(t, '_effectiveDocumentsCollected', m) },
  { key: 'initialInterviews', label: '初回面談数', unit: 'count', getValue: (t, m) => getTotalFromLump(t, '_initialInterviews', m) },
  { key: 'effectiveInitialInterviews', label: '初回有効面談数', unit: 'count', getValue: (t, m) => getTotalFromLump(t, '_effectiveInitialInterviews', m) },
  { key: 'effectiveInterviewRate', label: '初回有効面談率(%)', unit: 'percent', getValue: (t, m) => {
      const interviews = getTotalFromLump(t, '_initialInterviews', m);
      const effective = getTotalFromLump(t, '_effectiveInitialInterviews', m);
      return interviews > 0 ? (effective / interviews) * 100 : 0;
  } },
  ...(Object.keys(GENERAL_KPIS) as Array<keyof typeof GENERAL_KPIS>).map(key => ({
    key: key as string,
    label: GENERAL_KPIS[key].label,
    unit: 'count' as const,
    getValue: (t: KpiTotals) => t[key] || 0,
  })),
];

const TREND_METRICS: TrendMetricDef[] = buildTrendMetrics();
const DEFAULT_TREND_METRIC_KEYS = ['scoutReplies'];

/**
 * Generalizes the old fixed 7-series month-over-month chart into a metric picker: there are too
 * many possible KPIs (general + per-media-suffix + derived rates) to show them all at once
 * without a cluttered, unreadable chart, so the caller's entries get bucketed by month and the
 * user chooses which of TREND_METRICS to plot. Works the same whether `entries` is one person's
 * data or several users' entries concatenated together (summed per month either way).
 */
const buildZeroKpiTotals = (allMedia: MediaEntry[]): KpiTotals => {
    const allKeys = buildAllKpiKeys(allMedia);
    return allKeys.reduce((acc, key) => { acc[key] = 0; return acc; }, {} as KpiTotals);
};

/**
 * Generalizes the old fixed 7-series month-over-month chart into a metric picker: there are too
 * many possible KPIs (general + per-media-suffix + derived rates) to show them all at once
 * without a cluttered, unreadable chart, so the caller's per-user entries get bucketed by month
 * and the user chooses which of TREND_METRICS to plot. Two display modes: "合計" sums every
 * selected metric across all of `perUserEntries` into one line each; "選択ユーザーで比較" instead
 * fixes on a single metric and draws one line per user, so trends are directly comparable
 * between people/teams. The 合計-only mode is used when there's nothing to compare (personal tab).
 */
const MonthlyTrendChart: React.FC<{ perUserEntries: { label: string; entries: KpiEntry[] }[]; allMedia: MediaEntry[] }> = ({ perUserEntries, allMedia }) => {
    const [compareMode, setCompareMode] = useState<'total' | 'byUser'>('total');
    const [selectedKeys, setSelectedKeys] = useState<string[]>(DEFAULT_TREND_METRIC_KEYS);
    const [selectedUserMetricKey, setSelectedUserMetricKey] = useState<string>(DEFAULT_TREND_METRIC_KEYS[0]);
    // 'all' sums across every media (existing behavior); a specific media id scopes every
    // metric's getValue to just that media. GENERAL_KPIS-derived metrics (候補者推薦数 onward)
    // aren't tagged by sourcing media at all, so they're hidden while a specific media is
    // selected — showing them would just repeat the same all-media value regardless of choice.
    const [selectedMediaId, setSelectedMediaId] = useState<string>('all');
    const toggleMetric = (key: string) => {
        setSelectedKeys(prev => prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]);
    };

    const mediaScope = useMemo(
        () => (selectedMediaId === 'all' ? allMedia : allMedia.filter(m => m.id === selectedMediaId)),
        [selectedMediaId, allMedia]
    );
    const visibleMetrics = useMemo(
        () => (selectedMediaId === 'all' ? TREND_METRICS : TREND_METRICS.filter(m => !(m.key in GENERAL_KPIS))),
        [selectedMediaId]
    );
    useEffect(() => {
        const visibleKeys = visibleMetrics.map(m => m.key);
        setSelectedKeys(prev => {
            const filtered = prev.filter(k => visibleKeys.includes(k));
            return filtered.length > 0 ? filtered : [visibleKeys[0]];
        });
        setSelectedUserMetricKey(prev => (visibleKeys.includes(prev) ? prev : visibleKeys[0]));
    }, [visibleMetrics]);

    const zeroTotals = useMemo(() => buildZeroKpiTotals(allMedia), [allMedia]);

    const monthlyTotalsByUser = useMemo(() => perUserEntries.map(({ label, entries }) => {
        const monthlyData: Record<string, KpiTotals> = {};
        entries.forEach(entry => {
            const month = entry.date.substring(0, 7); // YYYY-MM
            if (!monthlyData[month]) monthlyData[month] = { ...zeroTotals };
            (Object.keys(entry.values) as KpiKey[]).forEach(key => {
                monthlyData[month][key] += entry.values[key] || 0;
            });
        });
        return { label, monthlyData };
    }), [perUserEntries, zeroTotals]);

    const allMonthLabels = useMemo(() => {
        const set = new Set<string>();
        monthlyTotalsByUser.forEach(u => Object.keys(u.monthlyData).forEach(m => set.add(m)));
        return Array.from(set).sort();
    }, [monthlyTotalsByUser]);

    const hasAnyData = useMemo(() => perUserEntries.some(u => u.entries.length > 0), [perUserEntries]);
    const canCompareByUser = perUserEntries.length > 1;

    const active = useMemo(() => {
        if (compareMode === 'byUser' && canCompareByUser) {
            const def = visibleMetrics.find(m => m.key === selectedUserMetricKey);
            if (!def) return { labels: allMonthLabels, datasets: [], maxRate: 0, hasCount: false, hasRate: false };
            let maxRate = 0;
            const datasets = monthlyTotalsByUser.map((u, i) => {
                const data = allMonthLabels.map(month => def.getValue(u.monthlyData[month] || zeroTotals, mediaScope));
                if (def.unit === 'percent') maxRate = Math.max(maxRate, ...data);
                const color = MONTHLY_TREND_COLORS[i % MONTHLY_TREND_COLORS.length];
                return { label: u.label, data, borderColor: color, backgroundColor: color, yAxisID: def.unit === 'percent' ? 'y-axis-rate' : 'y-axis-count', tension: 0.2, fill: false };
            });
            return { labels: allMonthLabels, datasets, maxRate, hasCount: def.unit === 'count', hasRate: def.unit === 'percent' };
        }

        const activeDefs = visibleMetrics.filter(m => selectedKeys.includes(m.key));
        let maxRate = 0;
        const datasets = activeDefs.map(def => {
            const data = allMonthLabels.map(month => monthlyTotalsByUser.reduce(
                (sum, u) => sum + def.getValue(u.monthlyData[month] || zeroTotals, mediaScope), 0
            ));
            if (def.unit === 'percent') maxRate = Math.max(maxRate, ...data);
            const color = MONTHLY_TREND_COLORS[TREND_METRICS.indexOf(def) % MONTHLY_TREND_COLORS.length];
            return { label: def.label, data, borderColor: color, backgroundColor: color, yAxisID: def.unit === 'percent' ? 'y-axis-rate' : 'y-axis-count', tension: 0.2, fill: false };
        });
        return { labels: allMonthLabels, datasets, maxRate, hasCount: activeDefs.some(d => d.unit === 'count'), hasRate: activeDefs.some(d => d.unit === 'percent') };
    }, [compareMode, canCompareByUser, selectedUserMetricKey, selectedKeys, allMonthLabels, monthlyTotalsByUser, mediaScope, zeroTotals, visibleMetrics]);

    const options = useMemo(() => ({
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
            mode: 'index' as const,
            intersect: false,
        },
        plugins: {
            title: { display: false },
            legend: { position: 'top' as const },
        },
        scales: {
            ...(active.hasCount ? {
                'y-axis-count': {
                    type: 'linear' as const,
                    display: true,
                    position: 'left' as const,
                    title: { display: true, text: '件数', font: { size: 14 } },
                    grid: { drawOnChartArea: true },
                    beginAtZero: true,
                },
            } : {}),
            ...(active.hasRate ? {
                'y-axis-rate': {
                    type: 'linear' as const,
                    display: true,
                    position: 'right' as const,
                    title: { display: true, text: '%', font: { size: 14 } },
                    grid: { drawOnChartArea: false },
                    beginAtZero: true,
                    max: active.maxRate > 0 ? Math.ceil(active.maxRate * 1.2) : 10,
                },
            } : {}),
            x: {
               grid: { display: false }
            }
        },
    }), [active.hasCount, active.hasRate, active.maxRate]);

    return (
        <div>
            <div className="pipeline-sort-controls" style={{ marginBottom: '0.75rem' }}>
                <span>媒体で切り替え:</span>
                <button
                    type="button"
                    className={selectedMediaId === 'all' ? 'active' : ''}
                    onClick={() => setSelectedMediaId('all')}
                >
                    全媒体合計
                </button>
                {allMedia.map(m => (
                    <button
                        key={m.id}
                        type="button"
                        className={selectedMediaId === m.id ? 'active' : ''}
                        onClick={() => setSelectedMediaId(m.id)}
                    >
                        {m.name}{m.isArchived ? '（アーカイブ済み）' : ''}
                    </button>
                ))}
            </div>
            {selectedMediaId !== 'all' && (
                <p className="gmail-scout-message">
                    単一媒体を表示中は、候補者推薦数以降（媒体を区別しない項目）は選択できません。
                </p>
            )}
            {canCompareByUser && (
                <div className="trend-compare-mode">
                    <label>
                        <input type="radio" name="trend-compare-mode" checked={compareMode === 'total'} onChange={() => setCompareMode('total')} />
                        合計で表示
                    </label>
                    <label>
                        <input type="radio" name="trend-compare-mode" checked={compareMode === 'byUser'} onChange={() => setCompareMode('byUser')} />
                        選択ユーザーで比較
                    </label>
                </div>
            )}
            {compareMode === 'byUser' && canCompareByUser ? (
                <div className="trend-metric-picker">
                    <label className="trend-metric-select-label">
                        比較する項目:
                        <select value={selectedUserMetricKey} onChange={(e) => setSelectedUserMetricKey(e.target.value)}>
                            {visibleMetrics.map(def => <option key={def.key} value={def.key}>{def.label}</option>)}
                        </select>
                    </label>
                </div>
            ) : (
                <div className="trend-metric-picker">
                    {visibleMetrics.map(def => (
                        <label key={def.key} className="trend-metric-checkbox">
                            <input type="checkbox" checked={selectedKeys.includes(def.key)} onChange={() => toggleMetric(def.key)} />
                            {def.label}
                        </label>
                    ))}
                </div>
            )}
            {!hasAnyData ? (
                <p className="no-data-message">実績データがありません。</p>
            ) : active.datasets.length === 0 ? (
                <p className="no-data-message">表示する項目を選択してください。</p>
            ) : (
                <div className="chart-container">
                    <Line options={options as any} data={active as any} />
                </div>
            )}
        </div>
    );
};


const DayOfWeekReplyRateChart: React.FC<{ data: any }> = ({ data }) => {
    const options = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                display: false,
            },
            title: {
                display: false,
            },
            tooltip: {
                callbacks: {
                    label: function(context: any) {
                        let label = context.dataset.label || '';
                        if (label) {
                            label += ': ';
                        }
                        if (context.parsed.y !== null) {
                            label += context.parsed.y.toFixed(1) + '%';
                        }
                        return label;
                    }
                }
            }
        },
        scales: {
            y: {
                beginAtZero: true,
                title: {
                    display: true,
                    text: '返信率 (%)'
                },
                ticks: {
                    callback: function(value: any) {
                        return value + '%';
                    }
                }
            },
            x: {
               grid: {
                  display: false
               }
            }
        },
    };

    return (
        <div className="chart-container">
            <Bar options={options} data={data} />
        </div>
    );
};


/** Aggregates several users' entries/weekly targets into the same shape WeeklySummary expects, for reuse in the team/all-users views. */
function computeAggregateWeeklyData(
  users: string[],
  allUsersData: Record<string, UserData>,
  activeMedia: MediaEntry[],
  weekStartDate: Date
): { data: WeeklyData; weeklyKpiTargets: Record<KpiKey, number> } {
  const weekStart = weekStartDate.getTime();
  const weekEnd = new Date(weekStartDate).setDate(weekStartDate.getDate() + 6);
  const allKeys = buildAllKpiKeys(activeMedia);

  const weeklyTotals = {} as KpiTotals;
  const weeklyKpiTargets = {} as Record<KpiKey, number>;
  allKeys.forEach(key => { weeklyTotals[key] = 0; weeklyKpiTargets[key] = 0; });

  users.forEach(user => {
    const userData = allUsersData[user];
    if (!userData) return;
    (userData.entries || []).forEach(entry => {
      const entryTime = new Date(entry.date).getTime();
      if (entryTime >= weekStart && entryTime <= weekEnd) {
        allKeys.forEach(key => { weeklyTotals[key] += entry.values[key] || 0; });
      }
    });
    const targets = userData.weeklyKpiTargets || {};
    allKeys.forEach(key => { weeklyKpiTargets[key] += targets[key] || 0; });
  });

  const mediaStats = activeMedia.map(source => {
    const sourceKey = source.id;
    return {
      source: source.name,
      id: source.id,
      scoutsSent: weeklyTotals[`${sourceKey}_scoutsSent` as KpiKey] || 0,
      scoutReplies: weeklyTotals[`${sourceKey}_scoutReplies` as KpiKey] || 0,
      effectiveReplies: weeklyTotals[`${sourceKey}_effectiveReplies` as KpiKey] || 0,
      documentsCollected: weeklyTotals[`${sourceKey}_documentsCollected` as KpiKey] || 0,
      effectiveDocumentsCollected: weeklyTotals[`${sourceKey}_effectiveDocumentsCollected` as KpiKey] || 0,
      initialInterviews: weeklyTotals[`${sourceKey}_initialInterviews` as KpiKey] || 0,
      effectiveInitialInterviews: weeklyTotals[`${sourceKey}_effectiveInitialInterviews` as KpiKey] || 0,
    };
  });

  const totalCandidatesSubmitted = weeklyTotals.candidatesSubmitted || 0;
  const totalInitialInterviews = mediaStats.reduce((sum, stat) => sum + stat.initialInterviews, 0);

  return { data: { mediaStats, totalCandidatesSubmitted, totalInitialInterviews }, weeklyKpiTargets };
}

function weeklyMediaStatsToCsvRows(data: WeeklyData, weeklyKpiTargets: Record<KpiKey, number>): string {
  let csv = '媒体,スカウト数,目標,返信数,目標,有効返信数,目標,書類回収数,目標,有効書類回収数,目標,初回面談数,目標,初回有効面談数,目標\n';
  data.mediaStats.forEach(stat => {
    const sourceKey = stat.id;
    const target = (suffix: string) => weeklyKpiTargets[`${sourceKey}_${suffix}` as KpiKey] || 0;
    csv += [
      `"${stat.source}"`,
      stat.scoutsSent, target('scoutsSent'),
      stat.scoutReplies, target('scoutReplies'),
      stat.effectiveReplies, target('effectiveReplies'),
      stat.documentsCollected, target('documentsCollected'),
      stat.effectiveDocumentsCollected, target('effectiveDocumentsCollected'),
      stat.initialInterviews, target('initialInterviews'),
      stat.effectiveInitialInterviews, target('effectiveInitialInterviews'),
    ].join(',') + '\n';
  });
  return csv;
}

/**
 * Builds a CSV covering everything shown in the team/all-users dashboard: each member's
 * monthly progress (same columns as the on-screen table), the team's aggregate weekly
 * summary, and each member's individual weekly summary.
 */
function buildTeamProgressCsv(
  label: string,
  users: string[],
  allUsersData: Record<string, UserData>,
  allMedia: MediaEntry[],
  weekStartDate: Date
): string {
  const activeMedia = allMedia.filter(m => !m.isArchived);
  const now = new Date();
  const generalKeys = Object.keys(GENERAL_KPIS) as Array<keyof typeof GENERAL_KPIS>;

  let csv = '﻿';
  csv += `チーム別進捗レポート: ${label}\n`;
  csv += `出力日: ${now.getFullYear()}/${now.getMonth() + 1}/${now.getDate()}\n\n`;

  csv += `月次進捗（${now.getFullYear()}年${now.getMonth() + 1}月）\n`;
  csv += [
    'ユーザー', 'スカウト送信数', 'スカウト返信数', 'スカウト返信率(%)', '有効返信数', '有効返信率(%)',
    '書類回収数', '書類回収目標', '有効書類回収数', '有効書類回収目標',
    '初回面談数', '初回面談目標', '初回有効面談数', '有効面談率(%)',
    ...generalKeys.map(k => GENERAL_KPIS[k].label),
  ].join(',') + '\n';

  users.forEach(user => {
    const userData = allUsersData[user];
    if (!userData) return;
    const displayName = userData.displayName || user;
    const monthlyTotals = calculateMonthlyTotals(userData.entries || [], allMedia);
    const kpiTargets = { ...buildDefaultKpiTargets(allMedia), ...(userData.kpiTargets || {}) };

    const sent = getTotalFromLump(monthlyTotals, '_scoutsSent', allMedia);
    const replies = getTotalFromLump(monthlyTotals, '_scoutReplies', allMedia);
    const effectiveReplies = getTotalFromLump(monthlyTotals, '_effectiveReplies', allMedia);
    const replyRate = sent > 0 ? (replies / sent) * 100 : 0;
    const effectiveReplyRate = replies > 0 ? (effectiveReplies / replies) * 100 : 0;

    // Actuals (numerator) include archived media, since their historical performance still
    // counts — but targets (denominator) must only sum activeMedia: the 月次目標設定 form only
    // ever lets anyone edit targets for active media, so summing allMedia would silently add in
    // stale/default target values for archived media nobody can see or edit, making this total
    // never match what was actually set in the settings form.
    const documentsCollected = getTotalFromLump(monthlyTotals, '_documentsCollected', allMedia);
    const documentsCollectedTarget = getTotalFromLump(kpiTargets, '_documentsCollected', activeMedia);
    const effectiveDocumentsCollected = getTotalFromLump(monthlyTotals, '_effectiveDocumentsCollected', allMedia);
    const effectiveDocumentsCollectedTarget = getTotalFromLump(kpiTargets, '_effectiveDocumentsCollected', activeMedia);

    const initialInterviews = getTotalFromLump(monthlyTotals, '_initialInterviews', allMedia);
    const initialInterviewsTarget = getTotalFromLump(kpiTargets, '_initialInterviews', activeMedia);
    const effectiveInitialInterviews = getTotalFromLump(monthlyTotals, '_effectiveInitialInterviews', allMedia);
    const effectiveInterviewRate = initialInterviews > 0 ? (effectiveInitialInterviews / initialInterviews) * 100 : 0;

    csv += [
      `"${displayName}"`,
      sent, replies, replyRate.toFixed(1), effectiveReplies, effectiveReplyRate.toFixed(1),
      documentsCollected, documentsCollectedTarget,
      effectiveDocumentsCollected, effectiveDocumentsCollectedTarget,
      initialInterviews, initialInterviewsTarget,
      effectiveInitialInterviews, effectiveInterviewRate.toFixed(1),
      ...generalKeys.map(k => monthlyTotals[k] || 0),
    ].join(',') + '\n';
  });
  csv += '\n';

  const { data: aggData, weeklyKpiTargets: aggTargets } = computeAggregateWeeklyData(users, allUsersData, activeMedia, weekStartDate);
  const weekEnd = new Date(weekStartDate);
  weekEnd.setDate(weekStartDate.getDate() + 6);
  const fmt = (d: Date) => `${d.getMonth() + 1}/${d.getDate()}`;
  csv += `週間サマリー（合計） ${fmt(weekStartDate)}~${fmt(weekEnd)}\n`;
  csv += weeklyMediaStatsToCsvRows(aggData, aggTargets);
  csv += `候補者推薦数,${aggData.totalCandidatesSubmitted}\n`;
  csv += `初回面談数,${aggData.totalInitialInterviews}\n\n`;

  csv += `メンバー別週間サマリー\n`;
  users.forEach(user => {
    const userData = allUsersData[user];
    if (!userData) return;
    const displayName = userData.displayName || user;
    const { data, weeklyKpiTargets } = computeAggregateWeeklyData([user], allUsersData, activeMedia, weekStartDate);
    csv += `${displayName}\n`;
    csv += weeklyMediaStatsToCsvRows(data, weeklyKpiTargets);
    csv += '\n';
  });

  return csv;
}

function mediaStatsFromTotals(totals: KpiTotals, activeMedia: MediaEntry[]): WeeklyMediaStats[] {
  return activeMedia.map(source => {
    const sourceKey = source.id;
    return {
      source: source.name,
      id: source.id,
      scoutsSent: totals[`${sourceKey}_scoutsSent` as KpiKey] || 0,
      scoutReplies: totals[`${sourceKey}_scoutReplies` as KpiKey] || 0,
      effectiveReplies: totals[`${sourceKey}_effectiveReplies` as KpiKey] || 0,
      documentsCollected: totals[`${sourceKey}_documentsCollected` as KpiKey] || 0,
      effectiveDocumentsCollected: totals[`${sourceKey}_effectiveDocumentsCollected` as KpiKey] || 0,
      initialInterviews: totals[`${sourceKey}_initialInterviews` as KpiKey] || 0,
      effectiveInitialInterviews: totals[`${sourceKey}_effectiveInitialInterviews` as KpiKey] || 0,
    };
  });
}

/** Same media-breakdown shape as weeklyMediaStatsToCsvRows, but without target columns — a
 * custom date range has no corresponding weekly/monthly target to compare against. */
function mediaStatsToCsvRowsNoTarget(mediaStats: WeeklyMediaStats[]): string {
  let csv = '媒体,スカウト数,返信数,有効返信数,書類回収数,有効書類回収数,初回面談数,初回有効面談数,返信率(%),有効返信率(%)\n';
  mediaStats.forEach(stat => {
    const replyRate = stat.scoutsSent > 0 ? (stat.scoutReplies / stat.scoutsSent) * 100 : 0;
    const effectiveReplyRate = stat.scoutReplies > 0 ? (stat.effectiveReplies / stat.scoutReplies) * 100 : 0;
    csv += [
      `"${stat.source}"`,
      stat.scoutsSent, stat.scoutReplies, stat.effectiveReplies,
      stat.documentsCollected, stat.effectiveDocumentsCollected,
      stat.initialInterviews, stat.effectiveInitialInterviews,
      replyRate.toFixed(1), effectiveReplyRate.toFixed(1),
    ].join(',') + '\n';
  });
  return csv;
}

/**
 * Same overall shape as buildTeamProgressCsv (per-member totals + aggregate/per-member media
 * breakdown), but for an arbitrary custom date range instead of "this month" + "this week" —
 * so there are no weekly/monthly targets to show alongside the actuals.
 */
function buildTeamProgressCsvForRange(
  label: string,
  users: string[],
  allUsersData: Record<string, UserData>,
  allMedia: MediaEntry[],
  startDate: Date,
  endDate: Date
): string {
  const activeMedia = allMedia.filter(m => !m.isArchived);
  const generalKeys = Object.keys(GENERAL_KPIS) as Array<keyof typeof GENERAL_KPIS>;
  const fmtDate = (d: Date) => `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;

  let csv = '﻿';
  csv += `進捗レポート（カスタム期間）: ${label}\n`;
  csv += `期間: ${fmtDate(startDate)} ~ ${fmtDate(endDate)}\n\n`;

  csv += `メンバー別実績（期間内合計）\n`;
  csv += [
    'ユーザー', 'スカウト送信数', 'スカウト返信数', 'スカウト返信率(%)', '有効返信数', '有効返信率(%)',
    '書類回収数', '有効書類回収数', '初回面談数', '初回有効面談数', '有効面談率(%)',
    ...generalKeys.map(k => GENERAL_KPIS[k].label),
  ].join(',') + '\n';

  const perUserTotals: { displayName: string; totals: KpiTotals }[] = [];
  users.forEach(user => {
    const userData = allUsersData[user];
    if (!userData) return;
    const displayName = userData.displayName || user;
    const totals = calculateTotalsForRange(userData.entries || [], allMedia, startDate, endDate);
    perUserTotals.push({ displayName, totals });

    const sent = getTotalFromLump(totals, '_scoutsSent', allMedia);
    const replies = getTotalFromLump(totals, '_scoutReplies', allMedia);
    const effectiveReplies = getTotalFromLump(totals, '_effectiveReplies', allMedia);
    const replyRate = sent > 0 ? (replies / sent) * 100 : 0;
    const effectiveReplyRate = replies > 0 ? (effectiveReplies / replies) * 100 : 0;
    const documentsCollected = getTotalFromLump(totals, '_documentsCollected', allMedia);
    const effectiveDocumentsCollected = getTotalFromLump(totals, '_effectiveDocumentsCollected', allMedia);
    const initialInterviews = getTotalFromLump(totals, '_initialInterviews', allMedia);
    const effectiveInitialInterviews = getTotalFromLump(totals, '_effectiveInitialInterviews', allMedia);
    const effectiveInterviewRate = initialInterviews > 0 ? (effectiveInitialInterviews / initialInterviews) * 100 : 0;

    csv += [
      `"${displayName}"`,
      sent, replies, replyRate.toFixed(1), effectiveReplies, effectiveReplyRate.toFixed(1),
      documentsCollected, effectiveDocumentsCollected,
      initialInterviews, effectiveInitialInterviews, effectiveInterviewRate.toFixed(1),
      ...generalKeys.map(k => totals[k] || 0),
    ].join(',') + '\n';
  });
  csv += '\n';

  const aggregateTotals = {} as KpiTotals;
  buildAllKpiKeys(allMedia).forEach(key => { aggregateTotals[key] = 0; });
  perUserTotals.forEach(({ totals }) => {
    (Object.keys(totals) as KpiKey[]).forEach(key => {
      aggregateTotals[key] = (aggregateTotals[key] || 0) + (totals[key] || 0);
    });
  });

  csv += `媒体別実績（合計）\n`;
  csv += mediaStatsToCsvRowsNoTarget(mediaStatsFromTotals(aggregateTotals, activeMedia));
  csv += `候補者推薦数,${aggregateTotals.candidatesSubmitted || 0}\n\n`;

  csv += `メンバー別 媒体実績\n`;
  perUserTotals.forEach(({ displayName, totals }) => {
    csv += `${displayName}\n`;
    csv += mediaStatsToCsvRowsNoTarget(mediaStatsFromTotals(totals, activeMedia));
    csv += '\n';
  });

  return csv;
}

const WeeklySummaryTable: React.FC<{
  data: WeeklyData;
  weeklyKpiTargets: Record<KpiKey, number>;
}> = ({ data, weeklyKpiTargets }) => {
  return (
      <div className="weekly-summary-content">
        <div className="weekly-summary-totals">
          <div className="total-item">
            <span>候補者推薦数</span>
            <strong>{data.totalCandidatesSubmitted}</strong>
          </div>
          <div className="total-item">
            <span>初回面談数</span>
            <strong>{data.totalInitialInterviews}</strong>
          </div>
        </div>
        <table className="weekly-summary-table">
          <thead>
            <tr>
              <th>媒体</th>
              <th>スカウト数</th>
              <th>返信数</th>
              <th>有効返信数</th>
              <th>書類回収数</th>
              <th>有効書類回収数</th>
              <th>初回面談数</th>
              <th>初回有効面談数</th>
              <th>返信率</th>
              <th>有効返信率</th>
            </tr>
          </thead>
          <tbody>
            {data.mediaStats.map(({ source, id, scoutsSent, scoutReplies, effectiveReplies, documentsCollected, effectiveDocumentsCollected, initialInterviews, effectiveInitialInterviews }) => {
              const replyRate = scoutsSent > 0 ? (scoutReplies / scoutsSent) * 100 : 0;
              const effectiveReplyRate = scoutReplies > 0 ? (effectiveReplies / scoutReplies) * 100 : 0;
              const sourceKey = id;

              const scoutsKey = `${sourceKey}_scoutsSent` as KpiKey;
              const repliesKey = `${sourceKey}_scoutReplies` as KpiKey;
              const effectiveRepliesKey = `${sourceKey}_effectiveReplies` as KpiKey;
              const documentsCollectedKey = `${sourceKey}_documentsCollected` as KpiKey;
              const effectiveDocumentsCollectedKey = `${sourceKey}_effectiveDocumentsCollected` as KpiKey;
              const interviewsKey = `${sourceKey}_initialInterviews` as KpiKey;
              const effectiveInterviewsKey = `${sourceKey}_effectiveInitialInterviews` as KpiKey;

              const scoutsTarget = weeklyKpiTargets[scoutsKey] || 0;
              const repliesTarget = weeklyKpiTargets[repliesKey] || 0;
              const effectiveRepliesTarget = weeklyKpiTargets[effectiveRepliesKey] || 0;
              const documentsCollectedTarget = weeklyKpiTargets[documentsCollectedKey] || 0;
              const effectiveDocumentsCollectedTarget = weeklyKpiTargets[effectiveDocumentsCollectedKey] || 0;
              const interviewsTarget = weeklyKpiTargets[interviewsKey] || 0;
              const effectiveInterviewsTarget = weeklyKpiTargets[effectiveInterviewsKey] || 0;

              const scoutsProgress = scoutsTarget > 0 ? Math.min((scoutsSent / scoutsTarget) * 100, 100) : 0;
              const repliesProgress = repliesTarget > 0 ? Math.min((scoutReplies / repliesTarget) * 100, 100) : 0;
              const effectiveRepliesProgress = effectiveRepliesTarget > 0 ? Math.min((effectiveReplies / effectiveRepliesTarget) * 100, 100) : 0;
              const documentsCollectedProgress = documentsCollectedTarget > 0 ? Math.min((documentsCollected / documentsCollectedTarget) * 100, 100) : 0;
              const effectiveDocumentsCollectedProgress = effectiveDocumentsCollectedTarget > 0 ? Math.min((effectiveDocumentsCollected / effectiveDocumentsCollectedTarget) * 100, 100) : 0;
              const interviewsProgress = interviewsTarget > 0 ? Math.min((initialInterviews / interviewsTarget) * 100, 100) : 0;
              const effectiveInterviewsProgress = effectiveInterviewsTarget > 0 ? Math.min((effectiveInitialInterviews / effectiveInterviewsTarget) * 100, 100) : 0;

              return (
                <tr key={source}>
                  <td>{source}</td>
                  <td className="summary-progress-cell">
                    <span>{scoutsSent} / {scoutsTarget}</span>
                    <div className="mini-progress-bar"><div className="progress-bar-fill" style={{ width: `${scoutsProgress}%` }}></div></div>
                  </td>
                  <td className="summary-progress-cell">
                    <span>{scoutReplies} / {repliesTarget}</span>
                    <div className="mini-progress-bar"><div className="progress-bar-fill" style={{ width: `${repliesProgress}%` }}></div></div>
                  </td>
                  <td className="summary-progress-cell">
                    <span>{effectiveReplies} / {effectiveRepliesTarget}</span>
                    <div className="mini-progress-bar"><div className="progress-bar-fill" style={{ width: `${effectiveRepliesProgress}%`, backgroundColor: 'var(--info-color)' }}></div></div>
                  </td>
                  <td className="summary-progress-cell">
                    <span>{documentsCollected} / {documentsCollectedTarget}</span>
                    <div className="mini-progress-bar"><div className="progress-bar-fill" style={{ width: `${documentsCollectedProgress}%` }}></div></div>
                  </td>
                  <td className="summary-progress-cell">
                    <span>{effectiveDocumentsCollected} / {effectiveDocumentsCollectedTarget}</span>
                    <div className="mini-progress-bar"><div className="progress-bar-fill" style={{ width: `${effectiveDocumentsCollectedProgress}%`, backgroundColor: 'var(--info-color)' }}></div></div>
                  </td>
                  <td className="summary-progress-cell">
                    <span>{initialInterviews} / {interviewsTarget}</span>
                    <div className="mini-progress-bar"><div className="progress-bar-fill" style={{ width: `${interviewsProgress}%` }}></div></div>
                  </td>
                  <td className="summary-progress-cell">
                    <span>{effectiveInitialInterviews} / {effectiveInterviewsTarget}</span>
                    <div className="mini-progress-bar"><div className="progress-bar-fill" style={{ width: `${effectiveInterviewsProgress}%`, backgroundColor: 'var(--info-color)' }}></div></div>
                  </td>
                  <td>{replyRate.toFixed(1)}%</td>
                  <td>{effectiveReplyRate.toFixed(1)}%</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
  );
};

const WeeklySummary: React.FC<{
  weekStartDate: Date;
  data: WeeklyData;
  weeklyKpiTargets: Record<KpiKey, number>;
  onPrevWeek: () => void;
  onNextWeek: () => void;
}> = ({ weekStartDate, data, weeklyKpiTargets, onPrevWeek, onNextWeek }) => {
  const endDate = new Date(weekStartDate);
  endDate.setDate(weekStartDate.getDate() + 6);

  const formatDate = (d: Date) => `${d.getMonth() + 1}月${d.getDate()}日`;
  const weekRange = `${formatDate(weekStartDate)} - ${formatDate(endDate)}`;

  const isThisWeek = getStartOfWeek(new Date()).getTime() === weekStartDate.getTime();

  return (
    <div className="weekly-summary-container">
      <div className="weekly-summary-header">
        <h3>{weekRange}</h3>
        <div>
          <button onClick={onPrevWeek} aria-label="前の週へ">&lt; 前の週</button>
          <button onClick={onNextWeek} disabled={isThisWeek} aria-label="次の週へ">次の週 &gt;</button>
        </div>
      </div>
      <WeeklySummaryTable data={data} weeklyKpiTargets={weeklyKpiTargets} />
    </div>
  );
};


const MediaKpiCard: React.FC<{
    source: MediaEntry;
    monthlyTotals: KpiTotals;
    kpiTargets: Record<KpiKey, number>;
}> = ({ source, monthlyTotals, kpiTargets }) => {
    const sourceKey = source.id;
    const scoutsKey = `${sourceKey}_scoutsSent` as KpiKey;
    const repliesKey = `${sourceKey}_scoutReplies` as KpiKey;
    const effectiveRepliesKey = `${sourceKey}_effectiveReplies` as KpiKey;
    const documentsCollectedKey = `${sourceKey}_documentsCollected` as KpiKey;
    const effectiveDocumentsCollectedKey = `${sourceKey}_effectiveDocumentsCollected` as KpiKey;
    const interviewsKey = `${sourceKey}_initialInterviews` as KpiKey;
    const effectiveInterviewsKey = `${sourceKey}_effectiveInitialInterviews` as KpiKey;

    const scoutsSent = monthlyTotals[scoutsKey] || 0;
    const scoutReplies = monthlyTotals[repliesKey] || 0;
    const effectiveReplies = monthlyTotals[effectiveRepliesKey] || 0;
    const documentsCollected = monthlyTotals[documentsCollectedKey] || 0;
    const effectiveDocumentsCollected = monthlyTotals[effectiveDocumentsCollectedKey] || 0;
    const initialInterviews = monthlyTotals[interviewsKey] || 0;
    const effectiveInitialInterviews = monthlyTotals[effectiveInterviewsKey] || 0;
    
    const scoutsTarget = kpiTargets[scoutsKey] || 0;
    const repliesTarget = kpiTargets[repliesKey] || 0;
    const effectiveRepliesTarget = kpiTargets[effectiveRepliesKey] || 0;
    const documentsCollectedTarget = kpiTargets[documentsCollectedKey] || 0;
    const effectiveDocumentsCollectedTarget = kpiTargets[effectiveDocumentsCollectedKey] || 0;
    const interviewsTarget = kpiTargets[interviewsKey] || 0;
    const effectiveInterviewsTarget = kpiTargets[effectiveInterviewsKey] || 0;

    const replyRate = scoutsSent > 0 ? (scoutReplies / scoutsSent) * 100 : 0;
    const effectiveReplyRate = scoutReplies > 0 ? (effectiveReplies / scoutReplies) * 100 : 0;

    const scoutsProgress = scoutsTarget > 0 ? Math.min((scoutsSent / scoutsTarget) * 100, 100) : 0;
    const repliesProgress = repliesTarget > 0 ? Math.min((scoutReplies / repliesTarget) * 100, 100) : 0;
    const effectiveRepliesProgress = effectiveRepliesTarget > 0 ? Math.min((effectiveReplies / effectiveRepliesTarget) * 100, 100) : 0;
    const documentsCollectedProgress = documentsCollectedTarget > 0 ? Math.min((documentsCollected / documentsCollectedTarget) * 100, 100) : 0;
    const effectiveDocumentsCollectedProgress = effectiveDocumentsCollectedTarget > 0 ? Math.min((effectiveDocumentsCollected / effectiveDocumentsCollectedTarget) * 100, 100) : 0;
    const interviewsProgress = interviewsTarget > 0 ? Math.min((initialInterviews / interviewsTarget) * 100, 100) : 0;
    const effectiveInterviewsProgress = effectiveInterviewsTarget > 0 ? Math.min((effectiveInitialInterviews / effectiveInterviewsTarget) * 100, 100) : 0;

    return (
        <div className="media-kpi-card" aria-label={`${source.name}の進捗`}>
            <h3>{source.name}</h3>
             <div className="reply-rate-group">
                <div className="reply-rate-section">
                    <span className="reply-rate-value" aria-label={`返信率 ${replyRate.toFixed(1)}パーセント`}>{replyRate.toFixed(1)}%</span>
                    <span className="reply-rate-label">返信率</span>
                </div>
                <div className="reply-rate-section">
                    <span className="reply-rate-value" aria-label={`有効返信率 ${effectiveReplyRate.toFixed(1)}パーセント`}>{effectiveReplyRate.toFixed(1)}%</span>
                    <span className="reply-rate-label">有効返信率</span>
                </div>
            </div>
            <div className="stat-item">
                <div className="stat-details">
                    <span>スカウト数</span>
                    <span aria-label={`目標${scoutsTarget}件中${scoutsSent}件`}>{scoutsSent} / {scoutsTarget}</span>
                </div>
                <div className="progress-bar" role="progressbar" aria-valuenow={scoutsSent} aria-valuemin={0} aria-valuemax={scoutsTarget}>
                    <div className="progress-bar-fill" style={{ width: `${scoutsProgress}%` }}></div>
                </div>
            </div>
            <div className="stat-item">
                <div className="stat-details">
                    <span>スカウト返信数</span>
                    <span aria-label={`目標${repliesTarget}件中${scoutReplies}件`}>{scoutReplies} / {repliesTarget}</span>
                </div>
                 <div className="progress-bar" role="progressbar" aria-valuenow={scoutReplies} aria-valuemin={0} aria-valuemax={repliesTarget}>
                    <div className="progress-bar-fill" style={{ width: `${repliesProgress}%` }}></div>
                </div>
            </div>
            <div className="stat-item">
                <div className="stat-details">
                    <span>有効返信数</span>
                    <span aria-label={`目標${effectiveRepliesTarget}件中${effectiveReplies}件`}>{effectiveReplies} / {effectiveRepliesTarget}</span>
                </div>
                 <div className="progress-bar" role="progressbar" aria-valuenow={effectiveReplies} aria-valuemin={0} aria-valuemax={effectiveRepliesTarget}>
                    <div className="progress-bar-fill" style={{ width: `${effectiveRepliesProgress}%`, backgroundColor: 'var(--info-color)' }}></div>
                </div>
            </div>
             <div className="stat-item">
                <div className="stat-details">
                    <span>書類回収数</span>
                    <span aria-label={`目標${documentsCollectedTarget}件中${documentsCollected}件`}>{documentsCollected} / {documentsCollectedTarget}</span>
                </div>
                 <div className="progress-bar" role="progressbar" aria-valuenow={documentsCollected} aria-valuemin={0} aria-valuemax={documentsCollectedTarget}>
                    <div className="progress-bar-fill" style={{ width: `${documentsCollectedProgress}%` }}></div>
                </div>
            </div>
             <div className="stat-item">
                <div className="stat-details">
                    <span>有効書類回収数</span>
                    <span aria-label={`目標${effectiveDocumentsCollectedTarget}件中${effectiveDocumentsCollected}件`}>{effectiveDocumentsCollected} / {effectiveDocumentsCollectedTarget}</span>
                </div>
                 <div className="progress-bar" role="progressbar" aria-valuenow={effectiveDocumentsCollected} aria-valuemin={0} aria-valuemax={effectiveDocumentsCollectedTarget}>
                    <div className="progress-bar-fill" style={{ width: `${effectiveDocumentsCollectedProgress}%`, backgroundColor: 'var(--info-color)' }}></div>
                </div>
            </div>
             <div className="stat-item">
                <div className="stat-details">
                    <span>初回面談数</span>
                    <span aria-label={`目標${interviewsTarget}件中${initialInterviews}件`}>{initialInterviews} / {interviewsTarget}</span>
                </div>
                 <div className="progress-bar" role="progressbar" aria-valuenow={initialInterviews} aria-valuemin={0} aria-valuemax={interviewsTarget}>
                    <div className="progress-bar-fill" style={{ width: `${interviewsProgress}%` }}></div>
                </div>
            </div>
            <div className="stat-item">
                <div className="stat-details">
                    <span>初回有効面談数</span>
                    {/* FIX: Corrected typo in variable name from effectiveInitialInterviewsTarget to effectiveInterviewsTarget */}
                    <span aria-label={`目標${effectiveInterviewsTarget}件中${effectiveInitialInterviews}件`}>{effectiveInitialInterviews} / {effectiveInterviewsTarget}</span>
                </div>
                 <div className="progress-bar" role="progressbar" aria-valuenow={effectiveInitialInterviews} aria-valuemin={0} aria-valuemax={effectiveInterviewsTarget}>
                    <div className="progress-bar-fill" style={{ width: `${effectiveInterviewsProgress}%`, backgroundColor: 'var(--info-color)' }}></div>
                </div>
            </div>
        </div>
    );
};

const DateEntryModal: React.FC<{
  date: string;
  initialValues: KpiTotals | null;
  activeMedia: MediaEntry[];
  onSave: (date: string, values: KpiTotals) => void;
  onNavigate: (currentDate: string, currentValues: KpiTotals, offsetDays: number) => void;
  onClose: () => void;
}> = ({ date, initialValues, activeMedia, onSave, onNavigate, onClose }) => {
  const [entryValues, setEntryValues] = useState<{ [key in KpiKey]?: number }>(
    initialValues || {}
  );
  const [gmailStatus, setGmailStatus] = useState<'idle' | 'loading' | 'error' | 'done'>('idle');
  const [gmailMessage, setGmailMessage] = useState('');
  const [gmailNeedsReauth, setGmailNeedsReauth] = useState(false);

  // This modal instance stays mounted across 前日/次日 navigation (only `date`/`initialValues`
  // change) instead of unmounting like a normal open/close does, so entryValues must be reset
  // by hand here — otherwise the form would keep showing the previous day's in-progress values
  // after navigating.
  useEffect(() => {
    setEntryValues(initialValues || {});
    setGmailStatus('idle');
    setGmailMessage('');
    setGmailNeedsReauth(false);
  }, [date]);

  const formattedDate = new Date(date + 'T00:00:00').toLocaleDateString('ja-JP', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target as { name: KpiKey; value: string };
    setEntryValues(prev => ({ ...prev, [name]: value === '' ? undefined : Number(value) }));
  };

  // Only zero/overwrite general + currently-active-media fields; any archived media's
  // historical value for this day (present in initialValues) is left untouched since
  // there's no input for it in this form anymore.
  const buildEditableKeys = (): KpiKey[] => [
    ...Object.keys(GENERAL_KPIS),
    ...activeMedia.flatMap(media => MEDIA_KPI_SUFFIXES.map(suffix => `${media.id}_${suffix}`)),
  ];

  const handleClear = () => {
      if (window.confirm('この日の実績をすべてクリアします。よろしいですか？')) {
          const emptyValues: KpiTotals = { ...(initialValues || {}) };
          buildEditableKeys().forEach(key => { emptyValues[key] = 0; });
          onSave(date, emptyValues);
      }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const valuesWithDefaults: KpiTotals = { ...(initialValues || {}) };
    buildEditableKeys().forEach(key => { valuesWithDefaults[key] = entryValues[key] || 0; });
    onSave(date, valuesWithDefaults);
  };

  // Saves the current day's in-progress input (same as the submit button) before moving,
  // so switching days never silently drops whatever was just typed.
  const handleNavigate = (offsetDays: number) => {
    const valuesWithDefaults: KpiTotals = { ...(initialValues || {}) };
    buildEditableKeys().forEach(key => { valuesWithDefaults[key] = entryValues[key] || 0; });
    onNavigate(date, valuesWithDefaults, offsetDays);
  };

  const handleFetchGmailReplies = async () => {
    const session = getCurrentSession();
    if (!session) {
      setGmailStatus('error');
      setGmailMessage('ログイン情報が確認できませんでした。再読み込みしてからお試しください。');
      setGmailNeedsReauth(false);
      return;
    }
    setGmailStatus('loading');
    setGmailMessage('');
    setGmailNeedsReauth(false);
    try {
      const result = await fetchScoutReplyCounts(session.accessToken, date);
      const appliedLabels: string[] = [];
      setEntryValues(prev => {
        const next = { ...prev };
        activeMedia.forEach(media => {
          const count = result.counts[media.id];
          if (count !== undefined) {
            next[`${media.id}_scoutReplies` as KpiKey] = count;
            appliedLabels.push(`${media.name}: ${count}件`);
          }
        });
        return next;
      });
      setGmailStatus('done');
      setGmailMessage(
        appliedLabels.length > 0
          ? `Gmailの返信通知メールから反映しました（${appliedLabels.join(' / ')}）。内容を確認のうえ保存してください。`
          : 'この日付に該当するスカウト返信メールは見つかりませんでした。'
      );
    } catch (err) {
      setGmailStatus('error');
      if (err instanceof GmailPermissionError) {
        setGmailMessage('Gmailの読み取り権限がまだ許可されていません。下のボタンから許可してください。');
        setGmailNeedsReauth(true);
      } else {
        setGmailMessage(err instanceof Error ? err.message : 'Gmailの取得に失敗しました。');
      }
    }
  };

  const handleReauthorizeGmail = async () => {
    setGmailStatus('loading');
    setGmailMessage('');
    try {
      await reauthorizeWithConsent();
      await handleFetchGmailReplies();
    } catch (err) {
      setGmailStatus('error');
      setGmailMessage(err instanceof Error ? err.message : 'ログインに失敗しました。');
    }
  };

  const isFormEmpty = Object.values(entryValues).every(val => val === undefined || val === 0 || val === null);
  const isSaveDisabled = isFormEmpty && (!initialValues || Object.values(initialValues).every(v => v === 0));
  const canClear = initialValues && Object.values(initialValues).some(v => v > 0);

  return (
    <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="modal-title">
      <div className="modal-content" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header date-entry-modal-header">
          <button type="button" onClick={() => handleNavigate(-1)} className="secondary-action-button date-entry-nav-button" aria-label="前日の実績入力へ">
            &lt; 前日
          </button>
          <h3 id="modal-title">{formattedDate} の実績入力</h3>
          <button type="button" onClick={() => handleNavigate(1)} className="secondary-action-button date-entry-nav-button" aria-label="次日の実績入力へ">
            次日 &gt;
          </button>
          <button onClick={onClose} className="close-button" aria-label="閉じる">&times;</button>
        </div>
        <form id="date-entry-form" className="modal-body" onSubmit={handleSubmit}>
          <fieldset className="general-kpi-fieldset">
            <legend className="sr-only">全体実績</legend>
            {(Object.keys(GENERAL_KPIS) as (keyof typeof GENERAL_KPIS)[]).map(key => (
              <div key={key} className="form-group">
                <label htmlFor={`modal-${key}`}>{GENERAL_KPIS[key].label}</label>
                <input
                  type="number"
                  id={`modal-${key}`}
                  name={key}
                  value={entryValues[key] ?? ''}
                  onChange={handleInputChange}
                  min="0"
                  placeholder="0"
                  aria-label={`${GENERAL_KPIS[key].label}の数値を入力`}
                />
              </div>
            ))}
          </fieldset>

          <div className="media-kpi-section">
            <h3 className="sub-section-title">媒体別実績</h3>
            <div className="gmail-scout-fetch-bar">
              <button type="button" onClick={handleFetchGmailReplies} disabled={gmailStatus === 'loading'} className="secondary-action-button">
                {gmailStatus === 'loading' ? 'Gmailを確認中...' : 'Gmailから返信数を取得'}
              </button>
              {gmailNeedsReauth && (
                <button type="button" onClick={handleReauthorizeGmail} className="secondary-action-button">
                  Gmailの権限を許可する
                </button>
              )}
              {gmailMessage && (
                <span className={`gmail-scout-message ${gmailStatus === 'error' ? 'is-error' : ''}`}>{gmailMessage}</span>
              )}
            </div>
            {/* One compact row per media instead of a tall stacked fieldset per media —
                the same 7 fields laid out vertically per source was the main source of the
                excessive scrolling in this modal. */}
            <div className="daily-entry-table-container">
              <table className="daily-entry-table">
                <thead>
                  <tr>
                    <th>媒体</th>
                    <th>スカウト数</th>
                    <th>返信数</th>
                    <th>有効返信数</th>
                    <th>書類回収数</th>
                    <th>有効書類回収数</th>
                    <th>初回面談数</th>
                    <th>初回有効面談数</th>
                  </tr>
                </thead>
                <tbody>
                  {activeMedia.map(source => {
                    const sourceKey = source.id;
                    const scoutsKey = `${sourceKey}_scoutsSent` as KpiKey;
                    const repliesKey = `${sourceKey}_scoutReplies` as KpiKey;
                    const effectiveRepliesKey = `${sourceKey}_effectiveReplies` as KpiKey;
                    const documentsCollectedKey = `${sourceKey}_documentsCollected` as KpiKey;
                    const effectiveDocumentsCollectedKey = `${sourceKey}_effectiveDocumentsCollected` as KpiKey;
                    const interviewsKey = `${sourceKey}_initialInterviews` as KpiKey;
                    const effectiveInterviewsKey = `${sourceKey}_effectiveInitialInterviews` as KpiKey;
                    const rowFields: { key: KpiKey; label: string }[] = [
                      { key: scoutsKey, label: 'スカウト数' },
                      { key: repliesKey, label: 'スカウト返信数' },
                      { key: effectiveRepliesKey, label: '有効返信数' },
                      { key: documentsCollectedKey, label: '書類回収数' },
                      { key: effectiveDocumentsCollectedKey, label: '有効書類回収数' },
                      { key: interviewsKey, label: '初回面談数' },
                      { key: effectiveInterviewsKey, label: '初回有効面談数' },
                    ];
                    return (
                      <tr key={source.id}>
                        <th scope="row">{source.name}</th>
                        {rowFields.map(field => (
                          <td key={field.key}>
                            <input
                              type="number"
                              id={`modal-${field.key}`}
                              name={field.key}
                              value={entryValues[field.key] ?? ''}
                              onChange={handleInputChange}
                              min="0"
                              placeholder="0"
                              aria-label={`${source.name} ${field.label}を入力`}
                            />
                          </td>
                        ))}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </form>
        <div className="modal-footer">
          {canClear && <button type="button" onClick={handleClear} className="reset-button" style={{ marginRight: 'auto' }}>実績をクリア</button>}
          <button type="button" onClick={onClose} className="cancel-button">キャンセル</button>
          <button type="submit" form="date-entry-form" className="submit-button" disabled={isSaveDisabled}>実績を保存</button>
        </div>
      </div>
    </div>
  );
};

const BulkGmailReplyImportModal: React.FC<{
  allMedia: MediaEntry[];
  entriesByDate: Map<string, KpiTotals>;
  onApply: (countsByDate: Record<string, Record<string, number>>) => void;
  onClose: () => void;
}> = ({ allMedia, entriesByDate, onApply, onClose }) => {
  const toDateInputValue = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const todayStr = toDateInputValue(new Date());
  const defaultStartStr = (() => {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    return toDateInputValue(d);
  })();

  const [startDate, setStartDate] = useState(defaultStartStr);
  const [endDate, setEndDate] = useState(todayStr);
  const [status, setStatus] = useState<'idle' | 'loading' | 'error' | 'preview'>('idle');
  const [message, setMessage] = useState('');
  const [needsReauth, setNeedsReauth] = useState(false);
  const [progress, setProgress] = useState<{ done: number; total: number } | null>(null);
  const [result, setResult] = useState<ScoutReplyRangeResult | null>(null);

  const rangeDays = Math.round((new Date(endDate + 'T00:00:00').getTime() - new Date(startDate + 'T00:00:00').getTime()) / 86400000) + 1;

  const runScan = async () => {
    const session = getCurrentSession();
    if (!session) {
      setStatus('error');
      setMessage('ログイン情報が確認できませんでした。再読み込みしてからお試しください。');
      return;
    }
    if (startDate > endDate) {
      setStatus('error');
      setMessage('開始日は終了日より前の日付にしてください。');
      return;
    }
    setStatus('loading');
    setMessage('');
    setNeedsReauth(false);
    setProgress(null);
    try {
      const rangeResult = await fetchScoutReplyCountsForRange(session.accessToken, startDate, endDate, (done, total) => setProgress({ done, total }));
      setResult(rangeResult);
      setStatus('preview');
    } catch (err) {
      setStatus('error');
      if (err instanceof GmailPermissionError) {
        setMessage('Gmailの読み取り権限がまだ許可されていません。下のボタンから許可してください。');
        setNeedsReauth(true);
      } else {
        setMessage(err instanceof Error ? err.message : 'Gmailの取得に失敗しました。');
      }
    }
  };

  const handleReauthorize = async () => {
    setStatus('loading');
    setMessage('');
    try {
      await reauthorizeWithConsent();
      await runScan();
    } catch (err) {
      setStatus('error');
      setMessage(err instanceof Error ? err.message : 'ログインに失敗しました。');
    }
  };

  const previewRows = useMemo(() => {
    if (!result) return [];
    return Object.keys(result.countsByDate).sort().map(dateStr => {
      const mediaCounts = result.countsByDate[dateStr];
      const existingValues = entriesByDate.get(dateStr);
      const cells = Object.keys(mediaCounts).map(mediaId => {
        const media = allMedia.find(m => m.id === mediaId);
        const fetchedCount = mediaCounts[mediaId];
        const currentValue = existingValues?.[`${mediaId}_scoutReplies` as KpiKey] || 0;
        return { mediaId, mediaName: media?.name || mediaId, fetchedCount, currentValue, willChange: currentValue !== fetchedCount };
      });
      return { dateStr, cells };
    }).filter(row => row.cells.length > 0);
  }, [result, allMedia, entriesByDate]);

  const totalChanges = previewRows.reduce((acc, row) => acc + row.cells.filter(c => c.willChange).length, 0);

  return (
    <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="bulk-gmail-modal-title">
      <div className="modal-content" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h3 id="bulk-gmail-modal-title">Gmailから返信数を一括取得</h3>
          <button onClick={onClose} className="close-button" aria-label="閉じる">&times;</button>
        </div>
        <div className="modal-body">
          <p className="modal-description">
            指定した期間のGmailからスカウト返信通知メール（BIZ/RDS/Doda/Liiga）を検索し、日付ごとの返信数をまとめて反映します。
            既に入力済みの日も「スカウト返信数」のみ取得結果で上書きされ、他の項目は変更されません。
          </p>
          <div className="bulk-gmail-date-range">
            <div className="form-group">
              <label htmlFor="bulk-gmail-start">開始日</label>
              <input type="date" id="bulk-gmail-start" value={startDate} max={endDate} onChange={(e) => setStartDate(e.target.value)} />
            </div>
            <div className="form-group">
              <label htmlFor="bulk-gmail-end">終了日</label>
              <input type="date" id="bulk-gmail-end" value={endDate} min={startDate} max={todayStr} onChange={(e) => setEndDate(e.target.value)} />
            </div>
          </div>
          {rangeDays > 120 && (
            <p className="gmail-scout-message is-error">期間が{rangeDays}日と長いため、取得に時間がかかる場合があります。</p>
          )}

          <div className="gmail-scout-fetch-bar">
            <button type="button" onClick={runScan} disabled={status === 'loading'} className="secondary-action-button">
              {status === 'loading' ? 'Gmailを確認中...' : 'Gmailを検索する'}
            </button>
            {needsReauth && (
              <button type="button" onClick={handleReauthorize} className="secondary-action-button">Gmailの権限を許可する</button>
            )}
            {status === 'loading' && progress && progress.total > 0 && (
              <span className="gmail-scout-message">{progress.done} / {progress.total} 件を確認中...</span>
            )}
            {message && <span className={`gmail-scout-message ${status === 'error' ? 'is-error' : ''}`}>{message}</span>}
          </div>

          {status === 'preview' && result && (
            <div className="bulk-gmail-preview">
              {previewRows.length === 0 ? (
                <p className="gmail-scout-message">この期間に該当するスカウト返信メールは見つかりませんでした（メール{result.totalScanned}件を確認）。</p>
              ) : (
                <>
                  <p className="gmail-scout-message">
                    {previewRows.length}日分・{result.totalMatched}件のメールを検出しました（うち{totalChanges}件が現在の入力値と異なります）。反映すると各日の「スカウト返信数」が上書きされます。
                  </p>
                  <div className="bulk-gmail-preview-table-container">
                    <table className="bulk-gmail-preview-table">
                      <thead>
                        <tr><th>日付</th><th>媒体</th><th>取得件数</th><th>現在の値</th></tr>
                      </thead>
                      <tbody>
                        {previewRows.map(row => row.cells.map((cell, i) => (
                          <tr key={`${row.dateStr}-${cell.mediaId}`}>
                            {i === 0 && <td rowSpan={row.cells.length}>{row.dateStr}</td>}
                            <td>{cell.mediaName}</td>
                            <td>{cell.fetchedCount}</td>
                            <td className={cell.willChange ? 'is-changed' : ''}>{cell.currentValue}</td>
                          </tr>
                        )))}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>
          )}
        </div>
        <div className="modal-footer">
          <button type="button" onClick={onClose} className="cancel-button">キャンセル</button>
          {status === 'preview' && result && previewRows.length > 0 && (
            <button type="button" onClick={() => onApply(result.countsByDate)} className="submit-button">
              反映する（{previewRows.length}日分）
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

/**
 * Imports a media platform's own daily scout-performance CSV export (BIZ/Doda) directly,
 * instead of relying on gmailScout.ts's Gmail-notification heuristic — these exports already
 * carry both スカウト数 and 返信数 per day straight from the source, so no reply-notification
 * parsing/deduping is needed. Same preview-then-apply UX as BulkGmailReplyImportModal.
 */
const MediaCsvImportModal: React.FC<{
  allMedia: MediaEntry[];
  entriesByDate: Map<string, KpiTotals>;
  onApply: (mediaId: ScoutCsvMediaId, countsByDate: Record<string, ScoutCsvDayCounts>) => void;
  onClose: () => void;
}> = ({ allMedia, entriesByDate, onApply, onClose }) => {
  const [selectedMediaId, setSelectedMediaId] = useState<ScoutCsvMediaId>('biz');
  const [status, setStatus] = useState<'idle' | 'parsing' | 'preview' | 'error'>('idle');
  const [message, setMessage] = useState('');
  const [result, setResult] = useState<ScoutCsvParseResult | null>(null);

  const mediaName = allMedia.find(m => m.id === selectedMediaId)?.name || selectedMediaId.toUpperCase();

  const handleMediaChange = (id: ScoutCsvMediaId) => {
    setSelectedMediaId(id);
    setStatus('idle');
    setMessage('');
    setResult(null);
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setStatus('parsing');
    setMessage('');
    setResult(null);
    try {
      const text = await decodeCsvFile(file);
      const parsed = parseScoutCsv(selectedMediaId, text);
      if (Object.keys(parsed.countsByDate).length === 0) {
        setStatus('error');
        setMessage('CSVから日付ごとのデータを取得できませんでした。フォーマットをご確認ください。');
        return;
      }
      setResult(parsed);
      setStatus('preview');
    } catch (err) {
      setStatus('error');
      setMessage(err instanceof Error ? err.message : 'CSVの読み込みに失敗しました。');
    } finally {
      e.target.value = '';
    }
  };

  const previewRows = useMemo(() => {
    if (!result) return [];
    return Object.keys(result.countsByDate).sort().map(dateStr => {
      const counts = result.countsByDate[dateStr];
      const existingValues = entriesByDate.get(dateStr);
      const currentSent = existingValues?.[`${selectedMediaId}_scoutsSent` as KpiKey] || 0;
      const currentReplies = existingValues?.[`${selectedMediaId}_scoutReplies` as KpiKey] || 0;
      return {
        dateStr,
        scoutsSent: counts.scoutsSent,
        scoutReplies: counts.scoutReplies,
        currentSent,
        currentReplies,
        willChange: currentSent !== counts.scoutsSent || currentReplies !== counts.scoutReplies,
      };
    });
  }, [result, entriesByDate, selectedMediaId]);

  return (
    <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="media-csv-modal-title">
      <div className="modal-content" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h3 id="media-csv-modal-title">媒体CSVから実績を取り込む</h3>
          <button onClick={onClose} className="close-button" aria-label="閉じる">&times;</button>
        </div>
        <div className="modal-body">
          <p className="modal-description">
            BIZ（ビズリーチ）またはDodaの管理画面からダウンロードした日次実績CSVを取り込み、スカウト数・返信数を反映します。
            既に入力済みの日も、この2項目だけが取り込んだ内容で上書きされ、他の項目は変更されません。
          </p>
          <div className="form-group">
            <label htmlFor="media-csv-select">対象媒体</label>
            <select
              id="media-csv-select"
              value={selectedMediaId}
              onChange={(e) => handleMediaChange(e.target.value as ScoutCsvMediaId)}
            >
              <option value="biz">BIZ（ビズリーチ）</option>
              <option value="doda">Doda</option>
            </select>
          </div>
          <div className="gmail-scout-fetch-bar">
            <input type="file" accept=".csv" onChange={handleFileChange} aria-label="CSVファイルを選択" />
            {status === 'parsing' && <span className="gmail-scout-message">読み込み中...</span>}
            {message && <span className={`gmail-scout-message ${status === 'error' ? 'is-error' : ''}`}>{message}</span>}
          </div>

          {status === 'preview' && result && (
            <div className="bulk-gmail-preview">
              <p className="gmail-scout-message">
                {previewRows.length}日分を検出しました（うち{previewRows.filter(r => r.willChange).length}件が現在の入力値と異なります）。
                反映すると各日の{mediaName}の「スカウト数」「返信数」が上書きされます。
              </p>
              <div className="bulk-gmail-preview-table-container">
                <table className="bulk-gmail-preview-table">
                  <thead>
                    <tr><th>日付</th><th>取得スカウト数</th><th>現在の値</th><th>取得返信数</th><th>現在の値</th></tr>
                  </thead>
                  <tbody>
                    {previewRows.map(row => (
                      <tr key={row.dateStr}>
                        <td>{row.dateStr}</td>
                        <td>{row.scoutsSent}</td>
                        <td className={row.currentSent !== row.scoutsSent ? 'is-changed' : ''}>{row.currentSent}</td>
                        <td>{row.scoutReplies}</td>
                        <td className={row.currentReplies !== row.scoutReplies ? 'is-changed' : ''}>{row.currentReplies}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
        <div className="modal-footer">
          <button type="button" onClick={onClose} className="cancel-button">キャンセル</button>
          {status === 'preview' && result && previewRows.length > 0 && (
            <button type="button" onClick={() => onApply(selectedMediaId, result.countsByDate)} className="submit-button">
              反映する（{previewRows.length}日分）
            </button>
          )}
        </div>
      </div>
    </div>
  );
};


const getTotalFromLump = (lump: { [key: string]: number | undefined }, kpiSuffix: string, allMedia: MediaEntry[]): number => {
    if (!lump) return 0;
    return allMedia.reduce((acc, media) => {
        const kpiKey = `${media.id}${kpiSuffix}`;
        return acc + (lump[kpiKey] || 0);
    }, 0);
};

const CalendarView: React.FC<{
  viewDate: Date;
  entriesByDate: Map<string, KpiTotals>;
  allMedia: MediaEntry[];
  onDayClick: (date: string) => void;
  onPrevMonth: () => void;
  onNextMonth: () => void;
}> = ({ viewDate, entriesByDate, allMedia, onDayClick, onPrevMonth, onNextMonth }) => {
  const year = viewDate.getFullYear();
  const month = viewDate.getMonth();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayLocalString = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  
  const firstDayOfMonth = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  const days = [];
  for (let i = 0; i < firstDayOfMonth; i++) {
    days.push(<div key={`empty-${i}`} className="calendar-day empty"></div>);
  }

  for (let i = 1; i <= daysInMonth; i++) {
    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(i).padStart(2, '0')}`;
    const data = entriesByDate.get(dateStr);
    const isToday = dateStr === todayLocalString;
    const hasData = data && Object.values(data).some(val => val > 0);
    
    const dayClasses = `calendar-day ${isToday ? 'today' : ''} ${hasData ? 'has-data' : ''}`;
    
    const summaryKpis = {
      interviewed: data ? getTotalFromLump(data, '_initialInterviews', allMedia) : 0,
      effectiveInterviewed: data ? getTotalFromLump(data, '_effectiveInitialInterviews', allMedia) : 0,
      submitted: data?.candidatesSubmitted || 0,
      collected: data ? getTotalFromLump(data, '_documentsCollected', allMedia) : 0,
      effectiveCollected: data ? getTotalFromLump(data, '_effectiveDocumentsCollected', allMedia) : 0,
      placed: data?.placements || 0,
    };

    days.push(
      <div key={i} className={dayClasses} onClick={() => onDayClick(dateStr)} role="button" tabIndex={0} aria-label={`${i}日, 実績入力`}>
        <div className="day-number">{i}</div>
        {hasData && (
          <div className="day-kpi-summary">
            {summaryKpis.interviewed > 0 && (
                <div className="day-kpi-item" title={`初回面談: ${summaryKpis.interviewed}`}>
                    <span className="kpi-text-label">初回:</span>
                    <span className="kpi-text-value">{summaryKpis.interviewed}</span>
                </div>
            )}
            {summaryKpis.effectiveInterviewed > 0 && (
                <div className="day-kpi-item" title={`初回有効面談: ${summaryKpis.effectiveInterviewed}`}>
                    <span className="kpi-text-label">有効:</span>
                    <span className="kpi-text-value">{summaryKpis.effectiveInterviewed}</span>
                </div>
            )}
            {summaryKpis.submitted > 0 && (
                <div className="day-kpi-item" title={`候補者推薦: ${summaryKpis.submitted}`}>
                    <span className="kpi-text-label">推薦:</span>
                    <span className="kpi-text-value">{summaryKpis.submitted}</span>
                </div>
            )}
            {summaryKpis.collected > 0 && (
                <div className="day-kpi-item" title={`書類回収: ${summaryKpis.collected}`}>
                    <span className="kpi-text-label">回収:</span>
                    <span className="kpi-text-value">{summaryKpis.collected}</span>
                </div>
            )}
            {summaryKpis.effectiveCollected > 0 && (
                <div className="day-kpi-item" title={`有効書類回収: ${summaryKpis.effectiveCollected}`}>
                    <span className="kpi-text-label">有効回収:</span>
                    <span className="kpi-text-value">{summaryKpis.effectiveCollected}</span>
                </div>
            )}
            {summaryKpis.placed > 0 && (
                <div className="day-kpi-item" title={`内定承諾: ${summaryKpis.placed}`}>
                    <span className="kpi-text-label placed">承諾:</span>
                    <span className="kpi-text-value placed">{summaryKpis.placed}</span>
                </div>
            )}
          </div>
        )}
      </div>
    );
  }

  return (
    <div className="calendar-container">
      <div className="calendar-header">
        <button onClick={onPrevMonth} aria-label="前の月へ">&lt; 前月</button>
        <h2>{`${year}年 ${month + 1}月`}</h2>
        <button onClick={onNextMonth} aria-label="次の月へ">次月 &gt;</button>
      </div>
      <div className="calendar-grid-header">
        {['日', '月', '火', '水', '木', '金', '土'].map(day => <div key={day}>{day}</div>)}
      </div>
      <div className="calendar-grid">
        {days}
      </div>
    </div>
  );
};


const TeamsModal: React.FC<{
    teams: Team[];
    isEditable: boolean;
    isAdmin: boolean;
    authorizedEditorEmails: string[];
    userOptions: { email: string; label: string }[];
    memberDepartments: Record<string, Department>;
    onClose: () => void;
    onCreateTeam: (name: string) => void;
    onRenameTeam: (teamId: string, name: string) => void;
    onDeleteTeam: (teamId: string) => void;
    onAddMember: (teamId: string, email: string) => void;
    onRemoveMember: (teamId: string, email: string) => void;
    onGrantEditor: (email: string) => void;
    onRevokeEditor: (email: string) => void;
    onSetMemberDepartment: (email: string, department: Department | null) => void;
}> = ({ teams, isEditable, isAdmin, authorizedEditorEmails, userOptions, memberDepartments, onClose, onCreateTeam, onRenameTeam, onDeleteTeam, onAddMember, onRemoveMember, onGrantEditor, onRevokeEditor, onSetMemberDepartment }) => {
    const [newTeamName, setNewTeamName] = useState('');
    const [editingTeamId, setEditingTeamId] = useState<string | null>(null);
    const [editedName, setEditedName] = useState('');
    const [memberInputs, setMemberInputs] = useState<Record<string, string>>({});
    const [newEditorEmail, setNewEditorEmail] = useState('');

    const labelByEmail = useMemo(() => new Map(userOptions.map(u => [u.email, u.label])), [userOptions]);

    const handleGrantEditor = (e: React.FormEvent) => {
        e.preventDefault();
        const email = newEditorEmail.trim();
        if (!email) return;
        if (!email.toLowerCase().endsWith('@bloom-firm.com')) {
            alert('bloom-firm.com のメールアドレスを入力してください。');
            return;
        }
        onGrantEditor(email);
        setNewEditorEmail('');
    };

    const handleCreate = (e: React.FormEvent) => {
        e.preventDefault();
        if (!newTeamName.trim()) return;
        onCreateTeam(newTeamName.trim());
        setNewTeamName('');
    };

    const handleStartEdit = (team: Team) => {
        setEditingTeamId(team.id);
        setEditedName(team.name);
    };

    const handleSaveEdit = (teamId: string) => {
        if (editedName.trim()) onRenameTeam(teamId, editedName.trim());
        setEditingTeamId(null);
    };

    const handleAddMember = (teamId: string, emailOverride?: string) => {
        // Lowercased before storing — this is free-typed (or pasted) input, not picked from a
        // list of known signed-in accounts, so a mismatched case here would silently make this
        // member invisible to every scope that indexes allUsersData directly by this string
        // (Google's own email casing is otherwise always consistent, since that comes straight
        // from account data rather than someone's keyboard).
        const email = (emailOverride ?? memberInputs[teamId] ?? '').trim().toLowerCase();
        if (!email) return;
        if (!email.endsWith('@bloom-firm.com')) {
            alert('bloom-firm.com のメールアドレスを入力してください。');
            return;
        }
        onAddMember(teamId, email);
        setMemberInputs(prev => ({ ...prev, [teamId]: '' }));
    };

    return (
        <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="teams-modal-title">
            <div className="modal-content" onClick={e => e.stopPropagation()}>
                <div className="modal-header">
                    <h3 id="teams-modal-title">チーム管理</h3>
                    <button onClick={onClose} className="close-button" aria-label="閉じる">&times;</button>
                </div>
                <div className="modal-body">
                    {isAdmin && (
                        <div className="teams-permission-section">
                            <h4 className="sub-section-title">チーム作成・編集権限の管理</h4>
                            <p className="modal-description">
                                ここに登録したメールアドレスの人は「チーム管理」からチームの作成・編集ができるようになります。
                            </p>
                            {authorizedEditorEmails.length === 0 ? (
                                <p className="no-data-message">まだ誰にも権限を付与していません。</p>
                            ) : (
                                <ul className="user-management-list">
                                    {authorizedEditorEmails.map(email => (
                                        <li key={email} className="user-management-item">
                                            <span className="user-management-name">
                                                {labelByEmail.get(email) || email}
                                                {labelByEmail.has(email) && <small style={{ color: 'var(--text-muted-color)' }}> ({email})</small>}
                                            </span>
                                            <button onClick={() => onRevokeEditor(email)} className="delete-user-button">削除</button>
                                        </li>
                                    ))}
                                </ul>
                            )}
                            <form onSubmit={handleGrantEditor} style={{ display: 'flex', gap: '0.5rem', marginTop: '0.5rem' }}>
                                <input
                                    type="email"
                                    value={newEditorEmail}
                                    onChange={(e) => setNewEditorEmail(e.target.value)}
                                    placeholder="example@bloom-firm.com"
                                    style={{ flex: 1 }}
                                />
                                <button type="submit" className="submit-button" disabled={!newEditorEmail.trim()}>権限を付与</button>
                            </form>
                            <hr style={{ margin: '1rem 0' }} />
                        </div>
                    )}
                    {isEditable && (
                        <div className="teams-permission-section">
                            <h4 className="sub-section-title">メンバーの所属部署</h4>
                            <p className="modal-description">
                                各メンバーがF+（Firm+）・AC（AssetCareer）のどちらに所属するかを設定します。ヘッダーの事業部切り替えで表示が絞り込まれます。
                            </p>
                            {userOptions.length === 0 ? (
                                <p className="no-data-message">まだユーザーデータがありません。</p>
                            ) : (
                                <ul className="user-management-list">
                                    {userOptions.map(u => (
                                        <li key={u.email} className="user-management-item">
                                            <span className="user-management-name">
                                                {u.label}
                                                {u.label !== u.email && <small style={{ color: 'var(--text-muted-color)' }}> ({u.email})</small>}
                                            </span>
                                            <select
                                                value={memberDepartments[u.email] || ''}
                                                onChange={(e) => onSetMemberDepartment(u.email, (e.target.value || null) as Department | null)}
                                            >
                                                <option value="">未設定</option>
                                                <option value="F+">Firm+</option>
                                                <option value="AC">AssetCareer</option>
                                            </select>
                                        </li>
                                    ))}
                                </ul>
                            )}
                            <hr style={{ margin: '1rem 0' }} />
                        </div>
                    )}
                    {!isEditable && (
                        <p className="no-data-message">
                            チームの作成・編集は権限を付与された人のみ可能です。閲覧のみできます。
                        </p>
                    )}
                    {isEditable && (
                        <form onSubmit={handleCreate} className="form-group" style={{ marginBottom: '1rem' }}>
                            <label htmlFor="new-team-name">新しいチーム名</label>
                            <input
                                id="new-team-name"
                                type="text"
                                value={newTeamName}
                                onChange={(e) => setNewTeamName(e.target.value)}
                                placeholder="例: 営業第一チーム"
                            />
                            <button type="submit" className="submit-button" disabled={!newTeamName.trim()} style={{ marginTop: '0.5rem' }}>
                                チームを作成
                            </button>
                        </form>
                    )}
                    {teams.length === 0 ? (
                        <p className="no-data-message">まだチームがありません。</p>
                    ) : (
                        <ul className="user-management-list">
                            {teams.map(team => (
                                <li key={team.id} className="user-management-item" style={{ flexDirection: 'column', alignItems: 'stretch' }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', width: '100%' }}>
                                        {editingTeamId === team.id ? (
                                            <input
                                                type="text"
                                                value={editedName}
                                                onChange={(e) => setEditedName(e.target.value)}
                                                className="user-management-input"
                                                autoFocus
                                                onKeyDown={(e) => e.key === 'Enter' && handleSaveEdit(team.id)}
                                            />
                                        ) : (
                                            <span className="user-management-name">{team.name}</span>
                                        )}
                                        {isEditable && (
                                            <div className="user-management-actions">
                                                {editingTeamId === team.id ? (
                                                    <>
                                                        <button onClick={() => handleSaveEdit(team.id)} className="save-user-button">保存</button>
                                                        <button onClick={() => setEditingTeamId(null)} className="cancel-user-button">キャンセル</button>
                                                    </>
                                                ) : (
                                                    <>
                                                        <button onClick={() => handleStartEdit(team)} className="edit-user-button">編集</button>
                                                        <button
                                                            onClick={() => window.confirm(`「${team.name}」を削除しますか？`) && onDeleteTeam(team.id)}
                                                            className="delete-user-button"
                                                        >
                                                            削除
                                                        </button>
                                                    </>
                                                )}
                                            </div>
                                        )}
                                    </div>
                                    <ul style={{ marginTop: '0.5rem', paddingLeft: '1rem' }}>
                                        {team.memberEmails.map(email => (
                                            <li key={email} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '0.25rem 0' }}>
                                                <span>
                                                    {labelByEmail.get(email) || email}
                                                    {labelByEmail.has(email) && <small style={{ color: 'var(--text-muted-color)' }}> ({email})</small>}
                                                </span>
                                                {isEditable && (
                                                    <button onClick={() => onRemoveMember(team.id, email)} className="delete-user-button">削除</button>
                                                )}
                                            </li>
                                        ))}
                                    </ul>
                                    {isEditable && (
                                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.5rem', marginTop: '0.5rem', alignItems: 'center' }}>
                                            <select
                                                value=""
                                                onChange={(e) => { if (e.target.value) handleAddMember(team.id, e.target.value); }}
                                                aria-label="登録済みユーザーから選択して追加"
                                            >
                                                <option value="">登録済みユーザーから選択...</option>
                                                {userOptions
                                                    .filter(u => !team.memberEmails.includes(u.email))
                                                    .map(u => <option key={u.email} value={u.email}>{u.label}</option>)}
                                            </select>
                                            <span style={{ color: 'var(--text-muted-color)', fontSize: '0.85rem' }}>または</span>
                                            <input
                                                type="email"
                                                placeholder="member@bloom-firm.com"
                                                value={memberInputs[team.id] || ''}
                                                onChange={(e) => setMemberInputs(prev => ({ ...prev, [team.id]: e.target.value }))}
                                                onKeyDown={(e) => e.key === 'Enter' && handleAddMember(team.id)}
                                            />
                                            <button onClick={() => handleAddMember(team.id)} className="submit-button">追加</button>
                                        </div>
                                    )}
                                </li>
                            ))}
                        </ul>
                    )}
                </div>
                <div className="modal-footer">
                    <button type="button" onClick={onClose} className="cancel-button">閉じる</button>
                </div>
            </div>
        </div>
    );
};

const MediaModal: React.FC<{
    allMedia: MediaEntry[];
    isEditable: boolean;
    onClose: () => void;
    onCreateMedia: (name: string) => void;
    onRenameMedia: (id: string, name: string) => void;
    onArchiveMedia: (id: string) => void;
    onUnarchiveMedia: (id: string) => void;
    onSetFeeRate: (id: string, feeRate: number | undefined) => void;
    onRefresh: () => void;
}> = ({ allMedia, isEditable, onClose, onCreateMedia, onRenameMedia, onArchiveMedia, onUnarchiveMedia, onSetFeeRate, onRefresh }) => {
    const [newMediaName, setNewMediaName] = useState('');
    const [editingMediaId, setEditingMediaId] = useState<string | null>(null);
    const [editedName, setEditedName] = useState('');
    const [feeRateInputs, setFeeRateInputs] = useState<Record<string, string>>({});

    const handleFeeRateChange = (id: string, value: string) => {
        setFeeRateInputs(prev => ({ ...prev, [id]: value }));
    };

    const handleFeeRateBlur = (id: string) => {
        const raw = feeRateInputs[id];
        if (raw === undefined) return;
        const trimmed = raw.trim();
        onSetFeeRate(id, trimmed === '' ? undefined : Number(trimmed));
    };

    const activeMedia = allMedia.filter(m => !m.isArchived);
    const archivedMedia = allMedia.filter(m => m.isArchived);

    const handleCreate = (e: React.FormEvent) => {
        e.preventDefault();
        if (!newMediaName.trim()) return;
        onCreateMedia(newMediaName.trim());
        setNewMediaName('');
    };

    const handleStartEdit = (media: MediaEntry) => {
        setEditingMediaId(media.id);
        setEditedName(media.name);
    };

    const handleSaveEdit = (id: string) => {
        if (editedName.trim()) onRenameMedia(id, editedName.trim());
        setEditingMediaId(null);
    };

    return (
        <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="media-modal-title">
            <div className="modal-content" onClick={e => e.stopPropagation()}>
                <div className="modal-header">
                    <h3 id="media-modal-title">媒体管理</h3>
                    <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
                        <button onClick={onRefresh} className="secondary-action-button">更新</button>
                        <button onClick={onClose} className="close-button" aria-label="閉じる">&times;</button>
                    </div>
                </div>
                <div className="modal-body">
                    {!isEditable && (
                        <p className="no-data-message">
                            媒体設定の編集は管理者（{MEDIA_ADMIN_EMAIL}）のみ可能です。閲覧のみできます。
                        </p>
                    )}
                    {isEditable && (
                        <form onSubmit={handleCreate} className="form-group" style={{ marginBottom: '1rem' }}>
                            <label htmlFor="new-media-name">新しい媒体名</label>
                            <input
                                id="new-media-name"
                                type="text"
                                value={newMediaName}
                                onChange={(e) => setNewMediaName(e.target.value)}
                                placeholder="例: リクナビNEXT"
                            />
                            <button type="submit" className="submit-button" disabled={!newMediaName.trim()} style={{ marginTop: '0.5rem' }}>
                                媒体を追加
                            </button>
                        </form>
                    )}

                    <h4 className="sub-section-title">利用中の媒体</h4>
                    {activeMedia.length === 0 ? (
                        <p className="no-data-message">媒体がありません。</p>
                    ) : (
                        <ul className="user-management-list">
                            {activeMedia.map(media => (
                                <li key={media.id} className="user-management-item" style={{ flexDirection: 'column', alignItems: 'stretch' }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', width: '100%' }}>
                                        {editingMediaId === media.id ? (
                                            <input
                                                type="text"
                                                value={editedName}
                                                onChange={(e) => setEditedName(e.target.value)}
                                                className="user-management-input"
                                                autoFocus
                                                onKeyDown={(e) => e.key === 'Enter' && handleSaveEdit(media.id)}
                                            />
                                        ) : (
                                            <span className="user-management-name">{media.name}</span>
                                        )}
                                        {isEditable && (
                                            <div className="user-management-actions">
                                                {editingMediaId === media.id ? (
                                                    <>
                                                        <button onClick={() => handleSaveEdit(media.id)} className="save-user-button">保存</button>
                                                        <button onClick={() => setEditingMediaId(null)} className="cancel-user-button">キャンセル</button>
                                                    </>
                                                ) : (
                                                    <>
                                                        <button onClick={() => handleStartEdit(media)} className="edit-user-button">改称</button>
                                                        <button
                                                            onClick={() => window.confirm(`「${media.name}」をアーカイブします。過去の実績データは保持されますが、入力フォームには表示されなくなります。よろしいですか？`) && onArchiveMedia(media.id)}
                                                            className="delete-user-button"
                                                        >
                                                            アーカイブ
                                                        </button>
                                                    </>
                                                )}
                                            </div>
                                        )}
                                    </div>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginTop: '0.5rem' }}>
                                        <label htmlFor={`media-fee-rate-${media.id}`} style={{ fontSize: '0.85rem', color: 'var(--text-muted-color)' }}>手数料率(%):</label>
                                        <input
                                            id={`media-fee-rate-${media.id}`}
                                            type="number"
                                            min="0"
                                            step="0.1"
                                            style={{ width: '5rem' }}
                                            placeholder="未設定"
                                            value={feeRateInputs[media.id] ?? (media.feeRate ?? '')}
                                            onChange={(e) => handleFeeRateChange(media.id, e.target.value)}
                                            onBlur={() => handleFeeRateBlur(media.id)}
                                            disabled={!isEditable}
                                        />
                                    </div>
                                </li>
                            ))}
                        </ul>
                    )}

                    {archivedMedia.length > 0 && (
                        <>
                            <h4 className="sub-section-title">アーカイブ済みの媒体</h4>
                            <ul className="user-management-list">
                                {archivedMedia.map(media => (
                                    <li key={media.id} className="user-management-item">
                                        <span className="user-management-name">{media.name}</span>
                                        {isEditable && (
                                            <div className="user-management-actions">
                                                <button onClick={() => onUnarchiveMedia(media.id)} className="edit-user-button">復元</button>
                                            </div>
                                        )}
                                    </li>
                                ))}
                            </ul>
                        </>
                    )}
                </div>
                <div className="modal-footer">
                    <button type="button" onClick={onClose} className="cancel-button">閉じる</button>
                </div>
            </div>
        </div>
    );
};


const DailyProgressCard: React.FC<{
    source: MediaEntry;
    todayTotals: KpiTotals;
    dailyKpiTargets: Record<KpiKey, number>;
}> = ({ source, todayTotals, dailyKpiTargets }) => {
    const sourceKey = source.id;
    const scoutsKey = `${sourceKey}_scoutsSent` as KpiKey;
    const repliesKey = `${sourceKey}_scoutReplies` as KpiKey;
    const effectiveRepliesKey = `${sourceKey}_effectiveReplies` as KpiKey;

    const scoutsSent = todayTotals[scoutsKey] || 0;
    const scoutReplies = todayTotals[repliesKey] || 0;
    const effectiveReplies = todayTotals[effectiveRepliesKey] || 0;
    const scoutsTarget = dailyKpiTargets[scoutsKey] || 0;
    const repliesTarget = dailyKpiTargets[repliesKey] || 0;
    const effectiveRepliesTarget = dailyKpiTargets[effectiveRepliesKey] || 0;

    const replyRate = scoutsSent > 0 ? (scoutReplies / scoutsSent) * 100 : 0;
    const effectiveReplyRate = scoutReplies > 0 ? (effectiveReplies / scoutReplies) * 100 : 0;

    const scoutsProgress = scoutsTarget > 0 ? Math.min((scoutsSent / scoutsTarget) * 100, 100) : 0;
    const repliesProgress = repliesTarget > 0 ? Math.min((scoutReplies / repliesTarget) * 100, 100) : 0;
    const effectiveRepliesProgress = effectiveRepliesTarget > 0 ? Math.min((effectiveReplies / effectiveRepliesTarget) * 100, 100) : 0;

    return (
        <div className="daily-progress-card">
            <h3>{source.name}</h3>
            <div className="reply-rate-group">
                <div className="reply-rate-section" style={{ borderBottom: 'none' }}>
                    <span className="reply-rate-value" style={{ fontSize: '1.75rem' }}>{replyRate.toFixed(1)}%</span>
                    <span className="reply-rate-label">返信率</span>
                </div>
                <div className="reply-rate-section" style={{ borderBottom: 'none' }}>
                    <span className="reply-rate-value" style={{ fontSize: '1.75rem' }}>{effectiveReplyRate.toFixed(1)}%</span>
                    <span className="reply-rate-label">有効返信率</span>
                </div>
            </div>
            <div className="stat-item">
                <div className="stat-details">
                    <span>スカウト数</span>
                    <span>{scoutsSent} / {scoutsTarget}</span>
                </div>
                <div className="progress-bar">
                    <div className="progress-bar-fill" style={{ width: `${scoutsProgress}%` }}></div>
                </div>
            </div>
            <div className="stat-item">
                <div className="stat-details">
                    <span>スカウト返信数</span>
                    <span>{scoutReplies} / {repliesTarget}</span>
                </div>
                 <div className="progress-bar">
                    <div className="progress-bar-fill" style={{ width: `${repliesProgress}%` }}></div>
                </div>
            </div>
            <div className="stat-item">
                <div className="stat-details">
                    <span>有効返信数</span>
                    <span>{effectiveReplies} / {effectiveRepliesTarget}</span>
                </div>
                 <div className="progress-bar">
                    <div className="progress-bar-fill" style={{ width: `${effectiveRepliesProgress}%`, backgroundColor: 'var(--info-color)' }}></div>
                </div>
            </div>
        </div>
    );
};

const DailyProgress: React.FC<{
    activeMedia: MediaEntry[];
    todayTotals: KpiTotals;
    dailyKpiTargets: Record<KpiKey, number>;
}> = ({ activeMedia, todayTotals, dailyKpiTargets }) => {
    const hasTargets = activeMedia.some(source => {
         const sourceKey = source.id;
         const scoutsKey = `${sourceKey}_scoutsSent` as KpiKey;
         const repliesKey = `${sourceKey}_scoutReplies` as KpiKey;
         const effectiveRepliesKey = `${sourceKey}_effectiveReplies` as KpiKey;
         return (dailyKpiTargets[scoutsKey] || 0) > 0
            || (dailyKpiTargets[repliesKey] || 0) > 0
            || (dailyKpiTargets[effectiveRepliesKey] || 0) > 0;
    });

    if (!hasTargets) {
        return <p className="no-data-message">日次目標が設定されていません。まずは「日次目標設定」から目標を入力してください。</p>;
    }

    return (
        <div className="daily-progress-container">
            {activeMedia.map(source => (
                <DailyProgressCard
                    key={source.id}
                    source={source}
                    todayTotals={todayTotals}
                    dailyKpiTargets={dailyKpiTargets}
                />
            ))}
        </div>
    );
};


const CandidateModal: React.FC<{
    onSave: (candidate: Candidate) => void;
    onClose: () => void;
    initialData?: Candidate | null;
    allMedia: MediaEntry[];
}> = ({ onSave, onClose, initialData, allMedia }) => {
    const activeMedia = allMedia.filter(m => !m.isArchived);
    const defaultCandidate: Candidate = {
        id: initialData?.id || `candidate_${Date.now()}`,
        name: '',
        salary: 0,
        currentSalary: 0,
        expectedAnnualSalary: 0,
        currentCompany: '',
        education: '',
        source: '',
        usingOtherAgents: false,
        applications: [],
        summary: '',
        resumeFiles: [],
        interviewAudioFile: null,
        interviewSummary: '',
        isHidden: false,
        createdAt: initialData?.createdAt || new Date().toISOString(),
    };

    const [candidate, setCandidate] = useState<Candidate>(() => {
        const initial = initialData || defaultCandidate;
        const hasOtherAgents = (initial as any).otherAgents && Array.isArray((initial as any).otherAgents) && (initial as any).otherAgents.length > 0;
        const resumeFiles = initial.resumeFiles || (initial.resumeFile ? [{ name: (initial.resumeFile as any).name }] : []);

        return {
            ...initial,
            currentSalary: initial.currentSalary || 0,
            expectedAnnualSalary: initial.expectedAnnualSalary || 0,
            currentCompany: initial.currentCompany || '',
            education: initial.education || '',
            usingOtherAgents: initial.usingOtherAgents || hasOtherAgents || false,
            isHidden: initial.isHidden || false,
            resumeFiles: resumeFiles,
            interviewAudioFile: initial.interviewAudioFile || null,
            interviewSummary: initial.interviewSummary || '',
        };
    });
    
    const [resumeDragActive, setResumeDragActive] = useState(false);
    const [audioDragActive, setAudioDragActive] = useState(false);
    const [isGenerating, setIsGenerating] = useState(false);
    const [isGeneratingInterviewSummary, setIsGeneratingInterviewSummary] = useState(false);
    const [isSearchingInterviewLogs, setIsSearchingInterviewLogs] = useState(false);
    const [interviewLogResults, setInterviewLogResults] = useState<InterviewLogFile[] | null>(null);
    const [isSummarizingInterviewLog, setIsSummarizingInterviewLog] = useState(false);
    const [interviewLogDragActive, setInterviewLogDragActive] = useState(false);
    const resumeInputRef = useRef<HTMLInputElement>(null);
    const audioInputRef = useRef<HTMLInputElement>(null);
    const interviewLogInputRef = useRef<HTMLInputElement>(null);

    const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
        const { name, value } = e.target;
        const numericFields = ['salary', 'currentSalary', 'expectedAnnualSalary'];
        setCandidate(prev => ({ ...prev, [name]: numericFields.includes(name) ? Number(value) : value }));
    };

    const handleApplicationChange = (index: number, field: keyof CompanyApplication, value: string) => {
        const newApplications = [...candidate.applications];
        const parsedValue: string | number | undefined = field === 'feeRate'
            ? (value === '' ? undefined : Number(value))
            : value;
        newApplications[index] = { ...newApplications[index], [field]: parsedValue };
        setCandidate(prev => ({ ...prev, applications: newApplications }));
    };

    const addApplication = () => {
        const newApp: CompanyApplication = { id: `app_${Date.now()}`, companyName: '', stage: '打診', nextAction: '' };
        setCandidate(prev => ({ ...prev, applications: [...prev.applications, newApp] }));
    };

    const removeApplication = (index: number) => {
        setCandidate(prev => ({ ...prev, applications: prev.applications.filter((_, i) => i !== index) }));
    };

    const handleResumeDrag = (e: React.DragEvent) => {
        e.preventDefault();
        e.stopPropagation();
        if (e.type === "dragenter" || e.type === "dragover") {
            setResumeDragActive(true);
        } else if (e.type === "dragleave") {
            setResumeDragActive(false);
        }
    };
    
    const fileToBase64 = (file: File): Promise<string> => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => resolve((reader.result as string).split(',')[1]);
            reader.onerror = error => reject(error);
        });
    };

    const handleResumeFile = async (files: FileList) => {
      if (!files || files.length === 0) return;

      const pdfFiles = Array.from(files).filter(file => file.type === 'application/pdf');
      if (pdfFiles.length === 0) {
          alert('PDFファイルのみアップロードできます。');
          return;
      }

      const existingFileNames = new Set(candidate.resumeFiles?.map(f => f.name) || []);
      const newFiles = pdfFiles.filter(file => !existingFileNames.has(file.name));

      if (newFiles.length === 0) {
        alert('アップロードしようとしたファイルは既に追加されています。');
        return;
      }
      
      setCandidate(prev => ({ ...prev, resumeFiles: [...(prev.resumeFiles || []), ...newFiles.map(f => ({ name: f.name }))] }));
      setIsGenerating(true);
      try {
          const ai = new GoogleGenAI({ apiKey: process.env.API_KEY as string });
          
          const fileParts = await Promise.all(newFiles.map(async (file) => {
              const base64Data = await fileToBase64(file);
              return {
                  inlineData: {
                      mimeType: file.type,
                      data: base64Data,
                  },
              };
          }));

          const textPart = {
              text: 'これらの履歴書から候補者の氏名、現在の勤務先企業名、最終学歴、そして学歴、職歴、スキルを重視した概要を抽出してください。複数のファイルがある場合は、情報を集約し、最も包括的な内容にしてください。',
          };

          const response = await ai.models.generateContent({
              model: 'gemini-2.5-flash',
              contents: { parts: [...fileParts, textPart] },
              config: {
                responseMimeType: "application/json",
                responseSchema: {
                    type: Type.OBJECT,
                    properties: {
                        name: {
                            type: Type.STRING,
                            description: '候補者の氏名',
                        },
                        currentCompany: {
                            type: Type.STRING,
                            description: '候補者の現在の勤務先企業名。見つからない場合は空文字を返す。',
                        },
                        education: {
                            type: Type.STRING,
                            description: '候補者の最終学歴。見つからない場合は空文字を返す。',
                        },
                        summary: {
                            type: Type.STRING,
                            description: '候補者の学歴、職歴、スキル、職務概要に焦点を当てた簡潔な要約。採用担当者が候補者の全体像を素早く把握できるように記述する。',
                        },
                    },
                    required: ["name", "summary", "currentCompany", "education"],
                },
              }
          });
          
          const jsonStr = response.text.trim();
          const parsedData = JSON.parse(jsonStr);

          setCandidate(prev => ({
              ...prev,
              name: parsedData.name?.trim() || prev.name,
              currentCompany: parsedData.currentCompany?.trim() || prev.currentCompany,
              education: parsedData.education?.trim() || prev.education,
              summary: parsedData.summary?.trim() || prev.summary,
          }));

      } catch (error) {
          console.error("Error generating summary:", error);
          alert('AIによる要約の生成中にエラーが発生しました。');
      } finally {
          setIsGenerating(false);
      }
    };

    const handleResumeDrop = (e: React.DragEvent) => {
        e.preventDefault();
        e.stopPropagation();
        setResumeDragActive(false);
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            handleResumeFile(e.dataTransfer.files);
        }
    };
    
    const handleResumeFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        e.preventDefault();
        if (e.target.files && e.target.files.length > 0) {
            handleResumeFile(e.target.files);
        }
    };

    const handleRemoveResumeFile = (fileNameToRemove: string) => {
        setCandidate(prev => ({
            ...prev,
            resumeFiles: prev.resumeFiles?.filter(file => file.name !== fileNameToRemove),
        }));
    };
    
    const onResumeButtonClick = () => {
        resumeInputRef.current?.click();
    };
    
    // --- Audio file handlers ---
    const handleAudioDrag = (e: React.DragEvent) => {
        e.preventDefault();
        e.stopPropagation();
        if (e.type === "dragenter" || e.type === "dragover") {
            setAudioDragActive(true);
        } else if (e.type === "dragleave") {
            setAudioDragActive(false);
        }
    };

    const handleAudioFile = async (files: FileList) => {
        if (!files || files.length === 0) return;

        const audioFile = Array.from(files).find(file => file.type.startsWith('audio/'));
        if (!audioFile) {
            alert('音声ファイル（MP3, WAV, etc.）のみアップロードできます。');
            return;
        }
        
        if (candidate.interviewAudioFile) {
            if (!window.confirm('既存の音声ファイルを置き換えますか？')) {
                return;
            }
        }

        setCandidate(prev => ({ ...prev, interviewAudioFile: { name: audioFile.name }, interviewSummary: '' }));
        setIsGeneratingInterviewSummary(true);

        try {
            const ai = new GoogleGenAI({ apiKey: process.env.API_KEY as string });

            const base64Data = await fileToBase64(audioFile);
            const audioPart = {
                inlineData: {
                    mimeType: audioFile.type,
                    data: base64Data,
                },
            };

            const textPart = {
                text: 'この面談の音声データの内容を、シンプルに要約してください。',
            };

            const response = await ai.models.generateContent({
                model: 'gemini-2.5-flash',
                contents: { parts: [audioPart, textPart] },
            });

            setCandidate(prev => ({
                ...prev,
                interviewSummary: response.text.trim(),
            }));

        } catch (error) {
            console.error("Error generating interview summary:", error);
            alert('AIによる面談要約の生成中にエラーが発生しました。');
            setCandidate(prev => ({ ...prev, interviewAudioFile: null, interviewSummary: '' }));
        } finally {
            setIsGeneratingInterviewSummary(false);
        }
    };

    const handleAudioDrop = (e: React.DragEvent) => {
        e.preventDefault();
        e.stopPropagation();
        setAudioDragActive(false);
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            handleAudioFile(e.dataTransfer.files);
        }
    };

    const handleAudioFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        e.preventDefault();
        if (e.target.files && e.target.files.length > 0) {
            handleAudioFile(e.target.files);
        }
    };

    const handleRemoveAudioFile = () => {
        setCandidate(prev => ({
            ...prev,
            interviewAudioFile: null,
            interviewSummary: '',
        }));
    };

    const onAudioButtonClick = () => {
        audioInputRef.current?.click();
    };

    const handleSearchInterviewLogs = async () => {
        if (!candidate.name.trim()) {
            alert('先に候補者名を入力してください。');
            return;
        }
        setIsSearchingInterviewLogs(true);
        setInterviewLogResults(null);
        try {
            const results = await searchInterviewLogsByName(candidate.name.trim());
            setInterviewLogResults(results);
            if (results.length === 0) {
                alert('Googleドライブ内に該当する面談ログ（Google Meetの議事録）が見つかりませんでした。');
            }
        } catch (error) {
            console.error('Error searching interview logs:', error);
            alert('面談ログの検索中にエラーが発生しました。');
        } finally {
            setIsSearchingInterviewLogs(false);
        }
    };

    // Shared by both the Drive-search result path and the manual-file-drop fallback below.
    // `contents` is whatever GoogleGenAI.generateContent accepts — a plain text prompt for
    // Drive-fetched/plain-text logs, or a multimodal {parts} array (inlineData + instruction)
    // for PDF/Word files, letting Gemini read the document itself instead of this app trying to
    // extract text from those formats client-side.
    const summarizeAndAppendInterviewLog = async (contents: any, sourceLabel: string, dateLabel: string) => {
        setIsSummarizingInterviewLog(true);
        try {
            const ai = new GoogleGenAI({ apiKey: process.env.API_KEY as string });
            const response = await ai.models.generateContent({ model: 'gemini-2.5-flash', contents });
            const entry = `--- 面談ログ「${sourceLabel}」(${dateLabel}) より ---\n${response.text.trim()}`;
            setCandidate(prev => ({
                ...prev,
                interviewSummary: prev.interviewSummary ? `${prev.interviewSummary}\n\n${entry}` : entry,
            }));
        } catch (error) {
            console.error('Error summarizing interview log:', error);
            alert('面談ログの要約生成中にエラーが発生しました。');
        } finally {
            setIsSummarizingInterviewLog(false);
        }
    };

    const handleUseInterviewLog = async (file: InterviewLogFile) => {
        try {
            const text = await exportGoogleDocAsText(file.id);
            const dateLabel = new Date(file.modifiedTime).toLocaleDateString('ja-JP');
            await summarizeAndAppendInterviewLog(
                `以下はGoogle Meetの面談議事録です。この内容をシンプルに要約してください。\n\n---\n${text}`,
                file.name,
                dateLabel
            );
            setInterviewLogResults(null);
        } catch (error) {
            console.error('Error fetching interview log from Drive:', error);
            alert('面談ログの取得中にエラーが発生しました。');
        }
    };

    // Fallback for when the automatic Drive search doesn't find the log (e.g. it's not a native
    // Google Doc, or search just misses it) — lets the user drop/select the transcript file
    // directly instead. Plain text is read client-side; PDF/Word files are instead sent to
    // Gemini as inline document data (same approach already used for resume PDFs above) since
    // this app has no PDF/DOCX text-extraction library of its own.
    const handleInterviewLogFile = async (files: FileList) => {
        if (!files || files.length === 0) return;
        const file = files[0];
        const dateLabel = new Date(file.lastModified).toLocaleDateString('ja-JP');
        try {
            if (file.type === 'text/plain' || file.name.toLowerCase().endsWith('.txt')) {
                const text = await file.text();
                await summarizeAndAppendInterviewLog(
                    `以下はGoogle Meetの面談議事録です。この内容をシンプルに要約してください。\n\n---\n${text}`,
                    file.name,
                    dateLabel
                );
            } else {
                const base64Data = await fileToBase64(file);
                await summarizeAndAppendInterviewLog(
                    {
                        parts: [
                            { inlineData: { mimeType: file.type, data: base64Data } },
                            { text: 'これはGoogle Meetなどの面談議事録のファイルです。この内容をシンプルに要約してください。' },
                        ],
                    },
                    file.name,
                    dateLabel
                );
            }
        } catch (error) {
            console.error('Error reading dropped interview log file:', error);
            alert('ファイルの読み込み中にエラーが発生しました。');
        }
    };

    const handleInterviewLogDrag = (e: React.DragEvent) => {
        e.preventDefault();
        e.stopPropagation();
        if (e.type === 'dragenter' || e.type === 'dragover') {
            setInterviewLogDragActive(true);
        } else if (e.type === 'dragleave') {
            setInterviewLogDragActive(false);
        }
    };

    const handleInterviewLogDrop = (e: React.DragEvent) => {
        e.preventDefault();
        e.stopPropagation();
        setInterviewLogDragActive(false);
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            handleInterviewLogFile(e.dataTransfer.files);
        }
    };

    const handleInterviewLogFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        e.preventDefault();
        if (e.target.files && e.target.files.length > 0) {
            handleInterviewLogFile(e.target.files);
        }
        e.target.value = '';
    };

    const onInterviewLogButtonClick = () => {
        interviewLogInputRef.current?.click();
    };


    const handleSubmit = (e: React.FormEvent) => {
        e.preventDefault();
        if (!candidate.name) {
            alert('候補者名は必須です。');
            return;
        }
        onSave(candidate);
        onClose();
    };
    
    const title = initialData ? '候補者情報を編集' : '新規候補者を追加';

    return (
      <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="candidate-modal-title">
        <div className="modal-content candidate-modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
                <h3 id="candidate-modal-title">{title}</h3>
                <button onClick={onClose} className="close-button" aria-label="閉じる">&times;</button>
            </div>
            <form id="candidate-form" className="modal-body" onSubmit={handleSubmit}>
              <div className="form-group-grid">
                  <div className="form-group form-group-span-3">
                      <label>レジュメ (PDF) - 複数可</label>
                      <p className="form-helper-text">PDFをアップロードすると、氏名、現職、学歴、概要が自動入力されます。複数ファイルを追加すると情報は統合されます。</p>
                       <div 
                          id="resume-drop-zone" 
                          className={`drop-zone ${resumeDragActive ? 'drag-active' : ''}`} 
                          onDragEnter={handleResumeDrag} 
                          onDragLeave={handleResumeDrag} 
                          onDragOver={handleResumeDrag} 
                          onDrop={handleResumeDrop}
                        >
                          <input ref={resumeInputRef} type="file" id="resume-upload" accept="application/pdf" onChange={handleResumeFileChange} multiple />
                          {candidate.resumeFiles && candidate.resumeFiles.length > 0 ? (
                              <div className="file-list">
                                  {candidate.resumeFiles.map(file => (
                                      <div key={file.name} className="file-item">
                                          <span>{file.name}</span>
                                          <button type="button" onClick={() => handleRemoveResumeFile(file.name)} className="remove-file-button" aria-label={`${file.name}を削除`}>&times;</button>
                                      </div>
                                  ))}
                              </div>
                          ) : (
                              <p onClick={onResumeButtonClick}>ここにPDFファイルをドラッグ＆ドロップ、またはクリックして選択</p>
                          )}
                      </div>
                  </div>
                  <div className="form-group form-group-span-3">
                    <label>面談音声ファイル</label>
                    <p className="form-helper-text">面談の音声ファイル（MP3, WAV等）をアップロードすると、AIが面談内容を要約します。</p>
                    <div
                      id="audio-drop-zone"
                      className={`drop-zone ${audioDragActive ? 'drag-active' : ''}`}
                      onDragEnter={handleAudioDrag}
                      onDragLeave={handleAudioDrag}
                      onDragOver={handleAudioDrag}
                      onDrop={handleAudioDrop}
                    >
                      <input ref={audioInputRef} type="file" id="audio-upload" accept="audio/*" onChange={handleAudioFileChange} />
                      {candidate.interviewAudioFile ? (
                        <div className="file-list">
                          <div className="file-item">
                            <span>{candidate.interviewAudioFile.name}</span>
                            <button type="button" onClick={handleRemoveAudioFile} className="remove-file-button" aria-label={`${candidate.interviewAudioFile.name}を削除`}>&times;</button>
                          </div>
                        </div>
                      ) : (
                        <p onClick={onAudioButtonClick}>ここに音声ファイルをドラッグ＆ドロップ、またはクリックして選択</p>
                      )}
                    </div>
                  </div>
                  <div className="form-group form-group-span-3">
                    <label>面談ログ（Google Meet議事録）</label>
                    <p className="form-helper-text">候補者名でGoogleドライブ内のMeet議事録を検索し、面談要約に追記できます。</p>
                    <button
                        type="button"
                        onClick={handleSearchInterviewLogs}
                        disabled={isSearchingInterviewLogs || isSummarizingInterviewLog}
                        className="secondary-action-button"
                    >
                        {isSearchingInterviewLogs ? '検索中...' : '候補者名でGoogleドライブを検索'}
                    </button>
                    {interviewLogResults && interviewLogResults.length > 0 && (
                        <ul className="user-management-list" style={{ marginTop: '0.5rem' }}>
                            {interviewLogResults.map(file => (
                                <li key={file.id} className="user-management-item">
                                    <span className="user-management-name">
                                        {file.name}
                                        <span style={{ color: '#888', fontSize: '0.85rem', marginLeft: '0.5rem' }}>
                                            {new Date(file.modifiedTime).toLocaleDateString('ja-JP')}
                                        </span>
                                    </span>
                                    <div className="user-management-actions">
                                        <button
                                            type="button"
                                            onClick={() => handleUseInterviewLog(file)}
                                            disabled={isSummarizingInterviewLog}
                                            className="save-user-button"
                                        >
                                            {isSummarizingInterviewLog ? '要約中...' : 'この面談ログを使う'}
                                        </button>
                                    </div>
                                </li>
                            ))}
                        </ul>
                    )}
                    <p className="form-helper-text" style={{ marginTop: '0.75rem' }}>
                        Driveから自動取得できない場合は、議事録のファイル（テキスト・PDF・Wordファイル）を直接アップロードしてください。
                    </p>
                    <div
                        id="interview-log-drop-zone"
                        className={`drop-zone ${interviewLogDragActive ? 'drag-active' : ''}`}
                        onDragEnter={handleInterviewLogDrag}
                        onDragLeave={handleInterviewLogDrag}
                        onDragOver={handleInterviewLogDrag}
                        onDrop={handleInterviewLogDrop}
                    >
                        <input
                            ref={interviewLogInputRef}
                            type="file"
                            id="interview-log-upload"
                            accept=".txt,text/plain,.pdf,application/pdf,.doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            onChange={handleInterviewLogFileChange}
                        />
                        <p onClick={onInterviewLogButtonClick}>
                            {isSummarizingInterviewLog ? '要約中...' : 'ここに議事録ファイル（.txt / .pdf / .doc・.docx）をドラッグ＆ドロップ、またはクリックして選択'}
                        </p>
                    </div>
                  </div>
              </div>

              <div className="form-grid">
                <div className="form-group">
                    <label htmlFor="name">候補者名 *</label>
                    <input type="text" id="name" name="name" value={candidate.name} onChange={handleChange} required />
                </div>
                <div className="form-group">
                    <label htmlFor="currentCompany">現職企業名</label>
                    <input type="text" id="currentCompany" name="currentCompany" value={candidate.currentCompany || ''} onChange={handleChange} placeholder="例: 株式会社ABC" />
                </div>
                 <div className="form-group">
                    <label htmlFor="education">最終学歴</label>
                    <input type="text" id="education" name="education" value={candidate.education || ''} onChange={handleChange} placeholder="例: 東京大学" />
                </div>
                <div className="form-group">
                    <label htmlFor="currentSalary">現職年収 (万円)</label>
                    <input type="number" id="currentSalary" name="currentSalary" value={candidate.currentSalary || ''} onChange={handleChange} placeholder="例: 500" />
                </div>
                <div className="form-group">
                    <label htmlFor="salary">希望年収 (万円)</label>
                    <input type="number" id="salary" name="salary" value={candidate.salary || ''} onChange={handleChange} placeholder="例: 650" />
                </div>
                <div className="form-group">
                    <label htmlFor="expectedAnnualSalary">想定年収 (万円)</label>
                    <input type="number" id="expectedAnnualSalary" name="expectedAnnualSalary" value={candidate.expectedAnnualSalary || ''} onChange={handleChange} placeholder="例: 600" />
                </div>
                <div className="form-group">
                    <label htmlFor="source">集客媒体</label>
                    <select id="source" name="source" value={candidate.source} onChange={handleChange}>
                        <option value="">選択してください</option>
                        {activeMedia.map(m => <option key={m.id} value={m.id}>{m.name}</option>)}
                        {candidate.source && candidate.source !== 'Other' && !activeMedia.some(m => m.id === candidate.source) && (
                            <option value={candidate.source}>
                                {(allMedia.find(m => m.id === candidate.source)?.name || candidate.source) + '（アーカイブ済み）'}
                            </option>
                        )}
                        <option value="Other">その他</option>
                    </select>
                </div>
                 <div className="form-group form-group-span-3">
                    <label>他エージェントの使用状況</label>
                    <div className="single-checkbox" style={{padding: 0}}>
                        <input 
                            type="checkbox"
                            id="usingOtherAgents"
                            name="usingOtherAgents"
                            checked={candidate.usingOtherAgents}
                            onChange={(e) => setCandidate(prev => ({ ...prev, usingOtherAgents: e.target.checked }))}
                        />
                        <label htmlFor="usingOtherAgents">他エージェントを利用している</label>
                    </div>
                </div>
                 <div className="form-group form-group-span-3">
                    <label htmlFor="summary">概要 (レジュメより)</label>
                    <div className="summary-wrapper">
                        <textarea 
                            id="summary" 
                            name="summary" 
                            value={candidate.summary} 
                            onChange={handleChange} 
                            rows={5}
                            placeholder={isGenerating ? "AIが履歴書を読み込んで要約を生成中です..." : "候補者の簡単な経歴や特徴"}
                            disabled={isGenerating}
                        />
                        {isGenerating && <div className="spinner"></div>}
                    </div>
                </div>
                <div className="form-group form-group-span-3">
                    <label htmlFor="interviewSummary">面談要約 (AI)</label>
                    <div className="summary-wrapper">
                        <textarea
                            id="interviewSummary"
                            name="interviewSummary"
                            value={candidate.interviewSummary}
                            onChange={handleChange}
                            rows={5}
                            placeholder={isGeneratingInterviewSummary ? "AIが音声データを解析し、要約を生成中です..." : "面談の音声ファイルをアップロードすると、ここに要約が自動生成されます。"}
                            disabled={isGeneratingInterviewSummary}
                        />
                        {isGeneratingInterviewSummary && <div className="spinner"></div>}
                    </div>
                 </div>
              </div>

              <div className="application-section">
                <h4 className="sub-section-title">選考状況</h4>
                {candidate.applications.length === 0 && <p className="no-data-message">まだ選考中の企業はありません。</p>}
                {candidate.applications.map((app, index) => (
                    <div key={app.id} className="application-card">
                       <button type="button" onClick={() => removeApplication(index)} className="remove-button" aria-label={`選考 ${index + 1} を削除`}>&times;</button>
                       <div className="form-group">
                          <label htmlFor={`companyName-${app.id}`}>企業名</label>
                          <input 
                            id={`companyName-${app.id}`}
                            type="text" 
                            placeholder="企業名" 
                            value={app.companyName} 
                            onChange={e => handleApplicationChange(index, 'companyName', e.target.value)}
                            aria-label={`企業名 ${index + 1}`}
                          />
                       </div>
                       <div className="form-group">
                         <label htmlFor={`stage-${app.id}`}>進捗状況</label>
                         <select 
                            id={`stage-${app.id}`}
                            value={app.stage} 
                            onChange={e => handleApplicationChange(index, 'stage', e.target.value)}
                            aria-label={`選考ステージ ${index + 1}`}
                          >
                              {PIPELINE_STAGES.map(stage => <option key={stage} value={stage}>{stage}</option>)}
                          </select>
                       </div>
                       <div className="form-group form-group-span-2">
                          <label htmlFor={`nextAction-${app.id}`}>次アクション</label>
                          <input
                            id={`nextAction-${app.id}`}
                            type="text"
                            placeholder="次アクション"
                            value={app.nextAction || ''}
                            onChange={e => handleApplicationChange(index, 'nextAction', e.target.value)}
                            aria-label={`次アクション ${index + 1}`}
                          />
                       </div>
                       <div className="form-group">
                          <label htmlFor={`scheduledDate-${app.id}`}>選考予定日</label>
                          <div className="scheduled-date-time-inputs">
                            <input
                              id={`scheduledDate-${app.id}`}
                              type="date"
                              value={app.scheduledDate || ''}
                              onChange={e => handleApplicationChange(index, 'scheduledDate', e.target.value)}
                              aria-label={`選考予定日 ${index + 1}`}
                            />
                            <input
                              id={`scheduledTime-${app.id}`}
                              type="time"
                              value={app.scheduledTime || ''}
                              onChange={e => handleApplicationChange(index, 'scheduledTime', e.target.value)}
                              aria-label={`開始時刻 ${index + 1}`}
                            />
                          </div>
                       </div>
                       <div className="form-group">
                          <label htmlFor={`expectedDecisionDate-${app.id}`}>意思決定時期</label>
                          <input
                            id={`expectedDecisionDate-${app.id}`}
                            type="date"
                            value={app.expectedDecisionDate || ''}
                            onChange={e => handleApplicationChange(index, 'expectedDecisionDate', e.target.value)}
                            aria-label={`意思決定時期 ${index + 1}`}
                          />
                       </div>
                       <div className="form-group">
                          <label htmlFor={`feeRate-${app.id}`}>fee料率(%)</label>
                          <input
                            id={`feeRate-${app.id}`}
                            type="number"
                            min="0"
                            step="0.1"
                            placeholder="例: 35"
                            value={app.feeRate ?? ''}
                            onChange={e => handleApplicationChange(index, 'feeRate', e.target.value)}
                            aria-label={`fee料率 ${index + 1}`}
                          />
                       </div>
                       <div className="form-group">
                          <label htmlFor={`offerConfidence-${app.id}`}>内定確度</label>
                          <select
                            id={`offerConfidence-${app.id}`}
                            value={app.offerConfidence || ''}
                            onChange={e => handleApplicationChange(index, 'offerConfidence', e.target.value)}
                            aria-label={`内定確度 ${index + 1}`}
                          >
                            <option value="">未設定</option>
                            {CONFIDENCE_GRADES.map(grade => <option key={grade} value={grade}>{grade}</option>)}
                          </select>
                       </div>
                       <div className="form-group">
                          <label htmlFor={`acceptanceConfidence-${app.id}`}>入社確度</label>
                          <select
                            id={`acceptanceConfidence-${app.id}`}
                            value={app.acceptanceConfidence || ''}
                            onChange={e => handleApplicationChange(index, 'acceptanceConfidence', e.target.value)}
                            aria-label={`入社確度 ${index + 1}`}
                          >
                            <option value="">未設定</option>
                            {CONFIDENCE_GRADES.map(grade => <option key={grade} value={grade}>{grade}</option>)}
                          </select>
                       </div>
                       <div className="form-group form-group-span-2">
                          <label htmlFor={`memo-${app.id}`}>メモ</label>
                          <textarea
                            id={`memo-${app.id}`}
                            placeholder="選考状況に関するメモ"
                            value={app.memo || ''}
                            onChange={e => handleApplicationChange(index, 'memo', e.target.value)}
                            aria-label={`メモ ${index + 1}`}
                            rows={2}
                          />
                       </div>
                    </div>
                ))}
                <button type="button" onClick={addApplication} className="add-button">+ 選考企業を追加</button>
              </div>

              { (resumeDragActive || audioDragActive || interviewLogDragActive) && (
                <div
                  id="drag-file-element"
                  onDragEnter={audioDragActive ? handleAudioDrag : interviewLogDragActive ? handleInterviewLogDrag : handleResumeDrag}
                  onDragLeave={audioDragActive ? handleAudioDrag : interviewLogDragActive ? handleInterviewLogDrag : handleResumeDrag}
                  onDragOver={audioDragActive ? handleAudioDrag : interviewLogDragActive ? handleInterviewLogDrag : handleResumeDrag}
                  onDrop={audioDragActive ? handleAudioDrop : interviewLogDragActive ? handleInterviewLogDrop : handleResumeDrop}
                ></div>
              )}
            </form>
            <div className="modal-footer">
                <button type="button" onClick={onClose} className="cancel-button">キャンセル</button>
                <button type="submit" form="candidate-form" className="submit-button">保存</button>
            </div>
        </div>
      </div>
    );
};


const ApplicationModal: React.FC<{
  isOpen: boolean;
  onClose: () => void;
  onSave: (application: CompanyApplication) => void;
  candidateName: string;
  initialData: CompanyApplication | null; // null for new, object for edit
}> = ({ isOpen, onClose, onSave, candidateName, initialData }) => {
    const [application, setApplication] = useState<CompanyApplication>({
        id: '', companyName: '', stage: '打診', nextAction: '', scheduledDate: '', memo: ''
    });

    useEffect(() => {
        if (isOpen) {
            if (initialData) {
                setApplication(initialData); // Editing existing application
            } else {
                // Adding new application
                setApplication({
                    id: `app_${Date.now()}`,
                    companyName: '',
                    stage: '打診',
                    nextAction: '',
                    scheduledDate: '',
                    memo: ''
                });
            }
        }
    }, [isOpen, initialData]);

    if (!isOpen) return null;

    const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
        const { name, value } = e.target;
        if (name === 'feeRate') {
            setApplication(prev => ({ ...prev, feeRate: value === '' ? undefined : Number(value) }));
            return;
        }
        setApplication(prev => ({ ...prev, [name]: value }));
    };

    const handleSubmit = (e: React.FormEvent) => {
        e.preventDefault();
        if (!application.companyName.trim()) {
            alert('企業名は必須です。');
            return;
        }
        onSave(application);
        onClose();
    };

    const modalTitle = `${candidateName}さん - ${initialData && initialData.companyName ? '選考情報を編集' : '選考情報を追加'}`;

    return (
        <div className="modal-overlay" onClick={onClose}>
            <div className="modal-content application-modal" onClick={e => e.stopPropagation()}>
                <div className="modal-header">
                    <h3 id="application-modal-title">{modalTitle}</h3>
                    <button onClick={onClose} className="close-button" aria-label="閉じる">&times;</button>
                </div>
                <form id="application-form" className="modal-body" onSubmit={handleSubmit}>
                    <div className="form-group">
                        <label htmlFor="companyName">企業名 *</label>
                        <input
                            type="text"
                            id="companyName"
                            name="companyName"
                            value={application.companyName}
                            onChange={handleChange}
                            required
                            autoFocus
                        />
                    </div>
                    <div className="form-group">
                        <label htmlFor="stage">進捗状況</label>
                        <select
                            id="stage"
                            name="stage"
                            value={application.stage}
                            onChange={handleChange}
                        >
                            {PIPELINE_STAGES.map(stage => <option key={stage} value={stage}>{stage}</option>)}
                        </select>
                    </div>
                    <div className="form-group">
                        <label htmlFor="nextAction">次アクション</label>
                        <input
                            type="text"
                            id="nextAction"
                            name="nextAction"
                            value={application.nextAction}
                            onChange={handleChange}
                        />
                    </div>
                    <div className="form-group">
                        <label htmlFor="scheduledDate">選考予定日</label>
                        <div className="scheduled-date-time-inputs">
                            <input
                                type="date"
                                id="scheduledDate"
                                name="scheduledDate"
                                value={application.scheduledDate || ''}
                                onChange={handleChange}
                            />
                            <input
                                type="time"
                                id="scheduledTime"
                                name="scheduledTime"
                                aria-label="開始時刻"
                                value={application.scheduledTime || ''}
                                onChange={handleChange}
                            />
                        </div>
                    </div>
                    <div className="form-group">
                        <label htmlFor="expectedDecisionDate">意思決定時期</label>
                        <input
                            type="date"
                            id="expectedDecisionDate"
                            name="expectedDecisionDate"
                            value={application.expectedDecisionDate || ''}
                            onChange={handleChange}
                        />
                    </div>
                    <div className="form-group">
                        <label htmlFor="feeRate">fee料率(%)</label>
                        <input
                            type="number"
                            id="feeRate"
                            name="feeRate"
                            min="0"
                            step="0.1"
                            placeholder="例: 35"
                            value={application.feeRate ?? ''}
                            onChange={handleChange}
                        />
                    </div>
                    <div className="form-group">
                        <label htmlFor="offerConfidence">内定確度</label>
                        <select
                            id="offerConfidence"
                            name="offerConfidence"
                            value={application.offerConfidence || ''}
                            onChange={handleChange}
                        >
                            <option value="">未設定</option>
                            {CONFIDENCE_GRADES.map(grade => <option key={grade} value={grade}>{grade}</option>)}
                        </select>
                    </div>
                    <div className="form-group">
                        <label htmlFor="acceptanceConfidence">入社確度</label>
                        <select
                            id="acceptanceConfidence"
                            name="acceptanceConfidence"
                            value={application.acceptanceConfidence || ''}
                            onChange={handleChange}
                        >
                            <option value="">未設定</option>
                            {CONFIDENCE_GRADES.map(grade => <option key={grade} value={grade}>{grade}</option>)}
                        </select>
                    </div>
                    <div className="form-group">
                        <label htmlFor="memo">メモ</label>
                        <textarea
                            id="memo"
                            name="memo"
                            value={application.memo || ''}
                            onChange={handleChange}
                            rows={3}
                            placeholder="選考状況に関するメモ"
                        />
                    </div>
                </form>
                <div className="modal-footer">
                    <button type="button" onClick={onClose} className="cancel-button">キャンセル</button>
                    <button type="submit" form="application-form" className="submit-button">保存</button>
                </div>
            </div>
        </div>
    );
};


/**
 * Registers or edits a candidate's 掘り起しリスト entry — a candidate-level (not per-
 * application) note that this person isn't being actively pursued right now but should be
 * revisited later. Saving here is what actually hides the candidate from the active pipeline
 * (see CandidatePipelineView's handleSaveRevival) and schedules the reminder on the calendar.
 */
const RevivalModal: React.FC<{
    isOpen: boolean;
    onClose: () => void;
    onSave: (data: { nextAction: string; nextActionDate: string }) => void;
    candidateName: string;
    initialData: { nextAction: string; nextActionDate: string } | null;
}> = ({ isOpen, onClose, onSave, candidateName, initialData }) => {
    const [nextAction, setNextAction] = useState('');
    const [nextActionDate, setNextActionDate] = useState('');

    useEffect(() => {
        if (isOpen) {
            setNextAction(initialData?.nextAction || '');
            setNextActionDate(initialData?.nextActionDate || '');
        }
    }, [isOpen, initialData]);

    if (!isOpen) return null;

    const handleSubmit = (e: React.FormEvent) => {
        e.preventDefault();
        if (!nextActionDate) {
            alert('次アクションの時期を入力してください。');
            return;
        }
        onSave({ nextAction: nextAction.trim(), nextActionDate });
    };

    return (
        <div className="modal-overlay" onClick={onClose}>
            <div className="modal-content" onClick={e => e.stopPropagation()}>
                <div className="modal-header">
                    <h3 id="revival-modal-title">{candidateName}さん - 掘り起しリストに登録</h3>
                    <button onClick={onClose} className="close-button" aria-label="閉じる">&times;</button>
                </div>
                <form id="revival-form" className="modal-body" onSubmit={handleSubmit}>
                    <p className="modal-description">
                        パイプライン上は非表示になり、掘り起しリストとして確認できるようになります。指定した時期にパイプラインカレンダー上にリマインドが表示されます。
                    </p>
                    <div className="form-group">
                        <label htmlFor="revival-next-action">次アクション</label>
                        <input
                            id="revival-next-action"
                            type="text"
                            placeholder="例: 半年後に改めて連絡"
                            value={nextAction}
                            onChange={(e) => setNextAction(e.target.value)}
                        />
                    </div>
                    <div className="form-group">
                        <label htmlFor="revival-next-action-date">次アクションの時期</label>
                        <input
                            id="revival-next-action-date"
                            type="date"
                            value={nextActionDate}
                            onChange={(e) => setNextActionDate(e.target.value)}
                            required
                        />
                    </div>
                </form>
                <div className="modal-footer">
                    <button type="button" onClick={onClose} className="cancel-button">キャンセル</button>
                    <button type="submit" form="revival-form" className="submit-button">保存</button>
                </div>
            </div>
        </div>
    );
};


/**
 * Opened by clicking a date on the pipeline calendar. Lets the signed-in user pick one of
 * their own candidates, then either an existing application (to reschedule/edit) or a brand
 * new one, for that date — the actual company/stage/memo fields are edited via the existing
 * ApplicationModal once a choice is made here.
 */
const PipelineDateScheduleModal: React.FC<{
    date: string;
    ownCandidates: Candidate[];
    onClose: () => void;
    onPickApplication: (candidate: Candidate, application: CompanyApplication | null) => void;
}> = ({ date, ownCandidates, onClose, onPickApplication }) => {
    const [selectedCandidateId, setSelectedCandidateId] = useState('');
    const selectedCandidate = ownCandidates.find(c => c.id === selectedCandidateId) || null;
    const dateLabel = new Date(date + 'T00:00:00').toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric', weekday: 'short' });

    return (
        <div className="modal-overlay" onClick={onClose}>
            <div className="modal-content" onClick={e => e.stopPropagation()}>
                <div className="modal-header">
                    <h3 id="schedule-modal-title">{dateLabel}の選考予定を入力</h3>
                    <button onClick={onClose} className="close-button" aria-label="閉じる">&times;</button>
                </div>
                <div className="modal-body">
                    {ownCandidates.length === 0 ? (
                        <p className="no-data-message">自分のパイプラインに候補者がいません。まず候補者を登録してください。</p>
                    ) : (
                        <>
                            <div className="form-group">
                                <label htmlFor="schedule-candidate-select">候補者を選択</label>
                                <select
                                    id="schedule-candidate-select"
                                    value={selectedCandidateId}
                                    onChange={e => setSelectedCandidateId(e.target.value)}
                                    autoFocus
                                >
                                    <option value="">選択してください</option>
                                    {ownCandidates.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                                </select>
                            </div>
                            {selectedCandidate && (
                                <div className="form-group">
                                    <label>選考企業を選択（既存を選ぶか、新しく追加）</label>
                                    <div className="schedule-modal-application-list">
                                        {selectedCandidate.applications.filter(app => !app.isHidden).map(app => (
                                            <button
                                                key={app.id}
                                                type="button"
                                                className="secondary-action-button"
                                                onClick={() => onPickApplication(selectedCandidate, app)}
                                            >
                                                {app.companyName} ({app.stage})
                                            </button>
                                        ))}
                                        <button
                                            type="button"
                                            className="add-button"
                                            onClick={() => onPickApplication(selectedCandidate, null)}
                                        >
                                            + 新しい選考を追加
                                        </button>
                                    </div>
                                </div>
                            )}
                        </>
                    )}
                </div>
                <div className="modal-footer">
                    <button type="button" onClick={onClose} className="cancel-button">キャンセル</button>
                </div>
            </div>
        </div>
    );
};


const PipelineDashboard: React.FC<{ candidates: Candidate[] }> = ({ candidates }) => {
    const STAGE_WEIGHTS: Record<PipelineStage, number> = {
        '打診': 1,
        '書類選考': 2,
        '適性検査': 3,
        'カジュアル面談': 4,
        '1次面接': 5,
        '2次面接': 6,
        '最終面接': 7,
        '内定': 8,
        '内定承諾': 9,
        'お見送り': 0,
        '選考辞退': 0,
    };

    const stageCounts = useMemo(() => {
        const counts = PIPELINE_STAGES.reduce((acc, stage) => {
            acc[stage] = 0;
            return acc;
        }, {} as Record<PipelineStage, number>);

        const visibleCandidates = candidates.filter(c => !c.isHidden);

        visibleCandidates.forEach(candidate => {
            const visibleApplications = candidate.applications.filter(app => !app.isHidden);
            if (visibleApplications.length === 0) {
                return;
            }

            let mostAdvancedStage: PipelineStage | null = null;
            let maxWeight = -1;

            visibleApplications.forEach(app => {
                const weight = STAGE_WEIGHTS[app.stage];
                if (weight > maxWeight) {
                    maxWeight = weight;
                    mostAdvancedStage = app.stage;
                }
            });

            if (mostAdvancedStage) {
                counts[mostAdvancedStage]++;
            }
        });

        return counts;
    }, [candidates]);

    return (
        <div className="pipeline-dashboard">
            {PIPELINE_STAGES.map(stage => (
                <div key={stage} className="pipeline-stage-card" style={{'--badge-color': STAGE_COLOR_MAP[stage]} as React.CSSProperties}>
                    <span className="stage-name">{stage}</span>
                    <span className="stage-count">{stageCounts[stage]}</span>
                </div>
            ))}
        </div>
    );
};


interface StageGrossProfit {
    stage: PipelineStage;
    count: number;
    estimableCount: number;
    revenue: number;
    cost: number;
    profit: number;
    entries: CompanyPipelineEntry[];
}

/**
 * Expected gross profit for one application: revenue is the client referral fee (candidate's
 * expected annual salary × the position's own fee rate — different positions for the same
 * candidate can negotiate different rates). The media's fee rate is a cut OF THAT REVENUE
 * (not of the candidate's salary directly) — i.e. 粗利 = 紹介料 − 紹介料×媒体手数料率.
 * Returns null when either the candidate's expected annual salary or the position's fee rate
 * hasn't been entered yet, so callers can tell "zero profit" apart from "not enough data".
 */
function computeApplicationGrossProfit(
    candidate: Candidate,
    application: CompanyApplication,
    mediaFeeRateById: Map<string, number>
): { revenue: number; cost: number; profit: number } | null {
    if (!candidate.expectedAnnualSalary || application.feeRate === undefined || application.feeRate === null) return null;
    const revenue = candidate.expectedAnnualSalary * (application.feeRate / 100);
    const mediaFeeRate = mediaFeeRateById.get(candidate.source) || 0;
    const cost = revenue * (mediaFeeRate / 100);
    return { revenue, cost, profit: revenue - cost };
}

/** Lower is better; a missing rating ranks worse than any explicit grade (CONFIDENCE_RANK.length, i.e. past 'C'). */
function confidenceScore(application: CompanyApplication): number {
    const offerRank = application.offerConfidence ? CONFIDENCE_RANK[application.offerConfidence] : CONFIDENCE_GRADES.length;
    const acceptanceRank = application.acceptanceConfidence ? CONFIDENCE_RANK[application.acceptanceConfidence] : CONFIDENCE_GRADES.length;
    return offerRank + acceptanceRank;
}

// 'nearestExpectedDecisionDate' isn't a real Candidate field — it's derived on the fly from
// applications (a candidate can have several, each with their own 意思決定時期) rather than
// stored directly, unlike the other sortable fields which are plain Candidate properties.
type PipelineSortKey = keyof Candidate | 'nearestExpectedDecisionDate';

/** Soonest 意思決定時期 among a candidate's visible applications, or null if none is set. */
function getNearestExpectedDecisionDate(candidate: Candidate): string | null {
    const dates = candidate.applications
        .filter(app => !app.isHidden && app.expectedDecisionDate)
        .map(app => app.expectedDecisionDate as string);
    if (dates.length === 0) return null;
    return dates.reduce((min, d) => (d < min ? d : min));
}

/** True if dateISO falls in the calendar month `monthOffset` months from today (0 = this month, 1 = next month). */
function isDateInRelativeMonth(dateISO: string, monthOffset: number): boolean {
    const d = new Date(dateISO + 'T00:00:00');
    const now = new Date();
    const target = new Date(now.getFullYear(), now.getMonth() + monthOffset, 1);
    return d.getFullYear() === target.getFullYear() && d.getMonth() === target.getMonth();
}

/**
 * A candidate interviewing at several companies at once can only actually be placed at one of
 * them — summing gross profit across all their applications would double-count the same
 * placement. Picks the single most-likely-to-close application (lowest combined offer +
 * acceptance confidence rank) per candidate instead.
 */
function pickBestApplicationPerCandidate(candidates: Candidate[]): CompanyPipelineEntry[] {
    const result: CompanyPipelineEntry[] = [];
    candidates.filter(c => !c.isHidden).forEach(candidate => {
        const visibleApps = candidate.applications.filter(app => !app.isHidden);
        if (visibleApps.length === 0) return;
        const best = visibleApps.reduce((a, b) => (confidenceScore(b) < confidenceScore(a) ? b : a));
        result.push({ candidate, application: best });
    });
    return result;
}

function computeGrossProfitByStage(candidates: Candidate[], allMedia: MediaEntry[]): StageGrossProfit[] {
    const mediaFeeRateById = new Map(allMedia.map(m => [m.id, m.feeRate || 0]));
    const totalsByStage = new Map<PipelineStage, StageGrossProfit>(
        PIPELINE_STAGES.map(stage => [stage, { stage, count: 0, estimableCount: 0, revenue: 0, cost: 0, profit: 0, entries: [] }])
    );

    pickBestApplicationPerCandidate(candidates).forEach(entry => {
        const { candidate, application } = entry;
        const bucket = totalsByStage.get(application.stage)!;
        bucket.count++;
        bucket.entries.push(entry);
        const result = computeApplicationGrossProfit(candidate, application, mediaFeeRateById);
        if (result) {
            bucket.estimableCount++;
            bucket.revenue += result.revenue;
            bucket.cost += result.cost;
            bucket.profit += result.profit;
        }
    });

    return PIPELINE_STAGES.map(stage => totalsByStage.get(stage)!);
}

const formatManYen = (n: number): string => `${Math.round(n).toLocaleString()}万円`;

/**
 * Shows expected gross profit (client referral fee minus the sourcing media's handling fee)
 * broken down by selection stage, plus a grand total across every stage except お見送り/
 * 選考辞退 (lost — they'll never generate revenue). Applications missing either the
 * candidate's expected annual salary or their own fee rate are counted but excluded from the
 * sums, and that gap is surfaced rather than silently under-reporting the total.
 */
const GrossProfitSummary: React.FC<{ candidates: Candidate[]; allMedia: MediaEntry[] }> = ({ candidates, allMedia }) => {
    const [selectedStageFilters, setSelectedStageFilters] = useState<PipelineStage[]>([]);
    const toggleStageFilter = (stage: PipelineStage) => {
        setSelectedStageFilters(prev => prev.includes(stage) ? prev.filter(s => s !== stage) : [...prev, stage]);
    };

    const stageTotals = useMemo(() => computeGrossProfitByStage(candidates, allMedia), [candidates, allMedia]);

    // With no explicit filter, cards show every stage that has data but the total still
    // excludes lost stages (they'll never generate revenue) — the same default as before this
    // filter existed. Once the user picks specific stages, both the cards and the total narrow
    // to exactly that selection, even if it includes お見送り/選考辞退 — that's an explicit choice.
    const visibleStageTotals = useMemo(() => {
        if (selectedStageFilters.length === 0) return stageTotals;
        return stageTotals.filter(s => selectedStageFilters.includes(s.stage));
    }, [stageTotals, selectedStageFilters]);

    const grandTotal = useMemo(() => {
        const base = selectedStageFilters.length > 0
            ? visibleStageTotals
            : stageTotals.filter(s => s.stage !== 'お見送り' && s.stage !== '選考辞退');
        return base.reduce((acc, s) => ({
            count: acc.count + s.count,
            estimableCount: acc.estimableCount + s.estimableCount,
            revenue: acc.revenue + s.revenue,
            cost: acc.cost + s.cost,
            profit: acc.profit + s.profit,
        }), { count: 0, estimableCount: 0, revenue: 0, cost: 0, profit: 0 });
    }, [stageTotals, visibleStageTotals, selectedStageFilters]);

    return (
        <div className="gross-profit-summary">
            <div className="pipeline-sort-controls">
                <span>選考フェーズで絞り込み:</span>
                {PIPELINE_STAGES.map(stage => (
                    <button
                        key={stage}
                        type="button"
                        onClick={() => toggleStageFilter(stage)}
                        className={selectedStageFilters.includes(stage) ? 'active' : ''}
                    >
                        {stage}
                    </button>
                ))}
                {selectedStageFilters.length > 0 && (
                    <button type="button" onClick={() => setSelectedStageFilters([])} className="secondary-action-button">
                        クリア
                    </button>
                )}
            </div>
            <div className="gross-profit-total-card">
                <div className="gross-profit-total-item">
                    <span>想定紹介料（合計）</span>
                    <strong>{formatManYen(grandTotal.revenue)}</strong>
                </div>
                <div className="gross-profit-total-item">
                    <span>想定媒体手数料（合計）</span>
                    <strong>{formatManYen(grandTotal.cost)}</strong>
                </div>
                <div className="gross-profit-total-item highlight">
                    <span>想定粗利（合計）</span>
                    <strong>{formatManYen(grandTotal.profit)}</strong>
                </div>
            </div>
            {grandTotal.count > grandTotal.estimableCount && (
                <p className="gross-profit-note">
                    ※ {selectedStageFilters.length > 0 ? '選択中のフェーズ' : 'お見送り・選考辞退を除く'}{grandTotal.count}件中、想定年収とfee料率が両方入力済みの{grandTotal.estimableCount}件のみを集計しています（残り{grandTotal.count - grandTotal.estimableCount}件は未入力のため対象外）。
                </p>
            )}
            <div className="detail-application-grid">
                {visibleStageTotals.filter(s => s.count > 0).map(s => (
                    <div key={s.stage} className="detail-application-card" style={{ borderLeftColor: STAGE_COLOR_MAP[s.stage] }}>
                        <div className="detail-card-header">
                            <span className="status-badge" style={{ '--badge-color': STAGE_COLOR_MAP[s.stage] } as React.CSSProperties}>{s.stage}</span>
                            <span className="company-pipeline-count">{s.count}件</span>
                        </div>
                        <div className="detail-card-body">
                            <div className="detail-card-item"><span>想定紹介料:</span><span>{formatManYen(s.revenue)}</span></div>
                            <div className="detail-card-item"><span>想定媒体手数料:</span><span>{formatManYen(s.cost)}</span></div>
                            <div className="detail-card-item"><span>想定粗利:</span><span>{formatManYen(s.profit)}</span></div>
                        </div>
                        {s.entries.length > 0 && (
                            <div className="company-pipeline-entries" style={{ marginTop: '0.75rem' }}>
                                {s.entries.map(({ candidate, application }) => (
                                    <div key={application.id} className="company-pipeline-entry">
                                        <span className="company-pipeline-entry-name">
                                            {candidate.name} - {application.companyName}
                                            {candidate.ownerLabel && <small> ({candidate.ownerLabel})</small>}
                                        </span>
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                ))}
            </div>
        </div>
    );
};


type PipelineCalendarEvent =
    | { kind: 'application'; candidate: Candidate; application: CompanyApplication }
    | { kind: 'revival'; candidate: Candidate };

/**
 * Month calendar of every visible application's manually-set scheduledDate, plus a 掘り起し
 * reminder event on each revival-listed candidate's nextActionDate. Reads whatever `candidates`
 * list the caller passes in, so it's automatically scoped by the pipeline's existing 自分/全
 * ユーザー/チーム/ユーザー別 switcher — no separate scope control needed here. Clicking empty
 * day space lets the signed-in user schedule (or edit) one of their own candidates'
 * applications for that date, via onDayClick. Clicking an existing event bar instead jumps
 * straight to editing THAT application/revival entry — but only when it belongs to the
 * signed-in user (stops propagation so it doesn't also trigger onDayClick); other people's
 * events are shown for visibility only and aren't clickable.
 */
const PipelineCalendarView: React.FC<{
    candidates: Candidate[];
    viewDate: Date;
    currentUserEmail: string;
    onPrevMonth: () => void;
    onNextMonth: () => void;
    onDayClick: (dateStr: string) => void;
    onEditApplication: (candidate: Candidate, application: CompanyApplication) => void;
    onEditRevival: (candidate: Candidate) => void;
}> = ({ candidates, viewDate, currentUserEmail, onPrevMonth, onNextMonth, onDayClick, onEditApplication, onEditRevival }) => {
    const year = viewDate.getFullYear();
    const month = viewDate.getMonth();

    const eventsByDate = useMemo(() => {
        const map = new Map<string, PipelineCalendarEvent[]>();
        // Application events only come from visible (non-hidden) candidates — once a candidate
        // is hidden (including via 掘り起しリスト, below), its past applications' scheduled
        // events stop cluttering the calendar.
        candidates.filter(c => !c.isHidden).forEach(c => {
            c.applications.filter(app => !app.isHidden && app.scheduledDate).forEach(app => {
                const list = map.get(app.scheduledDate!) || [];
                list.push({ kind: 'application', candidate: c, application: app });
                map.set(app.scheduledDate!, list);
            });
        });
        // Revival reminders come from EVERY candidate regardless of isHidden — a 掘り起しリスト
        // entry is always hidden by definition, but its reminder must still show.
        candidates.forEach(c => {
            if (c.revival?.nextActionDate) {
                const list = map.get(c.revival.nextActionDate) || [];
                list.push({ kind: 'revival', candidate: c });
                map.set(c.revival.nextActionDate, list);
            }
        });
        // Events without a start time (including every revival reminder) sort last, after every
        // timed application event on the same day.
        map.forEach(list => list.sort((a, b) => {
            const aTime = a.kind === 'application' ? (a.application.scheduledTime || '99:99') : '99:99';
            const bTime = b.kind === 'application' ? (b.application.scheduledTime || '99:99') : '99:99';
            return aTime.localeCompare(bTime);
        }));
        return map;
    }, [candidates]);

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayLocalString = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;

    const firstDayOfMonth = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();

    const days = [];
    for (let i = 0; i < firstDayOfMonth; i++) {
        days.push(<div key={`empty-${i}`} className="calendar-day empty"></div>);
    }

    for (let i = 1; i <= daysInMonth; i++) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(i).padStart(2, '0')}`;
        const events = eventsByDate.get(dateStr) || [];
        const isToday = dateStr === todayLocalString;

        days.push(
            <div
                key={i}
                className={`calendar-day pipeline-calendar-day ${isToday ? 'today' : ''} ${events.length > 0 ? 'has-data' : ''}`}
                onClick={() => onDayClick(dateStr)}
                role="button"
                tabIndex={0}
                aria-label={`${i}日, 選考予定を入力`}
                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onDayClick(dateStr); } }}
            >
                <div className="day-number">{i}</div>
                {events.length > 0 && (
                    <div className="pipeline-calendar-events">
                        {events.map((ev, idx) => {
                            const isOwnEvent = !ev.candidate.ownerEmail || ev.candidate.ownerEmail === currentUserEmail;
                            if (ev.kind === 'revival') {
                                const handleRevivalActivate = (e: React.SyntheticEvent) => {
                                    e.stopPropagation();
                                    onEditRevival(ev.candidate);
                                };
                                return (
                                    <div
                                        key={idx}
                                        className={`pipeline-calendar-event pipeline-calendar-event-revival ${isOwnEvent ? 'is-editable' : ''}`}
                                        title={`掘り起し: ${ev.candidate.name}${ev.candidate.revival?.nextAction ? ` / ${ev.candidate.revival.nextAction}` : ''}${ev.candidate.ownerLabel ? ` (${ev.candidate.ownerLabel})` : ''}${isOwnEvent ? ' — クリックして編集' : ''}`}
                                        role={isOwnEvent ? 'button' : undefined}
                                        tabIndex={isOwnEvent ? 0 : undefined}
                                        onClick={isOwnEvent ? handleRevivalActivate : undefined}
                                        onKeyDown={isOwnEvent ? (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleRevivalActivate(e); } } : undefined}
                                    >
                                        <span className="pipeline-calendar-event-stage">掘り起し</span>
                                        {ev.candidate.name}
                                    </div>
                                );
                            }
                            const handleEventActivate = (e: React.SyntheticEvent) => {
                                e.stopPropagation();
                                onEditApplication(ev.candidate, ev.application);
                            };
                            return (
                                <div
                                    key={idx}
                                    className={`pipeline-calendar-event ${isOwnEvent ? 'is-editable' : ''}`}
                                    style={{ '--badge-color': STAGE_COLOR_MAP[ev.application.stage] } as React.CSSProperties}
                                    title={`${ev.application.scheduledTime ? `${ev.application.scheduledTime} ` : ''}${ev.candidate.name} / ${ev.application.companyName} / ${ev.application.stage}${ev.candidate.ownerLabel ? ` (${ev.candidate.ownerLabel})` : ''}${isOwnEvent ? ' — クリックして編集' : ''}`}
                                    role={isOwnEvent ? 'button' : undefined}
                                    tabIndex={isOwnEvent ? 0 : undefined}
                                    onClick={isOwnEvent ? handleEventActivate : undefined}
                                    onKeyDown={isOwnEvent ? (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleEventActivate(e); } } : undefined}
                                >
                                    <span className="pipeline-calendar-event-stage">{STAGE_SHORT_LABELS[ev.application.stage]}</span>
                                    {ev.application.scheduledTime && <span className="pipeline-calendar-event-time">{ev.application.scheduledTime}</span>}
                                    {ev.candidate.name} - {ev.application.companyName}
                                </div>
                            );
                        })}
                    </div>
                )}
            </div>
        );
    }

    return (
        <div className="calendar-container pipeline-calendar-container">
            <div className="calendar-header">
                <button onClick={onPrevMonth} aria-label="前の月へ">&lt; 前月</button>
                <h2>{`${year}年 ${month + 1}月`}</h2>
                <button onClick={onNextMonth} aria-label="次の月へ">次月 &gt;</button>
            </div>
            <div className="calendar-grid-header">
                {['日', '月', '火', '水', '木', '金', '土'].map(day => <div key={day}>{day}</div>)}
            </div>
            <div className="calendar-grid pipeline-calendar-grid">
                {days}
            </div>
        </div>
    );
};


interface CompanyPipelineEntry {
    candidate: Candidate;
    application: CompanyApplication;
}

interface CompanyPipelineGroup {
    key: string;
    displayName: string;
    variants: string[]; // every distinct raw spelling seen, for transparency when names were merged
    entries: CompanyPipelineEntry[];
}

/** Groups every visible application by company, using normalizeCompanyName to merge spelling variants. */
function groupApplicationsByCompany(candidates: Candidate[]): CompanyPipelineGroup[] {
    const groups = new Map<string, { nameCounts: Map<string, number>; entries: CompanyPipelineEntry[] }>();

    candidates.filter(c => !c.isHidden).forEach(candidate => {
        candidate.applications.filter(app => !app.isHidden && app.companyName.trim()).forEach(application => {
            const rawName = application.companyName.trim();
            const key = normalizeCompanyName(rawName) || rawName;
            if (!groups.has(key)) groups.set(key, { nameCounts: new Map(), entries: [] });
            const group = groups.get(key)!;
            group.nameCounts.set(rawName, (group.nameCounts.get(rawName) || 0) + 1);
            group.entries.push({ candidate, application });
        });
    });

    return Array.from(groups.entries()).map(([key, group]) => {
        const variants = Array.from(group.nameCounts.keys());
        const displayName = variants.reduce((best, v) =>
            (group.nameCounts.get(v)! > group.nameCounts.get(best)!) ? v : best, variants[0]);
        return { key, displayName, variants, entries: group.entries };
    }).sort((a, b) => b.entries.length - a.entries.length);
}

/**
 * Shows every visible application grouped by company instead of by candidate. Company names
 * are merged via normalizeCompanyName so minor spelling/formatting differences (株式会社
 * prefix vs. suffix, full-width characters, extra spaces, etc.) don't split one company into
 * several groups — anyone typing the name slightly differently still lands in the same group.
 */
const CompanyPipelineView: React.FC<{
    candidates: Candidate[];
    currentUserEmail: string;
    onEditApplication: (candidate: Candidate, application: CompanyApplication) => void;
}> = ({ candidates, currentUserEmail, onEditApplication }) => {
    const [searchTerm, setSearchTerm] = useState('');

    const groups = useMemo(() => groupApplicationsByCompany(candidates), [candidates]);

    const filteredGroups = useMemo(() => {
        if (!searchTerm.trim()) return groups;
        const normalizedSearch = normalizeCompanyName(searchTerm);
        return groups.filter(g =>
            g.key.includes(normalizedSearch) || g.variants.some(v => v.includes(searchTerm))
        );
    }, [groups, searchTerm]);

    return (
        <div className="company-pipeline-view">
            <input
                type="text"
                placeholder="企業名で検索..."
                value={searchTerm}
                onChange={e => setSearchTerm(e.target.value)}
                className="search-input"
                style={{ marginBottom: '1rem' }}
            />
            {filteredGroups.length === 0 && <p className="no-data-message">該当する企業がありません。</p>}
            <div className="detail-application-grid">
                {filteredGroups.map(group => (
                    <div key={group.key} className="detail-application-card company-pipeline-card">
                        <div className="detail-card-header">
                            <strong>{group.displayName}</strong>
                            <span className="company-pipeline-count">{group.entries.length}件</span>
                        </div>
                        {group.variants.length > 1 && (
                            <p className="company-pipeline-variants">表記ゆれとして統合: {group.variants.join(' / ')}</p>
                        )}
                        <div className="company-pipeline-entries">
                            {group.entries.map(({ candidate, application }) => {
                                const candidateIsOwn = !candidate.ownerEmail || candidate.ownerEmail === currentUserEmail;
                                return (
                                    <div key={application.id} className="company-pipeline-entry">
                                        <span className="company-pipeline-entry-name">
                                            {candidate.name}
                                            {candidate.ownerLabel && <small> ({candidate.ownerLabel})</small>}
                                        </span>
                                        <span
                                            className={`status-badge ${candidateIsOwn ? 'status-badge-clickable' : ''}`}
                                            style={{ '--badge-color': STAGE_COLOR_MAP[application.stage] } as React.CSSProperties}
                                            onClick={candidateIsOwn ? () => onEditApplication(candidate, application) : undefined}
                                            role={candidateIsOwn ? 'button' : undefined}
                                            tabIndex={candidateIsOwn ? 0 : undefined}
                                            title={candidateIsOwn ? 'クリックして選考状況を更新' : undefined}
                                            onKeyDown={candidateIsOwn ? (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onEditApplication(candidate, application); } } : undefined}
                                        >
                                            {application.stage}
                                        </span>
                                        {application.scheduledDate && (
                                            <span className="company-pipeline-entry-date">
                                                {new Date(application.scheduledDate + 'T00:00:00').toLocaleDateString('ja-JP')}
                                                {application.scheduledTime ? ` ${application.scheduledTime}` : ''}
                                            </span>
                                        )}
                                    </div>
                                );
                            })}
                        </div>
                    </div>
                ))}
            </div>
        </div>
    );
};


const SourceEffectivenessReport: React.FC<{ candidates: Candidate[]; allMedia: MediaEntry[] }> = ({ candidates, allMedia }) => {
    const reportData = useMemo(() => {
        const visibleCandidates = candidates.filter(c => !c.isHidden);
        // Group by media name (active or archived — this is historical data, so archived
        // media must still be included). Falls back to the raw stored value for candidates
        // whose source predates the current media list.
        const mediaNameById = new Map<string, string>(allMedia.map(m => [m.id, m.name] as [string, string]));
        const resolveSourceLabel = (source: string) => source ? (mediaNameById.get(source) || source) : '未設定';

        const grouped = new Map<string, Candidate[]>();
        visibleCandidates.forEach(c => {
            const label = resolveSourceLabel(c.source);
            if (!grouped.has(label)) grouped.set(label, []);
            grouped.get(label)!.push(c);
        });

        const stats = Array.from(grouped.entries()).map(([sourceName, sourceCandidates]) => {
            const total = sourceCandidates.length;

            const reachedInterview = sourceCandidates.filter(c =>
                c.applications.some(app => !app.isHidden && PIPELINE_STAGES.indexOf(app.stage) >= PIPELINE_STAGES.indexOf('1次面接'))
            ).length;

            const receivedOffer = sourceCandidates.filter(c =>
                c.applications.some(app => !app.isHidden && PIPELINE_STAGES.indexOf(app.stage) >= PIPELINE_STAGES.indexOf('内定'))
            ).length;

            const placements = sourceCandidates.filter(c =>
                c.applications.some(app => !app.isHidden && app.stage === '内定承諾')
            ).length;

            const placementRate = total > 0 ? (placements / total) * 100 : 0;

            return {
                source: sourceName,
                total,
                reachedInterview,
                receivedOffer,
                placements,
                placementRate,
            };
        }).sort((a, b) => b.placements - a.placements || b.total - a.total);

        return stats;
    }, [candidates, allMedia]);

    if (reportData.length === 0) {
        return <p className="no-data-message">分析対象の候補者データがありません。</p>;
    }

    return (
        <div className="all-users-table-container">
            <table className="all-users-table">
                <thead>
                    <tr>
                        <th>媒体</th>
                        <th>候補者数</th>
                        <th>面接到達数</th>
                        <th>内定数</th>
                        <th>決定数</th>
                        <th>決定率</th>
                    </tr>
                </thead>
                <tbody>
                    {reportData.map(stat => (
                        <tr key={stat.source}>
                            <td>{stat.source}</td>
                            <td>{stat.total}</td>
                            <td>{stat.reachedInterview}</td>
                            <td>{stat.receivedOffer}</td>
                            <td>{stat.placements}</td>
                            <td className="progress-cell">
                                <span>{stat.placementRate.toFixed(1)}%</span>
                                <div className="mini-progress-bar">
                                    <div 
                                      className="progress-bar-fill" 
                                      style={{ width: `${stat.placementRate}%`, backgroundColor: 'var(--info-color)' }}
                                    ></div>
                                </div>
                            </td>
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
};


const CandidatePipelineView: React.FC<{
    candidates: Candidate[];
    allMedia: MediaEntry[];
    onSave: (candidate: Candidate) => void;
    onToggleVisibility: (candidateId: string) => void;
    currentUserEmail: string;
    scope: 'personal' | 'all_users' | 'team' | 'user';
    onScopeChange: (scope: 'personal' | 'all_users' | 'team' | 'user') => void;
    teams: Team[];
    selectedTeamId: string | null;
    onSelectedTeamIdChange: (teamId: string | null) => void;
    userOptions: { email: string; label: string }[];
    selectedUserEmail: string | null;
    onSelectedUserEmailChange: (email: string | null) => void;
    isLoadingAggregate: boolean;
}> = ({ candidates, allMedia, onSave, onToggleVisibility, currentUserEmail, scope, onScopeChange, teams, selectedTeamId, onSelectedTeamIdChange, userOptions, selectedUserEmail, onSelectedUserEmailChange, isLoadingAggregate }) => {
    const isOwn = (c: Candidate) => !c.ownerEmail || c.ownerEmail === currentUserEmail;
    const [isModalOpen, setIsModalOpen] = useState(false);
    const [editingCandidate, setEditingCandidate] = useState<Candidate | null>(null);
    const [searchTerm, setSearchTerm] = useState('');
    // 'active' = normal pipeline candidates; 'revival' = 掘り起しリスト entries specifically;
    // 'hidden' = any other hidden (e.g. plainly archived/rejected) candidate.
    const [visibilityFilter, setVisibilityFilter] = useState<'active' | 'hidden' | 'revival'>('active');
    // Multi-select, like selectedStageFilters below — empty means no filtering by 意思決定時期.
    const [selectedDecisionTimingFilters, setSelectedDecisionTimingFilters] = useState<('thisMonth' | 'nextMonth')[]>([]);
    const toggleDecisionTimingFilter = (key: 'thisMonth' | 'nextMonth') => {
        setSelectedDecisionTimingFilters(prev => prev.includes(key) ? prev.filter(k => k !== key) : [...prev, key]);
    };
    const [selectedStageFilters, setSelectedStageFilters] = useState<PipelineStage[]>([]);
    // Narrows the team scope down to specific members (empty = show every member of the
    // selected team). Reset whenever the selected team changes, since a different team's
    // member list makes any previously-checked emails meaningless.
    const [selectedMemberFilters, setSelectedMemberFilters] = useState<string[]>([]);
    useEffect(() => { setSelectedMemberFilters([]); }, [selectedTeamId]);
    const toggleMemberFilter = (email: string) => {
        setSelectedMemberFilters(prev => prev.includes(email) ? prev.filter(e => e !== email) : [...prev, email]);
    };
    const labelByEmailForPipeline = useMemo(() => new Map(userOptions.map(u => [u.email, u.label])), [userOptions]);
    // team.memberEmails is free-typed and may not exactly match the casing of a member's actual
    // sign-in email (canonicalized here via userOptions, which is always Google-cased) — without
    // this, a mis-cased entry would show a raw, unlabeled email AND fail to match candidates'
    // (canonically-tagged) ownerEmail, silently filtering out that member's candidates entirely.
    const canonicalEmailByLower = useMemo(() => new Map(userOptions.map(u => [normalizeEmail(u.email), u.email])), [userOptions]);
    const teamMemberOptions = useMemo(() => {
        if (scope !== 'team' || !selectedTeamId) return [];
        const memberEmails = teams.find(t => t.id === selectedTeamId)?.memberEmails || [];
        const canonicalEmails = Array.from(new Set(
            memberEmails.map(email => canonicalEmailByLower.get(normalizeEmail(email)) || email)
        ));
        return canonicalEmails.map(email => ({ email, label: labelByEmailForPipeline.get(email) || email }));
    }, [scope, selectedTeamId, teams, canonicalEmailByLower, labelByEmailForPipeline]);
    const [sortConfig, setSortConfig] = useState<{ key: PipelineSortKey; direction: 'asc' | 'desc' } | null>({ key: 'createdAt', direction: 'desc'});
    const [expandedCandidateId, setExpandedCandidateId] = useState<string | null>(null);
    const [isReportVisible, setIsReportVisible] = useState(false);
    const [isCalendarVisible, setIsCalendarVisible] = useState(true);
    const [isCompanyPipelineVisible, setIsCompanyPipelineVisible] = useState(false);
    const [isGrossProfitVisible, setIsGrossProfitVisible] = useState(false);
    const [calendarViewDate, setCalendarViewDate] = useState(new Date());
    const [showHiddenApps, setShowHiddenApps] = useState(false);
    
    // State for the new Application Modal
    const [isApplicationModalOpen, setIsApplicationModalOpen] = useState(false);
    const [applicationModalData, setApplicationModalData] = useState<{
        candidate: Candidate | null,
        application: CompanyApplication | null
    }>({ candidate: null, application: null });

    // State for scheduling from a pipeline-calendar date click
    const [scheduleModalDate, setScheduleModalDate] = useState<string | null>(null);

    // State for the 掘り起しリスト registration/edit modal
    const [revivalModalCandidate, setRevivalModalCandidate] = useState<Candidate | null>(null);

    const handleAdd = () => {
        setEditingCandidate(null);
        setIsModalOpen(true);
    };
    
    const handleEdit = (candidate: Candidate) => {
        if (!isOwn(candidate)) return;
        setEditingCandidate(candidate);
        setIsModalOpen(true);
    };

    const handleToggleExpand = (candidateId: string) => {
        const isOpeningNew = expandedCandidateId !== candidateId;
        if (isOpeningNew) {
            setShowHiddenApps(false); // Reset when expanding a new candidate
        }
        setExpandedCandidateId(prevId => (prevId === candidateId ? null : candidateId));
    };

    const handleOpenApplicationModal = (candidate: Candidate, application: CompanyApplication | null) => {
        if (!isOwn(candidate)) return;
        setApplicationModalData({ candidate, application });
        setIsApplicationModalOpen(true);
    };

    const handlePickApplicationForSchedule = (candidate: Candidate, application: CompanyApplication | null) => {
        if (!isOwn(candidate) || !scheduleModalDate) return;
        const prefilled: CompanyApplication = application
            ? { ...application, scheduledDate: scheduleModalDate }
            : { id: `app_${Date.now()}`, companyName: '', stage: '打診', nextAction: '', scheduledDate: scheduleModalDate, memo: '' };
        setScheduleModalDate(null);
        setApplicationModalData({ candidate, application: prefilled });
        setIsApplicationModalOpen(true);
    };

    const handleCloseApplicationModal = () => {
        setIsApplicationModalOpen(false);
        setApplicationModalData({ candidate: null, application: null });
    };

    const handleSaveApplication = (applicationData: CompanyApplication) => {
        const { candidate } = applicationModalData;
        if (!candidate) return;

        const isEditing = candidate.applications.some(app => app.id === applicationData.id);
        
        const updatedApplications = isEditing
            ? candidate.applications.map(app => app.id === applicationData.id ? applicationData : app)
            : [...candidate.applications, applicationData];
        
        const updatedCandidate = {
            ...candidate,
            applications: updatedApplications,
        };

        onSave(updatedCandidate);
    };

    const handleToggleApplicationVisibility = (candidate: Candidate, applicationId: string) => {
        if (!isOwn(candidate)) return;
        const updatedApplications = candidate.applications.map(app =>
            app.id === applicationId ? { ...app, isHidden: !app.isHidden } : app
        );
        const updatedCandidate = { ...candidate, applications: updatedApplications };
        onSave(updatedCandidate);
    };

    const handleOpenRevivalModal = (candidate: Candidate) => {
        if (!isOwn(candidate)) return;
        setRevivalModalCandidate(candidate);
    };

    const handleCloseRevivalModal = () => setRevivalModalCandidate(null);

    // Registering here is what actually hides the candidate — it's parked for future
    // re-engagement rather than being actively pursued, so it drops out of the normal pipeline
    // list and surfaces instead under the 掘り起しリスト filter (and its nextActionDate as a
    // calendar reminder).
    const handleSaveRevival = (data: { nextAction: string; nextActionDate: string }) => {
        if (!revivalModalCandidate) return;
        onSave({ ...revivalModalCandidate, revival: data, isHidden: true });
        handleCloseRevivalModal();
    };

    // "掘り起しリストから解除" — clears the revival entry and restores the candidate to the
    // active pipeline in one step (rather than just un-hiding while leaving stale revival data).
    const handleRemoveFromRevivalList = (candidate: Candidate) => {
        if (!isOwn(candidate)) return;
        const { revival, ...rest } = candidate;
        onSave({ ...rest, isHidden: false });
    };

    const handleExportCSV = () => {
        const dataToExport = candidates.filter(c => !c.isHidden);
        
        if (dataToExport.length === 0) {
            alert('エクスポート対象の候補者がいません。');
            return;
        }

        const mediaFeeRateById = new Map<string, number>(allMedia.map(m => [m.id, m.feeRate || 0] as [string, number]));

        const headers = [
            '氏名', '担当者', '現職企業名', '最終学歴', '現年収(万円)', '希望年収(万円)', '想定年収(万円)',
            '集客媒体', '他エージェント使用状況', '登録日', '概要',
            '応募企業名', '進捗状況', '次アクション', '意思決定時期', '内定確度', '入社確度',
            '想定紹介料(万円)', '想定媒体手数料(万円)', '想定粗利(万円)'
        ];

        const escapeCSV = (str: string | number | undefined | null): string => {
            if (str === null || str === undefined) return '';
            const stringified = String(str);
            if (stringified.includes(',') || stringified.includes('"') || stringified.includes('\n')) {
                const escaped = stringified.replace(/"/g, '""');
                return `"${escaped}"`;
            }
            return stringified;
        };

        const rows: string[][] = [];
        dataToExport.forEach(candidate => {
            const commonData = [
                escapeCSV(candidate.name),
                escapeCSV(candidate.ownerLabel || candidate.ownerEmail || ''),
                escapeCSV(candidate.currentCompany),
                escapeCSV(candidate.education),
                escapeCSV(candidate.currentSalary),
                escapeCSV(candidate.salary),
                escapeCSV(candidate.expectedAnnualSalary),
                escapeCSV(candidate.source),
                candidate.usingOtherAgents ? 'あり' : 'なし',
                new Date(candidate.createdAt).toLocaleDateString('ja-JP'),
                escapeCSV(candidate.summary),
            ];

            const emptyCommonData = Array(commonData.length).fill('');
            const visibleApps = candidate.applications.filter(app => !app.isHidden);

            if (visibleApps.length > 0) {
                visibleApps.forEach((app, index) => {
                    const profit = computeApplicationGrossProfit(candidate, app, mediaFeeRateById);
                    const applicationData = [
                        escapeCSV(app.companyName),
                        escapeCSV(app.stage),
                        escapeCSV(app.nextAction),
                        escapeCSV(app.expectedDecisionDate || ''),
                        escapeCSV(app.offerConfidence || ''),
                        escapeCSV(app.acceptanceConfidence || ''),
                        escapeCSV(profit ? Math.round(profit.revenue) : ''),
                        escapeCSV(profit ? Math.round(profit.cost) : ''),
                        escapeCSV(profit ? Math.round(profit.profit) : ''),
                    ];
                    if (index === 0) {
                        rows.push([...commonData, ...applicationData]);
                    } else {
                        rows.push([...emptyCommonData, ...applicationData]);
                    }
                });
            } else {
                 rows.push([...commonData, '', '', '', '', '', '', '', '', '']);
            }
        });

        // --- Pipeline Summary Calculation: overall total + per-owner breakdown ---
        const STAGE_WEIGHTS: Record<PipelineStage, number> = {
            '打診': 1, '書類選考': 2, '適性検査': 3, 'カジュアル面談': 4, '1次面接': 5, '2次面接': 6,
            '最終面接': 7, '内定': 8, '内定承諾': 9, 'お見送り': 0, '選考辞退': 0,
        };
        const computeStageCounts = (list: Candidate[]): Record<PipelineStage, number> => {
            const counts = PIPELINE_STAGES.reduce((acc, stage) => {
                acc[stage] = 0;
                return acc;
            }, {} as Record<PipelineStage, number>);
            list.forEach(candidate => {
                const visibleApps = candidate.applications.filter(app => !app.isHidden);
                if (visibleApps.length === 0) return;
                let mostAdvancedStage: PipelineStage | null = null;
                let maxWeight = -1;
                visibleApps.forEach(app => {
                    const weight = STAGE_WEIGHTS[app.stage];
                    if (weight > maxWeight) {
                        maxWeight = weight;
                        mostAdvancedStage = app.stage;
                    }
                });
                if (mostAdvancedStage) counts[mostAdvancedStage]++;
            });
            return counts;
        };

        const totalStageCounts = computeStageCounts(dataToExport);

        // Group by owner (falls back to a single "自分" bucket when there's no owner label,
        // i.e. a personal-only export) for the per-person breakdown.
        const ownerOrder: string[] = [];
        const candidatesByOwner = new Map<string, Candidate[]>();
        dataToExport.forEach(candidate => {
            const ownerKey = candidate.ownerLabel || candidate.ownerEmail || '自分';
            if (!candidatesByOwner.has(ownerKey)) {
                candidatesByOwner.set(ownerKey, []);
                ownerOrder.push(ownerKey);
            }
            candidatesByOwner.get(ownerKey)!.push(candidate);
        });

        const summaryRows: string[] = [];
        summaryRows.push('');
        summaryRows.push('パイプラインサマリー（選択中の全員合計）');
        summaryRows.push('ステージ,候補者数');
        PIPELINE_STAGES.forEach(stage => {
            summaryRows.push(`${escapeCSV(stage)},${totalStageCounts[stage]}`);
        });

        summaryRows.push('');
        summaryRows.push('パイプラインサマリー（担当者別）');
        summaryRows.push(['担当者', ...PIPELINE_STAGES].map(escapeCSV).join(','));
        ownerOrder.forEach(ownerKey => {
            const counts = computeStageCounts(candidatesByOwner.get(ownerKey)!);
            summaryRows.push([escapeCSV(ownerKey), ...PIPELINE_STAGES.map(stage => counts[stage])].join(','));
        });

        const candidateCsvPart = [headers.join(','), ...rows.map(row => row.join(','))].join('\n');
        const summaryCsvPart = summaryRows.join('\n');
        const csvContent = '\uFEFF' + candidateCsvPart + '\n\n' + summaryCsvPart;
        
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        link.setAttribute('href', url);
        link.setAttribute('download', `CandidatePipeline_${today}.csv`);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };


    const filteredCandidates = useMemo(() => {
        return candidates.filter(c => {
            const matchesSearch = c.name.toLowerCase().includes(searchTerm.toLowerCase());
            const matchesVisibility = visibilityFilter === 'revival'
                ? !!c.revival
                : visibilityFilter === 'hidden'
                ? c.isHidden === true && !c.revival
                : !c.isHidden;
            const matchesStage = selectedStageFilters.length === 0 || c.applications.some(
                app => !app.isHidden && selectedStageFilters.includes(app.stage)
            );
            const matchesMember = scope !== 'team' || selectedMemberFilters.length === 0
                || (!!c.ownerEmail && selectedMemberFilters.includes(c.ownerEmail));
            const matchesDecisionTiming = selectedDecisionTimingFilters.length === 0 || (() => {
                const nearest = getNearestExpectedDecisionDate(c);
                if (!nearest) return false;
                return selectedDecisionTimingFilters.some(key => isDateInRelativeMonth(nearest, key === 'thisMonth' ? 0 : 1));
            })();
            return matchesSearch && matchesVisibility && matchesStage && matchesMember && matchesDecisionTiming;
        });
    }, [candidates, searchTerm, visibilityFilter, selectedStageFilters, scope, selectedMemberFilters, selectedDecisionTimingFilters]);

    const sortedCandidates = useMemo(() => {
        let sortableItems = [...filteredCandidates];
        if (sortConfig !== null) {
            sortableItems.sort((a, b) => {
                if (sortConfig.key === 'nearestExpectedDecisionDate') {
                    const dateA = getNearestExpectedDecisionDate(a);
                    const dateB = getNearestExpectedDecisionDate(b);
                    // Candidates with no 意思決定時期 set always sort last, regardless of direction —
                    // "unknown" isn't meaningfully earlier or later than a known date.
                    if (!dateA && !dateB) return 0;
                    if (!dateA) return 1;
                    if (!dateB) return -1;
                    return sortConfig.direction === 'asc' ? dateA.localeCompare(dateB) : dateB.localeCompare(dateA);
                }

                const valA = a[sortConfig.key];
                const valB = b[sortConfig.key];

                if (typeof valA === 'string' && typeof valB === 'string') {
                    return sortConfig.direction === 'asc' ? valA.localeCompare(valB, 'ja') : valB.localeCompare(valA, 'ja');
                }
                if (valA < valB) {
                    return sortConfig.direction === 'asc' ? -1 : 1;
                }
                if (valA > valB) {
                    return sortConfig.direction === 'asc' ? 1 : -1;
                }
                return 0;
            });
        }
        return sortableItems;
    }, [filteredCandidates, sortConfig]);

    const requestSort = (key: PipelineSortKey) => {
        let direction: 'asc' | 'desc' = 'asc';
        if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') {
            direction = 'desc';
        }
        setSortConfig({ key, direction });
    };

    const toggleStageFilter = (stage: PipelineStage) => {
        setSelectedStageFilters(prev =>
            prev.includes(stage) ? prev.filter(s => s !== stage) : [...prev, stage]
        );
    };


    const getSortIndicator = (key: PipelineSortKey) => {
        if (!sortConfig || sortConfig.key !== key) return '';
        return sortConfig.direction === 'asc' ? ' ▲' : ' ▼';
    };

    const sortOptions: { key: PipelineSortKey, label: string }[] = [
      { key: 'createdAt', label: '登録日' },
      { key: 'name', label: '氏名' },
      { key: 'currentCompany', label: '現職企業名' },
      { key: 'currentSalary', label: '現職年収' },
      { key: 'nearestExpectedDecisionDate', label: '意思決定時期' },
    ];

    return (
        <section className="pipeline-container" aria-labelledby="pipeline-title">
            {isModalOpen && (
              <CandidateModal
                  onSave={onSave}
                  onClose={() => setIsModalOpen(false)}
                  initialData={editingCandidate}
                  allMedia={allMedia}
              />
            )}
            
            {isApplicationModalOpen && applicationModalData.candidate && (
                <ApplicationModal
                    isOpen={isApplicationModalOpen}
                    onClose={handleCloseApplicationModal}
                    onSave={handleSaveApplication}
                    candidateName={applicationModalData.candidate.name}
                    initialData={applicationModalData.application}
                />
            )}

            {scheduleModalDate && (
                <PipelineDateScheduleModal
                    date={scheduleModalDate}
                    ownCandidates={candidates.filter(c => isOwn(c) && !c.isHidden)}
                    onClose={() => setScheduleModalDate(null)}
                    onPickApplication={handlePickApplicationForSchedule}
                />
            )}

            <RevivalModal
                isOpen={!!revivalModalCandidate}
                onClose={handleCloseRevivalModal}
                onSave={handleSaveRevival}
                candidateName={revivalModalCandidate?.name || ''}
                initialData={revivalModalCandidate?.revival || null}
            />

            <div className="pipeline-header">
                <h2 id="pipeline-title" className="section-title">候補者パイプライン</h2>
                <div className="pipeline-controls">
                    <input
                        type="text"
                        placeholder="候補者名で検索..."
                        value={searchTerm}
                        onChange={e => setSearchTerm(e.target.value)}
                        className="search-input"
                    />
                     <button onClick={handleExportCSV} className="export-button">CSV出力</button>
                </div>
            </div>

            <div className="view-switcher" style={{ marginBottom: '1rem' }}>
                <button onClick={() => onScopeChange('personal')} disabled={scope === 'personal'}>自分</button>
                <button onClick={() => onScopeChange('all_users')} disabled={scope === 'all_users'}>全ユーザー</button>
                <button onClick={() => onScopeChange('team')} disabled={scope === 'team'}>チーム</button>
                <button onClick={() => onScopeChange('user')} disabled={scope === 'user'}>ユーザー別</button>
                {scope === 'team' && (
                    <select
                        value={selectedTeamId || ''}
                        onChange={(e) => onSelectedTeamIdChange(e.target.value || null)}
                        style={{ marginLeft: '0.75rem' }}
                    >
                        <option value="">チームを選択</option>
                        {teams.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
                    </select>
                )}
                {scope === 'user' && (
                    <select
                        value={selectedUserEmail || ''}
                        onChange={(e) => onSelectedUserEmailChange(e.target.value || null)}
                        style={{ marginLeft: '0.75rem' }}
                    >
                        <option value="">ユーザーを選択</option>
                        {userOptions.map(u => <option key={u.email} value={u.email}>{u.label}</option>)}
                    </select>
                )}
            </div>

            {scope !== 'personal' && isLoadingAggregate ? (
                <div className="loading-container">チームメンバーのデータをGoogleドライブから読み込み中...</div>
            ) : scope === 'team' && !selectedTeamId ? (
                <p className="no-data-message">チームを選択してください。チームがまだない場合は「チーム管理」から作成してください。</p>
            ) : scope === 'user' && !selectedUserEmail ? (
                <p className="no-data-message">ユーザーを選択してください。</p>
            ) : (
            <>
            <PipelineDashboard candidates={candidates} />

            <div className="source-effectiveness-section">
                <h3
                    id="gross-profit-title"
                    className="section-title collapsible-header"
                    onClick={() => setIsGrossProfitVisible(prev => !prev)}
                    onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setIsGrossProfitVisible(prev => !prev); }}}
                    role="button"
                    tabIndex={0}
                    aria-expanded={isGrossProfitVisible}
                    aria-controls="gross-profit-content"
                >
                    <span>想定粗利</span>
                    <span className={`toggle-icon ${isGrossProfitVisible ? 'open' : ''}`}>▼</span>
                </h3>
                <div id="gross-profit-content" className={`collapsible-content ${isGrossProfitVisible ? 'open' : ''}`}>
                    <GrossProfitSummary candidates={candidates} allMedia={allMedia} />
                </div>
            </div>

            <div className="source-effectiveness-section">
                <h3
                    id="pipeline-calendar-title"
                    className="section-title collapsible-header"
                    onClick={() => setIsCalendarVisible(prev => !prev)}
                    onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setIsCalendarVisible(prev => !prev); }}}
                    role="button"
                    tabIndex={0}
                    aria-expanded={isCalendarVisible}
                    aria-controls="pipeline-calendar-content"
                >
                    <span>選考スケジュール カレンダー</span>
                    <span className={`toggle-icon ${isCalendarVisible ? 'open' : ''}`}>▼</span>
                </h3>
                <div id="pipeline-calendar-content" className={`collapsible-content ${isCalendarVisible ? 'open' : ''}`}>
                    <PipelineCalendarView
                        candidates={candidates}
                        viewDate={calendarViewDate}
                        currentUserEmail={currentUserEmail}
                        onPrevMonth={() => setCalendarViewDate(d => new Date(d.getFullYear(), d.getMonth() - 1, 1))}
                        onNextMonth={() => setCalendarViewDate(d => new Date(d.getFullYear(), d.getMonth() + 1, 1))}
                        onDayClick={(dateStr) => setScheduleModalDate(dateStr)}
                        onEditApplication={handleOpenApplicationModal}
                        onEditRevival={handleOpenRevivalModal}
                    />
                </div>
            </div>

            <div className="source-effectiveness-section">
                <h3
                    id="company-pipeline-title"
                    className="section-title collapsible-header"
                    onClick={() => setIsCompanyPipelineVisible(prev => !prev)}
                    onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setIsCompanyPipelineVisible(prev => !prev); }}}
                    role="button"
                    tabIndex={0}
                    aria-expanded={isCompanyPipelineVisible}
                    aria-controls="company-pipeline-content"
                >
                    <span>企業別パイプライン状況</span>
                    <span className={`toggle-icon ${isCompanyPipelineVisible ? 'open' : ''}`}>▼</span>
                </h3>
                <div id="company-pipeline-content" className={`collapsible-content ${isCompanyPipelineVisible ? 'open' : ''}`}>
                    <CompanyPipelineView
                        candidates={candidates}
                        currentUserEmail={currentUserEmail}
                        onEditApplication={handleOpenApplicationModal}
                    />
                </div>
            </div>

            <div className="source-effectiveness-section">
                <h3
                    id="source-effectiveness-title"
                    className="section-title collapsible-header"
                    onClick={() => setIsReportVisible(prev => !prev)}
                    onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setIsReportVisible(prev => !prev); }}}
                    role="button"
                    tabIndex={0}
                    aria-expanded={isReportVisible}
                    aria-controls="source-effectiveness-content"
                >
                    <span>集客媒体別 決定率分析</span>
                    <span className={`toggle-icon ${isReportVisible ? 'open' : ''}`}>▼</span>
                </h3>
                <div id="source-effectiveness-content" className={`collapsible-content ${isReportVisible ? 'open' : ''}`}>
                    <SourceEffectivenessReport candidates={candidates} allMedia={allMedia} />
                </div>
            </div>

            {scope === 'personal' && (
            <div className="add-candidate-action-bar">
                <button onClick={handleAdd} className="add-candidate-large-button">
                    + 新規候補者を追加
                </button>
            </div>
            )}

             <div className="pipeline-list-controls">
                <div className="pipeline-sort-controls">
                  <span>選考フェーズで絞り込み:</span>
                  {PIPELINE_STAGES.map(stage => (
                     <button
                        key={stage}
                        onClick={() => toggleStageFilter(stage)}
                        className={selectedStageFilters.includes(stage) ? 'active' : ''}
                      >
                        {stage}
                      </button>
                  ))}
                  {selectedStageFilters.length > 0 && (
                      <button onClick={() => setSelectedStageFilters([])} className="secondary-action-button">
                          クリア
                      </button>
                  )}
                </div>
                <div className="pipeline-sort-controls">
                  <span>並び替え:</span>
                  {sortOptions.map(opt => (
                     <button
                        key={opt.key}
                        onClick={() => requestSort(opt.key)}
                        className={sortConfig?.key === opt.key ? 'active' : ''}
                      >
                        {opt.label}{getSortIndicator(opt.key)}
                      </button>
                  ))}
                </div>
                <div className="pipeline-sort-controls">
                  <span>表示対象:</span>
                  <button onClick={() => setVisibilityFilter('active')} className={visibilityFilter === 'active' ? 'active' : ''}>
                    表示中の候補者
                  </button>
                  <button onClick={() => setVisibilityFilter('revival')} className={visibilityFilter === 'revival' ? 'active' : ''}>
                    掘り起しリスト
                  </button>
                  <button onClick={() => setVisibilityFilter('hidden')} className={visibilityFilter === 'hidden' ? 'active' : ''}>
                    非表示（その他）
                  </button>
                </div>
                <div className="pipeline-sort-controls">
                  <span>意思決定時期で絞り込み:</span>
                  <button
                    onClick={() => toggleDecisionTimingFilter('thisMonth')}
                    className={selectedDecisionTimingFilters.includes('thisMonth') ? 'active' : ''}
                  >
                    今月見込み
                  </button>
                  <button
                    onClick={() => toggleDecisionTimingFilter('nextMonth')}
                    className={selectedDecisionTimingFilters.includes('nextMonth') ? 'active' : ''}
                  >
                    来月見込み
                  </button>
                  {selectedDecisionTimingFilters.length > 0 && (
                      <button onClick={() => setSelectedDecisionTimingFilters([])} className="secondary-action-button">
                          クリア
                      </button>
                  )}
                </div>
            </div>

            {scope === 'team' && teamMemberOptions.length > 0 && (
                <div className="comparison-user-selector">
                    <div className="comparison-user-selector-header">
                        <span>メンバーで絞り込み（未選択の場合は全員を表示）</span>
                        <div style={{ display: 'flex', gap: '0.5rem' }}>
                            <button onClick={() => setSelectedMemberFilters(teamMemberOptions.map(m => m.email))} className="secondary-action-button">全て選択</button>
                            <button onClick={() => setSelectedMemberFilters([])} className="secondary-action-button">選択をクリア</button>
                        </div>
                    </div>
                    <div className="comparison-user-checkbox-list">
                        {teamMemberOptions.map(m => (
                            <label key={m.email} className="comparison-user-checkbox">
                                <input
                                    type="checkbox"
                                    checked={selectedMemberFilters.includes(m.email)}
                                    onChange={() => toggleMemberFilter(m.email)}
                                />
                                {m.label}
                            </label>
                        ))}
                    </div>
                </div>
            )}

            <div className="candidate-list">
                {sortedCandidates.length > 0 ? sortedCandidates.map(c => {
                    const isExpanded = expandedCandidateId === c.id;
                    const visibleApplications = c.applications.filter(app => !app.isHidden);
                    const candidateIsOwn = isOwn(c);

                    return (
                        <div key={c.id} className="candidate-list-item">
                            <div className="candidate-card-main">
                                <div className="candidate-card-header">
                                    <h3
                                      onClick={() => handleEdit(c)}
                                      className={candidateIsOwn ? 'candidate-name-clickable' : ''}
                                      role={candidateIsOwn ? 'button' : undefined}
                                      tabIndex={candidateIsOwn ? 0 : undefined}
                                      onKeyDown={(e) => { if (candidateIsOwn && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); handleEdit(c); }}}
                                      title={candidateIsOwn ? `${c.name}を編集` : c.name}
                                    >
                                      {c.name}
                                    </h3>
                                    {c.ownerLabel && <span className="other-agent-tag">登録者: {c.ownerLabel}</span>}
                                    {c.usingOtherAgents && <span className="other-agent-tag">他エージェント利用中</span>}
                                    {candidateIsOwn && (
                                        <div className="candidate-card-actions">
                                            {visibilityFilter === 'revival' ? (
                                                <>
                                                    <button onClick={() => handleOpenRevivalModal(c)} className="edit-user-button">編集</button>
                                                    <button onClick={() => handleRemoveFromRevivalList(c)} className="secondary-action-button">掘り起しリストから解除</button>
                                                </>
                                            ) : (
                                                <>
                                                    <button onClick={() => handleOpenApplicationModal(c, null)} className="add-selection-button">+ 選考追加</button>
                                                    <button onClick={() => handleEdit(c)} className="edit-user-button">編集</button>
                                                    {visibilityFilter === 'active' && (
                                                        <button onClick={() => handleOpenRevivalModal(c)} className="secondary-action-button">掘り起しリストに追加</button>
                                                    )}
                                                    <button onClick={() => onToggleVisibility(c.id)} className={visibilityFilter === 'hidden' ? "secondary-action-button" : "delete-user-button"}>
                                                        {visibilityFilter === 'hidden' ? '再表示' : '非表示'}
                                                    </button>
                                                </>
                                            )}
                                        </div>
                                    )}
                                </div>
                                {visibilityFilter === 'revival' && c.revival && (
                                    <div className="revival-info-banner">
                                        次アクション: {c.revival.nextAction || '未設定'} / 予定日: {new Date(c.revival.nextActionDate + 'T00:00:00').toLocaleDateString('ja-JP')}
                                    </div>
                                )}
                                <div className="candidate-card-body">
                                    <div className="candidate-key-info">
                                        <div className="key-info-item"><span>現職:</span> {c.currentCompany || 'N/A'}</div>
                                        <div className="key-info-item"><span>学歴:</span> {c.education || 'N/A'}</div>
                                        <div className="key-info-item"><span>現年収:</span> {c.currentSalary ? `${c.currentSalary}万円` : 'N/A'}</div>
                                        <div className="key-info-item"><span>媒体:</span> {c.source || 'N/A'}</div>
                                    </div>
                                     <div className="candidate-application-summary">
                                        <span className="summary-label">選考状況 ({visibleApplications.length}件):</span>
                                        <div className="summary-badges">
                                            {visibleApplications.length > 0 ? (
                                                visibleApplications.map(app => (
                                                    <span
                                                        key={app.id}
                                                        className={`status-badge ${candidateIsOwn ? 'status-badge-clickable' : ''}`}
                                                        style={{'--badge-color': STAGE_COLOR_MAP[app.stage]} as React.CSSProperties}
                                                        title={candidateIsOwn ? `${app.companyName}: ${app.stage}（クリックして更新）` : `${app.companyName}: ${app.stage}`}
                                                        onClick={candidateIsOwn ? (e) => { e.stopPropagation(); handleOpenApplicationModal(c, app); } : undefined}
                                                        role={candidateIsOwn ? 'button' : undefined}
                                                        tabIndex={candidateIsOwn ? 0 : undefined}
                                                        onKeyDown={candidateIsOwn ? (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); e.stopPropagation(); handleOpenApplicationModal(c, app); } } : undefined}
                                                    >
                                                        {app.companyName}: {app.stage}
                                                    </span>
                                                ))
                                            ) : (
                                                <span className="no-status">未登録</span>
                                            )}
                                        </div>
                                    </div>
                                </div>
                                 <div 
                                    className="candidate-card-footer"
                                    onClick={() => handleToggleExpand(c.id)}
                                    onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleToggleExpand(c.id); }}}
                                    role="button"
                                    tabIndex={0}
                                    aria-expanded={isExpanded}
                                >
                                    <span>{isExpanded ? '詳細を閉じる' : '詳細を表示'}</span>
                                    <span className={`toggle-icon ${isExpanded ? 'open' : ''}`}>▼</span>
                                </div>
                            </div>
                            <div className={`candidate-detail-content ${isExpanded ? 'open' : ''}`}>
                                <div className="candidate-detail-view">
                                    <div className="candidate-info-section">
                                      <div className="candidate-info-item">
                                          <span className="info-label">希望年収</span>
                                          <span className="info-value">{c.salary ? `${c.salary}万円` : '未設定'}</span>
                                      </div>
                                      <div className="candidate-info-item">
                                          <span className="info-label">登録日</span>
                                          <span className="info-value">{new Date(c.createdAt).toLocaleDateString('ja-JP')}</span>
                                      </div>
                                      <div className="candidate-info-item">
                                          <span className="info-label">レジュメ</span>
                                          <span className="info-value">{c.resumeFiles && c.resumeFiles.length > 0 ? c.resumeFiles.map(f => f.name).join(', ') : '未登録'}</span>
                                      </div>
                                      <div className="candidate-info-item">
                                          <span className="info-label">面談音声ファイル</span>
                                          <span className="info-value">{c.interviewAudioFile ? c.interviewAudioFile.name : '未登録'}</span>
                                      </div>
                                      <div className="candidate-info-item summary-item">
                                          <span className="info-label">概要 (レジュメより)</span>
                                          <p className="info-value summary-text">{c.summary || '概要がありません。'}</p>
                                      </div>
                                      <div className="candidate-info-item summary-item">
                                          <span className="info-label">面談要約 (AI)</span>
                                          <p className="info-value summary-text">{c.interviewSummary || '面談要約はありません。'}</p>
                                      </div>
                                    </div>
                                    
                                    <div className="detail-view-controls">
                                      <h5 className="detail-view-subtitle">選考詳細</h5>
                                      <div className="single-checkbox">
                                          <input
                                              type="checkbox"
                                              id={`show-hidden-apps-${c.id}`}
                                              checked={showHiddenApps}
                                              onChange={e => setShowHiddenApps(e.target.checked)}
                                          />
                                          <label htmlFor={`show-hidden-apps-${c.id}`}>非表示の選考状況を表示</label>
                                      </div>
                                    </div>

                                    {(() => {
                                      const applicationsToShow = c.applications.filter(app => showHiddenApps ? true : !app.isHidden);
                                      return applicationsToShow.length > 0 ? (
                                          <div className="detail-application-grid">
                                              {applicationsToShow.map(app => (
                                                  <div key={app.id} className={`detail-application-card ${app.isHidden ? 'is-hidden' : ''}`}>
                                                      <div className="detail-card-header">
                                                          <strong>{app.companyName}</strong>
                                                          {candidateIsOwn && (
                                                          <div className="detail-card-actions">
                                                            <button onClick={() => handleOpenApplicationModal(c, app)} className="edit-user-button">編集</button>
                                                            <button onClick={() => handleToggleApplicationVisibility(c, app.id)} className={app.isHidden ? "secondary-action-button" : "delete-user-button"}>
                                                                {app.isHidden ? '再表示' : '非表示'}
                                                            </button>
                                                          </div>
                                                          )}
                                                      </div>
                                                      <div className="detail-card-body">
                                                          <div className="detail-card-item">
                                                              <span>進捗状況:</span>
                                                              <span
                                                                  className={`status-badge ${candidateIsOwn ? 'status-badge-clickable' : ''}`}
                                                                  style={{'--badge-color': STAGE_COLOR_MAP[app.stage]} as React.CSSProperties}
                                                                  onClick={candidateIsOwn ? () => handleOpenApplicationModal(c, app) : undefined}
                                                                  role={candidateIsOwn ? 'button' : undefined}
                                                                  tabIndex={candidateIsOwn ? 0 : undefined}
                                                                  title={candidateIsOwn ? 'クリックして選考状況を更新' : undefined}
                                                                  onKeyDown={candidateIsOwn ? (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleOpenApplicationModal(c, app); } } : undefined}
                                                              >
                                                                  {app.stage}
                                                              </span>
                                                          </div>
                                                          <div className="detail-card-item">
                                                              <span>次アクション:</span>
                                                              <span>{app.nextAction || '未設定'}</span>
                                                          </div>
                                                          <div className="detail-card-item">
                                                              <span>選考予定日:</span>
                                                              <span>
                                                                {app.scheduledDate
                                                                  ? `${new Date(app.scheduledDate + 'T00:00:00').toLocaleDateString('ja-JP')}${app.scheduledTime ? ` ${app.scheduledTime}` : ''}`
                                                                  : '未設定'}
                                                              </span>
                                                          </div>
                                                          <div className="detail-card-item">
                                                              <span>意思決定時期:</span>
                                                              <span>{app.expectedDecisionDate ? new Date(app.expectedDecisionDate + 'T00:00:00').toLocaleDateString('ja-JP') : '未設定'}</span>
                                                          </div>
                                                          <div className="detail-card-item">
                                                              <span>内定確度 / 入社確度:</span>
                                                              <span>{app.offerConfidence || '未設定'} / {app.acceptanceConfidence || '未設定'}</span>
                                                          </div>
                                                          {app.memo && (
                                                              <div className="detail-card-item detail-card-item-memo">
                                                                  <span>メモ:</span>
                                                                  <span className="detail-card-memo-text">{app.memo}</span>
                                                              </div>
                                                          )}
                                                      </div>
                                                  </div>
                                              ))}
                                          </div>
                                      ) : (
                                          <p className="no-data-message" style={{padding: '1rem 0'}}>選考中の企業情報はありません。</p>
                                      )
                                    })()}

                                    <div className="add-application-trigger">
                                        <button onClick={() => handleOpenApplicationModal(c, null)} className="add-button">+ 選考を追加</button>
                                    </div>
                                </div>
                            </div>
                        </div>
                    )
                }) : (
                    candidates.length === 0 ? (
                        <div className="candidate-list-empty-state">
                            <h3>まだ候補者が登録されていません</h3>
                            <p>最初の候補者情報を登録して、パイプライン管理を始めましょう。</p>
                        </div>
                    ) : (
                        <div className="no-data-message">
                            検索条件に一致する候補者が見つかりません。
                        </div>
                    )
                )}
            </div>
            </>
            )}
        </section>
    );
};


interface FunnelStageDef {
  key: string;
  label: string;
  getValue: (monthlyTotals: KpiTotals, allMedia: MediaEntry[]) => number;
}

const MEDIA_FUNNEL_STAGE_LABELS: Record<string, string> = {
  scoutsSent: 'スカウト数',
  scoutReplies: 'スカウト返信数',
  effectiveReplies: '有効返信数',
  documentsCollected: '書類回収数',
  effectiveDocumentsCollected: '有効書類回収数',
  initialInterviews: '初回面談数',
  effectiveInitialInterviews: '初回有効面談数',
};

// Covers every item entered in the daily entry form: the media-scoped sourcing-side steps
// (aggregated across all media via getTotalFromLump, in the same order as the daily entry
// table's columns) come first, since they happen before a candidate is ever submitted to a
// client, followed by GENERAL_KPIS (候補者推薦数 onward).
const FUNNEL_STAGES: FunnelStageDef[] = [
  ...MEDIA_KPI_SUFFIXES.map(suffix => ({
    key: suffix as string,
    label: MEDIA_FUNNEL_STAGE_LABELS[suffix],
    getValue: (totals: KpiTotals, allMedia: MediaEntry[]) => getTotalFromLump(totals, `_${suffix}`, allMedia),
  })),
  ...(Object.keys(GENERAL_KPIS) as Array<keyof typeof GENERAL_KPIS>).map(key => ({
    key: key as string,
    label: GENERAL_KPIS[key].label,
    getValue: (totals: KpiTotals) => totals[key] || 0,
  })),
];

// The media-scoped subset of FUNNEL_STAGES — used when viewing a single media's own funnel,
// since GENERAL_KPIS entries (候補者推薦数 onward) aren't tagged by sourcing media at all and
// would show an identical value regardless of which media is selected, which would be
// misleading in a per-media view.
const MEDIA_FUNNEL_STAGES = FUNNEL_STAGES.filter(s => (MEDIA_KPI_SUFFIXES as readonly string[]).includes(s.key));

/** Rate at index i is "stage[i] / stage[i-1]"; index 0 has no previous stage, so it's null. */
const computeConversionRates = (values: number[]): (number | null)[] =>
  values.map((v, i) => (i === 0 ? null : (values[i - 1] > 0 ? (v / values[i - 1]) * 100 : null)));

const findBottleneckIndex = (rates: (number | null)[]): number => {
  let minIdx = -1;
  let minVal = Infinity;
  rates.forEach((r, i) => {
    if (r !== null && r < minVal) {
      minVal = r;
      minIdx = i;
    }
  });
  return minIdx;
};

const FunnelAnalysisSection: React.FC<{
  users: string[];
  allUsersData: Record<string, UserData>;
  allMedia: MediaEntry[];
  periodOverride?: { start: Date; end: Date } | null;
  periodLabel?: string;
  perUserProgressStats?: any[];
  grossProfitStageTotals?: StageGrossProfit[];
}> = ({ users, allUsersData, allMedia, periodOverride = null, periodLabel = '今月', perUserProgressStats, grossProfitStageTotals }) => {
  const [isVisible, setIsVisible] = useState(false);
  const [aiSuggestion, setAiSuggestion] = useState('');
  const [isGeneratingSuggestion, setIsGeneratingSuggestion] = useState(false);
  const [userQuestion, setUserQuestion] = useState('');
  const [historyMonths, setHistoryMonths] = useState(6);
  // 'all' sums across every media (existing behavior); a specific media id scopes the funnel to
  // just that media's own numbers, using MEDIA_FUNNEL_STAGES only since GENERAL_KPIS isn't
  // per-media data.
  const [selectedMediaId, setSelectedMediaId] = useState<string>('all');

  const visibleStages = selectedMediaId === 'all' ? FUNNEL_STAGES : MEDIA_FUNNEL_STAGES;
  const mediaScope = useMemo(
    () => (selectedMediaId === 'all' ? allMedia : allMedia.filter(m => m.id === selectedMediaId)),
    [selectedMediaId, allMedia]
  );
  const mediaLabel = selectedMediaId === 'all' ? '全媒体合計' : (allMedia.find(m => m.id === selectedMediaId)?.name || selectedMediaId);

  const { totalStageValues, perUserStageValues } = useMemo(() => {
    const perUser: Record<string, number[]> = {};
    const totals = visibleStages.map(() => 0);
    users.forEach(email => {
      const data = allUsersData[email];
      if (!data) return;
      const periodTotals = periodOverride
        ? calculateTotalsForRange(data.entries || [], allMedia, periodOverride.start, periodOverride.end)
        : calculateMonthlyTotals(data.entries || [], allMedia);
      const values = visibleStages.map(stage => stage.getValue(periodTotals, mediaScope));
      perUser[email] = values;
      values.forEach((v, i) => { totals[i] += v; });
    });
    return { totalStageValues: totals, perUserStageValues: perUser };
  }, [users, allUsersData, allMedia, periodOverride, visibleStages, mediaScope]);

  const totalConversionRates = useMemo(() => computeConversionRates(totalStageValues), [totalStageValues]);
  const bottleneckIndex = useMemo(() => findBottleneckIndex(totalConversionRates), [totalConversionRates]);

  // A separate, always-monthly-bucketed trend series (independent of periodOverride, which is
  // just a single snapshot window) — this is what lets the AI compare "this month vs last
  // month" or extrapolate a simple trend, neither of which is possible from one totals object.
  const monthlyHistory = useMemo(() => {
    const now = new Date();
    const months: { label: string; start: Date; end: Date }[] = [];
    for (let i = historyMonths - 1; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      months.push({
        label: `${d.getFullYear()}年${d.getMonth() + 1}月`,
        start: new Date(d.getFullYear(), d.getMonth(), 1),
        end: new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59),
      });
    }
    return months.map(({ label, start, end }) => {
      const values = visibleStages.map(() => 0);
      users.forEach(email => {
        const data = allUsersData[email];
        if (!data) return;
        const rangeTotals = calculateTotalsForRange(data.entries || [], allMedia, start, end);
        visibleStages.forEach((stage, i) => { values[i] += stage.getValue(rangeTotals, mediaScope); });
      });
      return { label, values };
    });
  }, [users, allUsersData, allMedia, historyMonths, visibleStages, mediaScope]);

  const handleAskAi = async () => {
    setIsGeneratingSuggestion(true);
    setAiSuggestion('');
    try {
      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY as string });
      const lines: string[] = [`【全体ファネル実績（${periodLabel}・${mediaLabel}・全ユーザー合計）】`];
      visibleStages.forEach((stage, i) => {
        const rate = totalConversionRates[i];
        lines.push(`${stage.label}: ${totalStageValues[i]}件${rate !== null ? `（前段階からの歩留まり: ${rate.toFixed(1)}%）` : ''}`);
      });
      lines.push('', `【ユーザー別ファネル実績（${periodLabel}・${mediaLabel}）】`);
      users.forEach(email => {
        const data = allUsersData[email];
        if (!data) return;
        const label = data.displayName || email;
        const values = perUserStageValues[email] || visibleStages.map(() => 0);
        const rates = computeConversionRates(values);
        const parts = visibleStages.map((stage, i) =>
          `${stage.label} ${values[i]}件${rates[i] !== null ? `(${rates[i]!.toFixed(1)}%)` : ''}`
        );
        lines.push(`${label}: ${parts.join(', ')}`);
      });

      if (perUserProgressStats && perUserProgressStats.length > 0) {
        lines.push('', `【ユーザー別 書類回収・面談・その他KPI実績（${periodLabel}）】`);
        perUserProgressStats.forEach(stat => {
          if (!stat.userData) return;
          const generalParts = (stat.generalKpis || []).map((g: any) => `${g.label} ${g.value}件`).join(', ');
          lines.push(
            `${stat.displayName}: スカウト返信率 ${stat.replyRate.toFixed(1)}%(${stat.replies}/${stat.sent}), ` +
            `書類回収 ${stat.documentsCollected}件, 初回面談 ${stat.initialInterviews}件, ${generalParts}`
          );
        });
      }

      if (grossProfitStageTotals && grossProfitStageTotals.length > 0) {
        lines.push('', '【パイプライン想定粗利（フェーズ別・現在時点の全件）】');
        grossProfitStageTotals.forEach(s => {
          lines.push(`${s.stage}: ${s.count}件（うち粗利算出可能 ${s.estimableCount}件）, 想定粗利合計 ${formatManYen(s.profit)}`);
        });
      }

      if (monthlyHistory.length > 1) {
        lines.push('', `【月別実績推移（${mediaLabel}・全ユーザー合計・直近${monthlyHistory.length}ヶ月、最後の行が最新月）】`);
        monthlyHistory.forEach(({ label, values }) => {
          const parts = visibleStages.map((stage, i) => `${stage.label} ${values[i]}件`);
          lines.push(`${label}: ${parts.join(', ')}`);
        });
      }

      const question = userQuestion.trim() ||
        '各ステップの歩留まり（前段階からの通過率）を踏まえて、ボトルネックとなっている工程を指摘し、改善のための具体的な施策を日本語で3〜5点、簡潔に提案してください。';
      const prompt =
        `あなたは人材紹介会社の営業KPIを分析するアシスタントです。以下は採用エージェントの求人紹介パイプラインにおける${periodLabel}の実績データと、月別の推移データです。` +
        `このデータに基づいて、次の質問に日本語で分かりやすく回答してください。月別推移データがある場合は月同士の比較や今後の傾向の推定に活用してかまいませんが、` +
        `将来予測をする場合は必ず「過去の傾向に基づく推定であり保証ではない」旨を明記してください。データにない情報については推測せず、その旨を述べてください。\n\n` +
        `${lines.join('\n')}\n\n【質問】\n${question}`;

      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: prompt,
      });
      setAiSuggestion(response.text.trim());
    } catch (error) {
      console.error('Error generating AI insight:', error);
      alert('AIによる分析中にエラーが発生しました。');
    } finally {
      setIsGeneratingSuggestion(false);
    }
  };

  return (
    <section aria-labelledby="funnel-analysis-title">
      <h2
        id="funnel-analysis-title"
        className="section-title collapsible-header"
        onClick={() => setIsVisible(prev => !prev)}
        onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setIsVisible(prev => !prev); } }}
        role="button"
        tabIndex={0}
        aria-expanded={isVisible}
        aria-controls="funnel-analysis-content"
      >
        <span>歩留まり分析（ファネル）</span>
        <span className={`toggle-icon ${isVisible ? 'open' : ''}`}>▼</span>
      </h2>
      <div id="funnel-analysis-content" className={`collapsible-content ${isVisible ? 'open' : ''}`}>
        <div className="pipeline-sort-controls" style={{ marginBottom: '1rem' }}>
          <span>媒体で切り替え:</span>
          <button
            type="button"
            className={selectedMediaId === 'all' ? 'active' : ''}
            onClick={() => setSelectedMediaId('all')}
          >
            全媒体合計
          </button>
          {allMedia.map(m => (
            <button
              key={m.id}
              type="button"
              className={selectedMediaId === m.id ? 'active' : ''}
              onClick={() => setSelectedMediaId(m.id)}
            >
              {m.name}{m.isArchived ? '（アーカイブ済み）' : ''}
            </button>
          ))}
        </div>
        {selectedMediaId !== 'all' && (
          <p className="gmail-scout-message">
            単一媒体を表示中は、候補者推薦数以降（媒体を区別しないGENERAL_KPIS項目）はファネルから除外しています。
          </p>
        )}

        <h3 className="sub-section-title">全体ファネル（全ユーザー合計・{periodLabel}・{mediaLabel}）</h3>
        <div className="all-users-table-container">
          <table className="all-users-table">
            <thead>
              <tr><th>指標</th><th>件数</th><th>前段階からの歩留まり</th></tr>
            </thead>
            <tbody>
              {visibleStages.map((stage, i) => (
                <tr key={stage.key}>
                  <td>{stage.label}</td>
                  <td>{totalStageValues[i]}</td>
                  <td style={i === bottleneckIndex ? { color: 'crimson', fontWeight: 'bold' } : undefined}>
                    {totalConversionRates[i] !== null
                      ? `${totalConversionRates[i]!.toFixed(1)}%${i === bottleneckIndex ? '（ボトルネック）' : ''}`
                      : '—'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <h3 className="sub-section-title" style={{ marginTop: '1.5rem' }}>ユーザー別ファネル比較（{periodLabel}・{mediaLabel}）</h3>
        <div className="all-users-table-container">
          <table className="all-users-table">
            <thead>
              <tr>
                <th>ユーザー</th>
                <th>{visibleStages[0].label}（件数）</th>
                {visibleStages.slice(1).map(stage => <th key={stage.key}>{stage.label}（歩留まり %）</th>)}
              </tr>
            </thead>
            <tbody>
              {users.map(email => {
                const data = allUsersData[email];
                if (!data) return null;
                const label = data.displayName || email;
                const values = perUserStageValues[email] || visibleStages.map(() => 0);
                const rates = computeConversionRates(values);
                const userBottleneckIndex = findBottleneckIndex(rates);
                return (
                  <tr key={email}>
                    <td>{label}</td>
                    <td>{values[0]}</td>
                    {rates.slice(1).map((rate, idx) => {
                      const i = idx + 1;
                      return (
                        <td key={i} style={i === userBottleneckIndex ? { color: 'crimson', fontWeight: 'bold' } : undefined}>
                          {rate !== null ? `${rate.toFixed(1)}%` : '—'}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        <div className="ai-insight-panel">
          <label htmlFor="ai-insight-question">AIに質問・分析を依頼（空欄の場合は改善提案を生成します）</label>
          <textarea
            id="ai-insight-question"
            value={userQuestion}
            onChange={(e) => setUserQuestion(e.target.value)}
            placeholder="例：先月と比較してどう変化した？／今後3ヶ月の返信数を予測して"
            rows={2}
          />
          <div className="ai-insight-quick-questions">
            <button type="button" className="secondary-action-button" onClick={() => setUserQuestion('直近数ヶ月の月別実績を比較し、良くなっている点・悪化している点を指摘してください。')}>月次比較</button>
            <button type="button" className="secondary-action-button" onClick={() => setUserQuestion('月別実績の推移から、今後の傾向を推定してください（保証ではない推定である旨を明記してください）。')}>今後の予測</button>
            <button type="button" className="secondary-action-button" onClick={() => setUserQuestion('')}>質問をクリア</button>
          </div>
          <label htmlFor="ai-insight-history-months" style={{ marginTop: '0.25rem' }}>
            比較・予測に使う月別履歴データ:
            <select
              id="ai-insight-history-months"
              value={historyMonths}
              onChange={(e) => setHistoryMonths(Number(e.target.value))}
              style={{ marginLeft: '0.5rem' }}
            >
              <option value={3}>直近3ヶ月</option>
              <option value={6}>直近6ヶ月</option>
              <option value={12}>直近12ヶ月</option>
            </select>
          </label>
          <button type="button" onClick={handleAskAi} disabled={isGeneratingSuggestion} className="submit-button">
            {isGeneratingSuggestion ? 'AIが分析中...' : (userQuestion.trim() ? 'AIに聞く' : 'AIに改善提案をもらう')}
          </button>
          {aiSuggestion && (
            <div className="summary-wrapper" style={{ marginTop: '1rem' }}>
              <p className="info-value summary-text" style={{ whiteSpace: 'pre-wrap' }}>{aiSuggestion}</p>
            </div>
          )}
        </div>
      </div>
    </section>
  );
};

const formatPeriodDate = (d: Date): string => `${d.getMonth() + 1}/${d.getDate()}`;

const AllUsersDashboard: React.FC<{
  users: string[];
  allUsersData: Record<string, UserData>;
  allMedia: MediaEntry[];
  dayOfWeekReplyRateData: any | null;
  weekStartDate: Date;
  onPrevWeek: () => void;
  onNextWeek: () => void;
  visibility: { progress: boolean; dowRate: boolean; weeklySummary: boolean; memberWeeklySummary: boolean; grossProfit: boolean; monthlyTrend: boolean };
  toggleSection: (key: 'allUsersProgress' | 'allUsersDayOfWeekRate' | 'allUsersWeeklySummary' | 'allUsersMemberWeeklySummary' | 'allUsersGrossProfit' | 'allUsersMonthlyTrend') => void;
  showGrossProfit?: boolean;
  periodOverride?: { start: Date; end: Date } | null;
}> = ({ users, allUsersData, allMedia, dayOfWeekReplyRateData, weekStartDate, onPrevWeek, onNextWeek, visibility, toggleSection, showGrossProfit = true, periodOverride = null }) => {
  const activeMedia = useMemo(() => allMedia.filter(m => !m.isArchived), [allMedia]);
  const periodLabel = periodOverride ? `${formatPeriodDate(periodOverride.start)}〜${formatPeriodDate(periodOverride.end)}` : '今月';
  const { data: aggregateWeeklyData, weeklyKpiTargets: aggregateWeeklyKpiTargets } = useMemo(
    () => computeAggregateWeeklyData(users, allUsersData, activeMedia, weekStartDate),
    [users, allUsersData, activeMedia, weekStartDate]
  );
  const perMemberWeeklyData = useMemo(
    () => users.map(user => ({
      user,
      displayName: allUsersData[user]?.displayName || user,
      ...computeAggregateWeeklyData([user], allUsersData, activeMedia, weekStartDate),
    })),
    [users, allUsersData, activeMedia, weekStartDate]
  );
  const candidatesAcrossUsers = useMemo(
    () => users.flatMap(user => allUsersData[user]?.candidates || []),
    [users, allUsersData]
  );
  const perUserTrendEntries = useMemo(
    () => users.map(user => ({ label: allUsersData[user]?.displayName || user, entries: allUsersData[user]?.entries || [] })),
    [users, allUsersData]
  );
  const grossProfitStageTotals = useMemo(
    () => (showGrossProfit ? computeGrossProfitByStage(candidatesAcrossUsers, allMedia) : undefined),
    [showGrossProfit, candidatesAcrossUsers, allMedia]
  );

  // Hoisted out of the table's render loop so the same per-user period totals can also be
  // handed to the AI panel as context, instead of duplicating this calculation twice.
  const perUserProgressStats = useMemo(() => users.map(user => {
    const userData = allUsersData[user];
    if (!userData) return { user, displayName: user, userData: null as null };
    const displayName = userData.displayName || user;
    const totals = periodOverride
      ? calculateTotalsForRange(userData.entries || [], allMedia, periodOverride.start, periodOverride.end)
      : calculateMonthlyTotals(userData.entries || [], allMedia);
    const kpiTargets = { ...buildDefaultKpiTargets(allMedia), ...(userData.kpiTargets || {}) };

    const sent = getTotalFromLump(totals, '_scoutsSent', allMedia);
    const replies = getTotalFromLump(totals, '_scoutReplies', allMedia);
    const effectiveReplies = getTotalFromLump(totals, '_effectiveReplies', allMedia);
    const replyRate = sent > 0 ? (replies / sent) * 100 : 0;
    const effectiveReplyRate = replies > 0 ? (effectiveReplies / replies) * 100 : 0;

    // Actuals (numerator) include archived media — their historical performance still
    // counts — but targets (denominator) must only sum activeMedia; see note elsewhere in
    // this file on why summing allMedia for targets silently mismatches the settings form.
    const documentsCollected = getTotalFromLump(totals, '_documentsCollected', allMedia);
    const documentsCollectedTarget = getTotalFromLump(kpiTargets, '_documentsCollected', activeMedia);
    const effectiveDocumentsCollected = getTotalFromLump(totals, '_effectiveDocumentsCollected', allMedia);
    const effectiveDocumentsCollectedTarget = getTotalFromLump(kpiTargets, '_effectiveDocumentsCollected', activeMedia);

    const initialInterviews = getTotalFromLump(totals, '_initialInterviews', allMedia);
    const effectiveInitialInterviews = getTotalFromLump(totals, '_effectiveInitialInterviews', allMedia);
    const effectiveInterviewRate = initialInterviews > 0 ? (effectiveInitialInterviews / initialInterviews) * 100 : 0;
    const initialInterviewsTarget = getTotalFromLump(kpiTargets, '_initialInterviews', activeMedia);

    const generalKpis = (Object.keys(GENERAL_KPIS) as Array<keyof typeof GENERAL_KPIS>).map(key => ({
      key, label: GENERAL_KPIS[key].label, value: totals[key] || 0, target: kpiTargets[key] || 0,
    }));

    return {
      user, displayName, userData,
      sent, replies, effectiveReplies, replyRate, effectiveReplyRate,
      documentsCollected, documentsCollectedTarget, effectiveDocumentsCollected, effectiveDocumentsCollectedTarget,
      initialInterviews, effectiveInitialInterviews, effectiveInterviewRate, initialInterviewsTarget,
      generalKpis,
    };
  }), [users, allUsersData, allMedia, activeMedia, periodOverride]);

  return (
    <>
      <section aria-labelledby="all-users-dashboard-title">
        <h2 
            id="all-users-dashboard-title" 
            className="section-title collapsible-header"
            onClick={() => toggleSection('allUsersProgress')}
            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('allUsersProgress'); } }}
            role="button"
            tabIndex={0}
            aria-expanded={visibility.progress}
            aria-controls="all-users-progress-content"
        >
          <span>全ユーザーの進捗（{periodLabel}）</span>
          <span className={`toggle-icon ${visibility.progress ? 'open' : ''}`}>▼</span>
        </h2>
        <div id="all-users-progress-content" className={`collapsible-content ${visibility.progress ? 'open' : ''}`}>
          {periodOverride && (
            <p className="gmail-scout-message">指定期間の実績を表示しています。目標は月次で設定されているため、この表示では目標・達成率は表示していません。</p>
          )}
          <div className="all-users-table-container">
            <table className="all-users-table">
              <thead>
                <tr>
                  <th>ユーザー</th>
                  <th>スカウト返信率</th>
                  <th>有効返信率</th>
                  <th>書類回収合計</th>
                  <th>有効書類回収合計</th>
                  <th>初回面談数</th>
                  <th>有効面談率</th>
                  {(Object.keys(GENERAL_KPIS) as Array<keyof typeof GENERAL_KPIS>).map(key => (
                      <th key={key}>{GENERAL_KPIS[key].label}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {perUserProgressStats.map(stat => {
                  if (!stat.userData) {
                    return <tr key={stat.user}><td colSpan={Object.keys(GENERAL_KPIS).length + 7}>{stat.displayName}のデータがありません。</td></tr>;
                  }
                  const {
                    user, displayName, sent, replies, effectiveReplies, replyRate, effectiveReplyRate,
                    documentsCollected, documentsCollectedTarget, effectiveDocumentsCollected, effectiveDocumentsCollectedTarget,
                    initialInterviews, effectiveInitialInterviews, effectiveInterviewRate, initialInterviewsTarget,
                    generalKpis,
                  } = stat;
                  const documentsCollectedProgress = documentsCollectedTarget > 0 ? Math.min((documentsCollected / documentsCollectedTarget) * 100, 100) : 0;
                  const effectiveDocumentsCollectedProgress = effectiveDocumentsCollectedTarget > 0 ? Math.min((effectiveDocumentsCollected / effectiveDocumentsCollectedTarget) * 100, 100) : 0;
                  const initialInterviewsProgress = initialInterviewsTarget > 0 ? Math.min((initialInterviews / initialInterviewsTarget) * 100, 100) : 0;

                  return (
                    <tr key={user}>
                      <td>{displayName}</td>
                      <td className="progress-cell">
                        <span>{replyRate.toFixed(1)}%</span>
                        <div className="mini-progress-bar">
                          <div className="progress-bar-fill" style={{ width: `${Math.min(replyRate, 100)}%` }}></div>
                        </div>
                        <small>({replies}/{sent})</small>
                      </td>
                      <td className="progress-cell">
                        <span>{effectiveReplyRate.toFixed(1)}%</span>
                        <div className="mini-progress-bar">
                          <div className="progress-bar-fill" style={{ width: `${Math.min(effectiveReplyRate, 100)}%`, backgroundColor: 'var(--info-color)' }}></div>
                        </div>
                        <small>({effectiveReplies}/{replies})</small>
                      </td>
                      <td className="progress-cell">
                          <span>{documentsCollected}{!periodOverride && ` / ${documentsCollectedTarget}`}</span>
                          {!periodOverride && (
                            <div className="mini-progress-bar">
                                <div className="progress-bar-fill" style={{ width: `${documentsCollectedProgress}%` }}></div>
                            </div>
                          )}
                      </td>
                       <td className="progress-cell">
                          <span>{effectiveDocumentsCollected}{!periodOverride && ` / ${effectiveDocumentsCollectedTarget}`}</span>
                          {!periodOverride && (
                            <div className="mini-progress-bar">
                                <div className="progress-bar-fill" style={{ width: `${effectiveDocumentsCollectedProgress}%`, backgroundColor: 'var(--info-color)' }}></div>
                            </div>
                          )}
                      </td>
                      <td className="progress-cell">
                          <span>{initialInterviews}{!periodOverride && ` / ${initialInterviewsTarget}`}</span>
                          {!periodOverride && (
                            <div className="mini-progress-bar">
                                <div className="progress-bar-fill" style={{ width: `${initialInterviewsProgress}%` }}></div>
                            </div>
                          )}
                      </td>
                      <td className="progress-cell">
                        <span>{effectiveInterviewRate.toFixed(1)}%</span>
                        <div className="mini-progress-bar">
                          <div className="progress-bar-fill" style={{ width: `${Math.min(effectiveInterviewRate, 100)}%`, backgroundColor: 'var(--info-color)' }}></div>
                        </div>
                        <small>({effectiveInitialInterviews}/{initialInterviews})</small>
                      </td>
                      {generalKpis.map(({ key, value, target }) => {
                          const progress = target > 0 ? Math.min((value / target) * 100, 100) : 0;
                          return (
                              <td key={key} className="progress-cell">
                                  <span>{value}{!periodOverride && ` / ${target}`}</span>
                                  {!periodOverride && (
                                    <div className="mini-progress-bar">
                                        <div className="progress-bar-fill" style={{ width: `${progress}%` }}></div>
                                    </div>
                                  )}
                              </td>
                          );
                      })}
                    </tr>
                  );
                })}
                 {users.length === 0 && (
                    <tr>
                        <td colSpan={Object.keys(GENERAL_KPIS).length + 7}>表示するユーザーがいません。</td>
                    </tr>
                 )}
              </tbody>
            </table>
          </div>
        </div>
      </section>

      {showGrossProfit && (
      <section aria-labelledby="all-users-gross-profit-title">
        <h2
          id="all-users-gross-profit-title"
          className="section-title collapsible-header"
          onClick={() => toggleSection('allUsersGrossProfit')}
          onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('allUsersGrossProfit'); } }}
          role="button"
          tabIndex={0}
          aria-expanded={visibility.grossProfit}
          aria-controls="all-users-gross-profit-content"
        >
          <span>想定粗利（パイプライン合計）</span>
          <span className={`toggle-icon ${visibility.grossProfit ? 'open' : ''}`}>▼</span>
        </h2>
        <div id="all-users-gross-profit-content" className={`collapsible-content ${visibility.grossProfit ? 'open' : ''}`}>
          <GrossProfitSummary candidates={candidatesAcrossUsers} allMedia={allMedia} />
        </div>
      </section>
      )}

      <section aria-labelledby="all-users-weekly-summary-title">
        <h2
          id="all-users-weekly-summary-title"
          className="section-title collapsible-header"
          onClick={() => toggleSection('allUsersWeeklySummary')}
          onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('allUsersWeeklySummary'); } }}
          role="button"
          tabIndex={0}
          aria-expanded={visibility.weeklySummary}
          aria-controls="all-users-weekly-summary-content"
        >
          <span>週間サマリー（合計）</span>
          <span className={`toggle-icon ${visibility.weeklySummary ? 'open' : ''}`}>▼</span>
        </h2>
        <div id="all-users-weekly-summary-content" className={`collapsible-content ${visibility.weeklySummary ? 'open' : ''}`}>
          <WeeklySummary
            weekStartDate={weekStartDate}
            data={aggregateWeeklyData}
            weeklyKpiTargets={aggregateWeeklyKpiTargets}
            onPrevWeek={onPrevWeek}
            onNextWeek={onNextWeek}
          />
        </div>
      </section>

      <section aria-labelledby="all-users-member-weekly-summary-title">
        <h2
          id="all-users-member-weekly-summary-title"
          className="section-title collapsible-header"
          onClick={() => toggleSection('allUsersMemberWeeklySummary')}
          onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('allUsersMemberWeeklySummary'); } }}
          role="button"
          tabIndex={0}
          aria-expanded={visibility.memberWeeklySummary}
          aria-controls="all-users-member-weekly-summary-content"
        >
          <span>メンバー別 週間サマリー</span>
          <span className={`toggle-icon ${visibility.memberWeeklySummary ? 'open' : ''}`}>▼</span>
        </h2>
        <div id="all-users-member-weekly-summary-content" className={`collapsible-content ${visibility.memberWeeklySummary ? 'open' : ''}`}>
          {perMemberWeeklyData.length === 0 && <p className="no-data-message">表示するユーザーがいません。</p>}
          {perMemberWeeklyData.map(({ user, displayName, data, weeklyKpiTargets }) => (
            <div key={user} className="member-weekly-summary-item weekly-summary-container">
              <h3 className="sub-section-title">{displayName}</h3>
              <WeeklySummaryTable data={data} weeklyKpiTargets={weeklyKpiTargets} />
            </div>
          ))}
        </div>
      </section>

      <section aria-labelledby="all-users-monthly-trend-title">
        <h2
          id="all-users-monthly-trend-title"
          className="section-title collapsible-header"
          onClick={() => toggleSection('allUsersMonthlyTrend')}
          onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('allUsersMonthlyTrend'); } }}
          role="button"
          tabIndex={0}
          aria-expanded={visibility.monthlyTrend}
          aria-controls="all-users-monthly-trend-content"
        >
          <span>月別パフォーマンストレンド</span>
          <span className={`toggle-icon ${visibility.monthlyTrend ? 'open' : ''}`}>▼</span>
        </h2>
        <div id="all-users-monthly-trend-content" className={`collapsible-content ${visibility.monthlyTrend ? 'open' : ''}`}>
          <MonthlyTrendChart perUserEntries={perUserTrendEntries} allMedia={allMedia} />
        </div>
      </section>

      <FunnelAnalysisSection
        users={users}
        allUsersData={allUsersData}
        allMedia={allMedia}
        periodOverride={periodOverride}
        periodLabel={periodLabel}
        perUserProgressStats={perUserProgressStats}
        grossProfitStageTotals={grossProfitStageTotals}
      />

      {dayOfWeekReplyRateData && (
        <section aria-labelledby="all-users-dow-title">
          <h2 
            id="all-users-dow-title" 
            className="section-title collapsible-header"
            onClick={() => toggleSection('allUsersDayOfWeekRate')}
            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('allUsersDayOfWeekRate'); } }}
            role="button"
            tabIndex={0}
            aria-expanded={visibility.dowRate}
            aria-controls="all-users-dow-content"
          >
            <span>全ユーザー 曜日別返信率（{periodLabel}）</span>
            <span className={`toggle-icon ${visibility.dowRate ? 'open' : ''}`}>▼</span>
          </h2>
          <div id="all-users-dow-content" className={`collapsible-content ${visibility.dowRate ? 'open' : ''}`}>
            <DayOfWeekReplyRateChart data={dayOfWeekReplyRateData} />
          </div>
        </section>
      )}
    </>
  );
};

interface CustomPeriodReportProps {
  entries: KpiEntry[];
  allMedia: MediaEntry[];
}

interface ReportData {
  generalTotals: KpiTotals;
  mediaStats: Array<{
    source: string;
    scoutsSent: number;
    scoutReplies: number;
    effectiveReplies: number;
    documentsCollected: number;
    effectiveDocumentsCollected: number;
    initialInterviews: number;
    effectiveInitialInterviews: number;
  }>;
  totalScoutsSent: number;
  totalScoutReplies: number;
  totalEffectiveReplies: number;
  totalDocumentsCollected: number;
  totalEffectiveDocumentsCollected: number;
  totalInitialInterviews: number;
  totalEffectiveInitialInterviews: number;
}


const CustomPeriodReport: React.FC<CustomPeriodReportProps> = ({ entries, allMedia }) => {
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [reportData, setReportData] = useState<ReportData | null>(null);

  const handleGenerateReport = () => {
    if (!startDate || !endDate || new Date(startDate) > new Date(endDate)) {
      alert('有効な期間を選択してください。');
      setReportData(null);
      return;
    }
    
    const start = new Date(startDate + 'T00:00:00');
    const end = new Date(endDate + 'T00:00:00');

    const filteredEntries = entries.filter(entry => {
      const entryDate = new Date(entry.date + 'T00:00:00');
      return entryDate >= start && entryDate <= end;
    });

    if (filteredEntries.length === 0) {
        alert('指定された期間に実績データがありません。');
        setReportData(null);
        return;
    }

    const allKeys = buildAllKpiKeys(allMedia);
    const totals = allKeys.reduce((acc, key) => {
      acc[key] = 0;
      return acc;
    }, {} as KpiTotals);

    filteredEntries.forEach(entry => {
      allKeys.forEach(key => {
        totals[key] += entry.values[key] || 0;
      });
    });

    const mediaStats = allMedia.map(source => {
      const sourceKey = source.id;
      const scoutsKey = `${sourceKey}_scoutsSent` as KpiKey;
      const repliesKey = `${sourceKey}_scoutReplies` as KpiKey;
      const effectiveRepliesKey = `${sourceKey}_effectiveReplies` as KpiKey;
      const documentsCollectedKey = `${sourceKey}_documentsCollected` as KpiKey;
      const effectiveDocumentsCollectedKey = `${sourceKey}_effectiveDocumentsCollected` as KpiKey;
      const interviewsKey = `${sourceKey}_initialInterviews` as KpiKey;
      const effectiveInterviewsKey = `${sourceKey}_effectiveInitialInterviews` as KpiKey;
      return {
        source: source.name,
        scoutsSent: totals[scoutsKey],
        scoutReplies: totals[repliesKey],
        effectiveReplies: totals[effectiveRepliesKey],
        documentsCollected: totals[documentsCollectedKey],
        effectiveDocumentsCollected: totals[effectiveDocumentsCollectedKey],
        initialInterviews: totals[interviewsKey],
        effectiveInitialInterviews: totals[effectiveInterviewsKey],
      };
    });

    const totalScoutsSent = mediaStats.reduce((sum, stat) => sum + stat.scoutsSent, 0);
    const totalScoutReplies = mediaStats.reduce((sum, stat) => sum + stat.scoutReplies, 0);
    const totalEffectiveReplies = mediaStats.reduce((sum, stat) => sum + stat.effectiveReplies, 0);
    const totalDocumentsCollected = mediaStats.reduce((sum, stat) => sum + stat.documentsCollected, 0);
    const totalEffectiveDocumentsCollected = mediaStats.reduce((sum, stat) => sum + stat.effectiveDocumentsCollected, 0);
    const totalInitialInterviews = mediaStats.reduce((sum, stat) => sum + stat.initialInterviews, 0);
    const totalEffectiveInitialInterviews = mediaStats.reduce((sum, stat) => sum + stat.effectiveInitialInterviews, 0);
    
    setReportData({
      generalTotals: totals,
      mediaStats,
      totalScoutsSent,
      totalScoutReplies,
      totalEffectiveReplies,
      totalDocumentsCollected,
      totalEffectiveDocumentsCollected,
      totalInitialInterviews,
      totalEffectiveInitialInterviews
    });
  };

  const handleExportCSV = () => {
    if (!reportData) return;

    // BOM for Excel compatibility with UTF-8
    let csvContent = '\uFEFF';

    // Header for report period
    csvContent += `実績期間: ${startDate.replace(/-/g, '/')} ~ ${endDate.replace(/-/g, '/')}\n\n`;

    // General KPIs section
    csvContent += '全体実績\n';
    csvContent += '指標,数値\n';

    csvContent += `"書類回収数",${reportData.totalDocumentsCollected}\n`;
    csvContent += `"有効書類回収数",${reportData.totalEffectiveDocumentsCollected}\n`;
    csvContent += `"初回面談数",${reportData.totalInitialInterviews}\n`;
    csvContent += `"初回有効面談数",${reportData.totalEffectiveInitialInterviews}\n`;

    (Object.keys(GENERAL_KPIS) as Array<keyof typeof GENERAL_KPIS>).forEach(key => {
        const label = GENERAL_KPIS[key].label;
        const value = reportData.generalTotals[key] || 0;
        csvContent += `"${label}",${value}\n`;
    });
    
    const rate = reportData.totalInitialInterviews > 0 ? (reportData.totalEffectiveInitialInterviews / reportData.totalInitialInterviews) * 100 : 0;
    csvContent += `"有効面談率","${rate.toFixed(1)}%"\n`;

    csvContent += '\n'; // Spacer

    // Media KPIs section
    csvContent += '媒体別実績\n';
    csvContent += '媒体,スカウト数,返信数,有効返信数,書類回収数,有効書類回収数,初回面談数,初回有効面談数,返信率 (%),有効返信率 (%)\n';
    
    reportData.mediaStats.forEach(stat => {
        const replyRate = stat.scoutsSent > 0 ? (stat.scoutReplies / stat.scoutsSent) * 100 : 0;
        const effectiveReplyRate = stat.scoutReplies > 0 ? (stat.effectiveReplies / stat.scoutReplies) * 100 : 0;
        const row = [
            `"${stat.source}"`,
            stat.scoutsSent,
            stat.scoutReplies,
            stat.effectiveReplies,
            stat.documentsCollected,
            stat.effectiveDocumentsCollected,
            stat.initialInterviews,
            stat.effectiveInitialInterviews,
            replyRate.toFixed(1),
            effectiveReplyRate.toFixed(1),
        ].join(',');
        csvContent += `${row}\n`;
    });

    // Totals row for Media
    const totalReplyRate = reportData.totalScoutsSent > 0 ? (reportData.totalScoutReplies / reportData.totalScoutsSent) * 100 : 0;
    const totalEffectiveReplyRate = reportData.totalScoutReplies > 0 ? (reportData.totalEffectiveReplies / reportData.totalScoutReplies) * 100 : 0;
    const totalsRow = [
        '"合計"',
        reportData.totalScoutsSent,
        reportData.totalScoutReplies,
        reportData.totalEffectiveReplies,
        reportData.totalDocumentsCollected,
        reportData.totalEffectiveDocumentsCollected,
        reportData.totalInitialInterviews,
        reportData.totalEffectiveInitialInterviews,
        totalReplyRate.toFixed(1),
        totalEffectiveReplyRate.toFixed(1),
    ].join(',');
    csvContent += `${totalsRow}\n`;

    // Create blob and trigger download
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', `kpi_report_${startDate}_to_${endDate}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };
  
  const totalReplyRate = reportData ? (reportData.totalScoutsSent > 0 ? (reportData.totalScoutReplies / reportData.totalScoutsSent) * 100 : 0) : 0;
  const totalEffectiveReplyRate = reportData ? (reportData.totalScoutReplies > 0 ? (reportData.totalEffectiveReplies / reportData.totalScoutReplies) * 100 : 0) : 0;

  return (
    <div className="custom-report-container">
      <div className="report-controls">
        <div className="form-group">
          <label htmlFor="start-date">開始日</label>
          <input type="date" id="start-date" value={startDate} onChange={e => setStartDate(e.target.value)} aria-label="実績表示の開始日"/>
        </div>
        <div className="form-group">
          <label htmlFor="end-date">終了日</label>
          <input type="date" id="end-date" value={endDate} onChange={e => setEndDate(e.target.value)} aria-label="実績表示の終了日"/>
        </div>
        <button onClick={handleGenerateReport} className="submit-button" disabled={!startDate || !endDate}>表示</button>
        <button onClick={handleExportCSV} className="export-button" disabled={!reportData}>スプレッドシートで出力</button>
      </div>
      
      {reportData && (
        <div className="report-results">
          <h4 className="report-period-header">
            実績期間: {startDate.replace(/-/g, '/')} ~ {endDate.replace(/-/g, '/')}
          </h4>

          <h5 className="report-section-title">全体実績</h5>
          <div className="report-general-kpis kpi-dashboard">
             <div className="report-kpi-card">
                <span className="report-kpi-label">書類回収数</span>
                <span className="report-kpi-value">{reportData.totalDocumentsCollected}</span>
              </div>
               <div className="report-kpi-card">
                <span className="report-kpi-label">有効書類回収数</span>
                <span className="report-kpi-value">{reportData.totalEffectiveDocumentsCollected}</span>
              </div>
             <div className="report-kpi-card">
                <span className="report-kpi-label">初回面談数</span>
                <span className="report-kpi-value">{reportData.totalInitialInterviews}</span>
              </div>
               <div className="report-kpi-card">
                <span className="report-kpi-label">初回有効面談数</span>
                <span className="report-kpi-value">{reportData.totalEffectiveInitialInterviews}</span>
              </div>
            {(Object.keys(GENERAL_KPIS) as Array<keyof typeof GENERAL_KPIS>).map(key => (
              <div key={key} className="report-kpi-card">
                <span className="report-kpi-label">{GENERAL_KPIS[key].label}</span>
                <span className="report-kpi-value">{reportData.generalTotals[key] || 0}</span>
              </div>
            ))}
            <div className="report-kpi-card">
                <span className="report-kpi-label">有効面談率</span>
                <span className="report-kpi-value">
                {(() => {
                    const rate = reportData.totalInitialInterviews > 0 ? (reportData.totalEffectiveInitialInterviews / reportData.totalInitialInterviews) * 100 : 0;
                    return `${rate.toFixed(1)}%`;
                })()}
                </span>
            </div>
          </div>

          <h5 className="report-section-title">媒体別実績</h5>
          <div className="all-users-table-container">
            <table className="weekly-summary-table">
              <thead>
                <tr>
                  <th>媒体</th>
                  <th>スカウト数</th>
                  <th>返信数</th>
                  <th>有効返信数</th>
                  <th>書類回収数</th>
                  <th>有効書類回収数</th>
                  <th>初回面談数</th>
                  <th>初回有効面談数</th>
                  <th>返信率</th>
                  <th>有効返信率</th>
                </tr>
              </thead>
              <tbody>
                {reportData.mediaStats.map(stat => {
                  const replyRate = stat.scoutsSent > 0 ? (stat.scoutReplies / stat.scoutsSent) * 100 : 0;
                  const effectiveReplyRate = stat.scoutReplies > 0 ? (stat.effectiveReplies / stat.scoutReplies) * 100 : 0;
                  return (
                    <tr key={stat.source}>
                      <td>{stat.source}</td>
                      <td>{stat.scoutsSent}</td>
                      <td>{stat.scoutReplies}</td>
                      <td>{stat.effectiveReplies}</td>
                      <td>{stat.documentsCollected}</td>
                      <td>{stat.effectiveDocumentsCollected}</td>
                      <td>{stat.initialInterviews}</td>
                      <td>{stat.effectiveInitialInterviews}</td>
                      <td>{replyRate.toFixed(1)}%</td>
                      <td>{effectiveReplyRate.toFixed(1)}%</td>
                    </tr>
                  )
                })}
              </tbody>
              <tfoot>
                <tr>
                  <td>合計</td>
                  <td>{reportData.totalScoutsSent}</td>
                  <td>{reportData.totalScoutReplies}</td>
                  <td>{reportData.totalEffectiveReplies}</td>
                  <td>{reportData.totalDocumentsCollected}</td>
                  <td>{reportData.totalEffectiveDocumentsCollected}</td>
                  <td>{reportData.totalInitialInterviews}</td>
                  <td>{reportData.totalEffectiveInitialInterviews}</td>
                  <td>{totalReplyRate.toFixed(1)}%</td>
                  <td>{totalEffectiveReplyRate.toFixed(1)}%</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};


type SectionVisibilityKeys = 
  | 'monthlyProgress' | 'monthlyPerformance' | 'monthOverMonthPerformance'
  | 'weeklySummary' | 'dayOfWeekRate' | 'mediaProgress' 
  | 'monthlyTargetSettings' | 'weeklyTargetSettings' | 'dailyTargetSettings' | 'calendar' | 'history'
  | 'dailyProgress' | 'customPeriodReport'
  | 'allUsersProgress' | 'allUsersDayOfWeekRate' | 'allUsersWeeklySummary' | 'allUsersMemberWeeklySummary' | 'allUsersGrossProfit'
  | 'allUsersMonthlyTrend';


const App: React.FC = () => {
  // Authentication state (Google Sign-In, bloom-firm.com domain only)
  const [currentIdentity, setCurrentIdentity] = useState<GoogleIdentity | null>(null);
  const [isSigningIn, setIsSigningIn] = useState(false);
  const [authError, setAuthError] = useState('');

  // Multi-user state (the signed-in Google account is always the current user)
  const [users, setUsers] = useState<string[]>([]);
  const [currentUser, setCurrentUser] = useState<string | null>(null);
  const [isInitialized, setIsInitialized] = useState(false);

  // Google Drive storage state
  const [driveFileId, setDriveFileId] = useState<string | null>(null);
  const [isLoadingUserData, setIsLoadingUserData] = useState(false);
  const [legacyMigrationChoices, setLegacyMigrationChoices] = useState<string[] | null>(null);
  const [isLoadingAllUsers, setIsLoadingAllUsers] = useState(false);

  // View state
  const [view, setView] = useState<'personal_kpi' | 'all_users_kpi' | 'team_kpi' | 'pipeline'>('personal_kpi');
  const [allUsersData, setAllUsersData] = useState<Record<string, UserData>>({});

  // BCA事業部 header switcher — 'BCA' shows F+ and AC combined (not a real assignment of its
  // own); filters which members' data 全ユーザー/チーム別/パイプライン(全ユーザー) show.
  const [selectedDivision, setSelectedDivision] = useState<'BCA' | Department>('BCA');

  // Teams state
  const [teams, setTeams] = useState<Team[]>([]);
  const [teamsDriveFileId, setTeamsDriveFileId] = useState<string | null>(null);
  const [teamsOwnerEmail, setTeamsOwnerEmail] = useState<string | null>(null);
  const [teamsAuthorizedEditors, setTeamsAuthorizedEditors] = useState<string[]>([]);
  const [memberDepartments, setMemberDepartments] = useState<Record<string, Department>>({});
  const [isTeamsModalOpen, setIsTeamsModalOpen] = useState(false);
  const [selectedTeamId, setSelectedTeamId] = useState<string | null>(null);
  // Empty = no filter (show everyone) on the 全ユーザー tab; otherwise an ad-hoc selection of
  // specific users to compare, independent of the formal Team groupings.
  const [comparisonUserEmails, setComparisonUserEmails] = useState<string[]>([]);
  const [customExportStartDate, setCustomExportStartDate] = useState('');
  const [customExportEndDate, setCustomExportEndDate] = useState('');
  // Reuses the same start/end fields the custom-period CSV export already had: once explicitly
  // enabled (below), the 全ユーザー/チーム別 progress dashboard switches from "this month" to
  // this range too, instead of needing a second, separate period picker. Requires an explicit
  // enable step rather than auto-activating whenever both dates happen to be filled in, so
  // typing dates doesn't change the dashboard out from under the user before they're ready.
  const [isPeriodFilterEnabled, setIsPeriodFilterEnabled] = useState(false);
  const dashboardPeriodOverride = useMemo(() => {
    if (!isPeriodFilterEnabled || !customExportStartDate || !customExportEndDate) return null;
    return { start: new Date(customExportStartDate + 'T00:00:00'), end: new Date(customExportEndDate + 'T23:59:59') };
  }, [isPeriodFilterEnabled, customExportStartDate, customExportEndDate]);
  const handleTogglePeriodFilter = () => {
    if (dashboardPeriodOverride) {
      setIsPeriodFilterEnabled(false);
      // Returning to "this month" should also snap 週間サマリー back to the current real week,
      // rather than leaving it wherever 前月/次月 last moved it.
      setViewWeekStartDate(getStartOfWeek(new Date()));
      return;
    }
    if (!customExportStartDate || !customExportEndDate) {
      alert('開始日と終了日を指定してください。');
      return;
    }
    setIsPeriodFilterEnabled(true);
  };

  const formatDateInputValue = (d: Date): string =>
    `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

  // Steps the 全ユーザー/チーム別 dashboards a whole calendar month forward/back with one click,
  // instead of needing to type both dates by hand every time. Shifts relative to whichever
  // month is currently selected (so repeated clicks keep walking further back/forward), or from
  // today's month if the period filter isn't enabled yet. Also moves 週間サマリー (which has its
  // own independent 前週/次週 navigation) to the week containing this month's 1st, so it doesn't
  // silently stay stuck on a week from a different month while everything else moved — the
  // weekly nav remains usable afterwards for fine-tuning within/around that month.
  const handleShiftDashboardMonth = (offset: number) => {
    const reference = customExportStartDate ? new Date(customExportStartDate + 'T00:00:00') : new Date();
    const targetYear = reference.getFullYear();
    const targetMonth = reference.getMonth() + offset;
    const monthStart = new Date(targetYear, targetMonth, 1);
    const monthEnd = new Date(targetYear, targetMonth + 1, 0);
    setCustomExportStartDate(formatDateInputValue(monthStart));
    setCustomExportEndDate(formatDateInputValue(monthEnd));
    setIsPeriodFilterEnabled(true);
    setViewWeekStartDate(getStartOfWeek(monthStart));
  };

  // Independent period control for 個人実績's own 曜日別累積返信率 chart — separate from the
  // 全ユーザー/チーム別 dashboards' period fields above, since a period picked for one's own
  // chart has no reason to be coupled to whatever is selected on a different tab. Unfiltered
  // default is all-time (unlike the other dashboards' "this month" default), matching this
  // chart's existing behavior before this control existed.
  const [personalDowStartDate, setPersonalDowStartDate] = useState('');
  const [personalDowEndDate, setPersonalDowEndDate] = useState('');
  const [isPersonalDowPeriodEnabled, setIsPersonalDowPeriodEnabled] = useState(false);
  const personalDowPeriodOverride = useMemo(() => {
    if (!isPersonalDowPeriodEnabled || !personalDowStartDate || !personalDowEndDate) return null;
    return { start: new Date(personalDowStartDate + 'T00:00:00'), end: new Date(personalDowEndDate + 'T23:59:59') };
  }, [isPersonalDowPeriodEnabled, personalDowStartDate, personalDowEndDate]);
  const handleTogglePersonalDowPeriod = () => {
    if (personalDowPeriodOverride) {
      setIsPersonalDowPeriodEnabled(false);
      return;
    }
    if (!personalDowStartDate || !personalDowEndDate) {
      alert('開始日と終了日を指定してください。');
      return;
    }
    setIsPersonalDowPeriodEnabled(true);
  };
  const [pipelineScope, setPipelineScope] = useState<'personal' | 'all_users' | 'team' | 'user'>('personal');
  const [pipelineSelectedTeamId, setPipelineSelectedTeamId] = useState<string | null>(null);
  const [pipelineSelectedUserEmail, setPipelineSelectedUserEmail] = useState<string | null>(null);
  const [isEditingDisplayName, setIsEditingDisplayName] = useState(false);
  const [editedDisplayName, setEditedDisplayName] = useState('');

  // Media (scouting source) state — shared across all users, Drive-backed like Teams
  const [allMedia, setAllMedia] = useState<MediaEntry[]>(() => readMediaConfigCache<MediaConfig>()?.media || []);
  const [mediaDriveFileId, setMediaDriveFileId] = useState<string | null>(null);
  const [mediaOwnerEmail, setMediaOwnerEmail] = useState<string | null>(null);
  const [isLoadingMediaConfig, setIsLoadingMediaConfig] = useState(allMedia.length === 0);
  const [isMediaModalOpen, setIsMediaModalOpen] = useState(false);

  // --- Consolidated user data state ---
  const [currentUserData, setCurrentUserData] = useState<UserData | null>(null);
  const [hasSyncError, setHasSyncError] = useState(false);
  const [lastSyncedAt, setLastSyncedAt] = useState<number | null>(null);
  const [isForcingSync, setIsForcingSync] = useState(false);
  const [isBulkGmailModalOpen, setIsBulkGmailModalOpen] = useState(false);
  const [isMediaCsvModalOpen, setIsMediaCsvModalOpen] = useState(false);

  // UI state
  const [viewDate, setViewDate] = useState(new Date());
  const [selectedDate, setSelectedDate] = useState<string | null>(null);
  const [viewWeekStartDate, setViewWeekStartDate] = useState(getStartOfWeek(new Date()));
  const [sectionVisibility, setSectionVisibility] = useState({
    monthlyProgress: true,
    monthlyPerformance: true,
    monthOverMonthPerformance: true,
    weeklySummary: true,
    dayOfWeekRate: true,
    mediaProgress: true,
    monthlyTargetSettings: false,
    weeklyTargetSettings: false,
    dailyTargetSettings: false,
    calendar: true,
    dailyProgress: true,
    history: false,
    customPeriodReport: false,
    allUsersProgress: false,
    allUsersDayOfWeekRate: false,
    allUsersWeeklySummary: false,
    allUsersMemberWeeklySummary: false,
    allUsersGrossProfit: false,
    allUsersMonthlyTrend: false,
  });

  const toggleSection = (sectionKey: SectionVisibilityKeys) => {
    setSectionVisibility(prev => ({
        ...prev,
        [sectionKey]: !prev[sectionKey]
    }));
  };
  
  // Restore the Google session, if the access token from earlier in this browser tab's
  // session is still valid. We deliberately do NOT attempt a network re-auth here: any
  // token request not triggered by a real click gets blocked by the browser's popup
  // blocker, so it would only add a delay before showing the login screen anyway. Once
  // expired/cleared, signIn() (called from the login button) tries a silent-if-possible
  // flow first using the last-known email, so returning users still get a near-instant
  // re-login with just one click.
  useEffect(() => {
    const session = getCurrentSession();
    if (session) {
      setCurrentIdentity(session.identity);
    }
    setIsInitialized(true);
  }, []);

  // The signed-in Google account's email is always the current user; register it
  // in the local roster automatically (no manual "create user" step anymore).
  useEffect(() => {
    if (!currentIdentity) {
      setCurrentUser(null);
      return;
    }
    const email = currentIdentity.email;
    setCurrentUser(email);
    setUsers(prev => (prev.includes(email) ? prev : [...prev, email]));
  }, [currentIdentity]);


  // Load the signed-in user's data. Drive is the source of truth, but if we have a local
  // cache from a previous session we show it immediately (no loading spinner) and quietly
  // upgrade to the fresh Drive copy once it arrives — this removes the "wait for Drive"
  // delay on every normal app open, which is by far the most common case.
  useEffect(() => {
    if (!currentIdentity) {
      setCurrentUserData(null);
      setDriveFileId(null);
      return;
    }
    let cancelled = false;
    const email = currentIdentity.email;

    const normalize = (d: Partial<UserData>): UserData => ({
      entries: d.entries || [],
      // ownerEmail/ownerLabel are presentation-only tags applied when flattening candidates for
      // an aggregate (全ユーザー/チーム/ユーザー別) pipeline view — never meant to be persisted.
      // Editing one's own candidate while viewing such a scope used to save the tagged copy
      // straight through, baking a stray "登録者: (自分)" label into that one candidate forever
      // while untouched candidates stayed clean, which read as "some cards have the label, some
      // don't". Stripped here (self-heals anything already saved with the leak) and again at
      // the actual save point below (handleSaveCandidate) to stop it recurring.
      candidates: (d.candidates || []).map(({ ownerEmail, ownerLabel, ...c }) => c),
      kpiTargets: { ...defaultKpiTargets, ...(d.kpiTargets || {}) },
      weeklyKpiTargets: { ...defaultKpiTargets, ...(d.weeklyKpiTargets || {}) },
      dailyKpiTargets: { ...defaultKpiTargets, ...(d.dailyKpiTargets || {}) },
      displayName: d.displayName || currentIdentity.name,
    });

    const cached = readLocalCache<UserData>(email);
    if (cached) {
      setCurrentUserData(normalize(cached));
      setIsLoadingUserData(false);
    } else {
      setIsLoadingUserData(true);
    }

    (async () => {
      const result = await loadOwnData<UserData>(email);
      if (cancelled) return;

      if (result.data) {
        setDriveFileId(result.driveFileId);
        const fresh = normalize(result.data);
        setCurrentUserData(prev => {
          // If the user already started editing while this fetch was in flight, keep their
          // edits instead of clobbering them with the (now slightly stale) Drive snapshot.
          if (cached && prev && JSON.stringify(prev) !== JSON.stringify(normalize(cached))) {
            return prev;
          }
          // A prior save never made it to Drive (e.g. the session expired/was revoked first)
          // — the local cache is newer than what Drive just returned, so keep showing it rather
          // than flashing back to the stale Drive snapshot while retryPendingSyncIfNeeded below
          // pushes the correct data.
          if (cached && hasPendingSync(email)) {
            return normalize(cached);
          }
          return fresh;
        });
        // A prior save may have failed (e.g. an expired session right as it fired) and never
        // reached Drive — now that we have a confirmed-working session, retry it.
        retryPendingSyncIfNeeded(email, result.driveFileId, setDriveFileId);
      } else if (!cached) {
        // Brand-new signed-in user: offer to claim any pre-Google-login local data.
        const legacy = readLegacyAppData();
        const legacyNames = (legacy?.users || []).filter(name => legacy?.userData?.[name]);
        if (legacyNames.length > 0) {
          setLegacyMigrationChoices(legacyNames);
        }
        setCurrentUserData(normalize({}));
        setDriveFileId(null);
      }
      setIsLoadingUserData(false);
    })();
    return () => { cancelled = true; };
  }, [currentIdentity]);

  const handleClaimLegacyData = (legacyName: string | null) => {
    setLegacyMigrationChoices(null);
    if (!legacyName) return;
    const legacy = readLegacyAppData();
    const legacyUserData = legacy?.userData?.[legacyName];
    if (!legacyUserData) return;
    setCurrentUserData({
      entries: legacyUserData.entries || [],
      candidates: legacyUserData.candidates || [],
      kpiTargets: { ...defaultKpiTargets, ...(legacyUserData.kpiTargets || {}) },
      weeklyKpiTargets: { ...defaultKpiTargets, ...(legacyUserData.weeklyKpiTargets || {}) },
      dailyKpiTargets: { ...defaultKpiTargets, ...(legacyUserData.dailyKpiTargets || {}) },
      displayName: legacyName,
    });
  };

  const handleSaveDisplayName = (name: string) => {
    setCurrentUserData(prev => (prev ? { ...prev, displayName: name } : prev));
  };


  // Load every teammate's Drive-shared data (domain-wide, cross-device). This is fetched once
  // per sign-in session the first time either view is opened, not on every tab switch — flipping
  // back and forth between "全ユーザー"/"チーム別" and other tabs reuses what's already loaded.
  // Use the "更新" button in those views to force a fresh pull from Drive.
  const hasFetchedAllUsersRef = useRef(false);

  const fetchAllUsersData = useCallback(async () => {
    setIsLoadingAllUsers(true);
    try {
      const teammates = await loadAllTeammatesData<UserData>();
      const merged: Record<string, UserData> = {};
      teammates.forEach(({ email, data }) => {
        merged[email] = {
          entries: data.entries || [],
          candidates: data.candidates || [],
          kpiTargets: { ...defaultKpiTargets, ...(data.kpiTargets || {}) },
          weeklyKpiTargets: { ...defaultKpiTargets, ...(data.weeklyKpiTargets || {}) },
          dailyKpiTargets: { ...defaultKpiTargets, ...(data.dailyKpiTargets || {}) },
          displayName: data.displayName,
        };
      });
      setAllUsersData(merged);
      setUsers(Object.keys(merged));
    } catch (error) {
      console.error("Failed to load teammates' data from Drive", error);
    } finally {
      setIsLoadingAllUsers(false);
    }
  }, []);

  useEffect(() => {
    const needsAggregateData =
      view === 'all_users_kpi' || view === 'team_kpi' || (view === 'pipeline' && pipelineScope !== 'personal') || isTeamsModalOpen;
    if (!needsAggregateData || !isInitialized || !currentIdentity) return;
    if (hasFetchedAllUsersRef.current) return;
    hasFetchedAllUsersRef.current = true;
    fetchAllUsersData();
  }, [view, isInitialized, currentIdentity, fetchAllUsersData, pipelineScope, isTeamsModalOpen]);

  // Loads unconditionally after sign-in (like the media config below), not gated by
  // view/modal — memberDepartments now feeds the header's BCA/F+/AC division switcher, which
  // is visible everywhere, so this can no longer wait until the user happens to open Teams-
  // related UI. Firing once per sign-in (deps are just identity/init, not view/modal state)
  // also sidesteps the earlier redundant-refetch race this effect used to have to guard against
  // with a memoized "needsTeams" boolean.
  useEffect(() => {
    if (!currentIdentity || !isInitialized) return;
    let cancelled = false;
    (async () => {
      try {
        const result = await loadTeamsConfig<TeamsConfig>();
        if (cancelled) return;
        setTeams(result.data?.teams || []);
        setTeamsDriveFileId(result.driveFileId);
        setTeamsOwnerEmail(result.ownerEmail);
        setTeamsAuthorizedEditors(result.data?.authorizedEditorEmails || []);
        setMemberDepartments(result.data?.memberDepartments || {});
      } catch (error) {
        console.error('Failed to load teams config from Drive', error);
      }
    })();
    return () => { cancelled = true; };
  }, [currentIdentity, isInitialized]);

  // Only TEAMS_ADMIN_EMAIL and whoever they've explicitly granted access to (via TeamsModal's
  // permission section) can create/edit teams — file ownership no longer determines this, since
  // the full `drive` scope now lets any of these accounts actually write the shared file
  // regardless of who originally created it.
  const isTeamsAdmin = currentIdentity?.email === TEAMS_ADMIN_EMAIL;
  const isTeamsEditable = isTeamsAdmin || (!!currentIdentity && teamsAuthorizedEditors.includes(currentIdentity.email));

  // Load the shared media list once per sign-in. Unlike Teams, this is required for the KPI
  // forms to render at all, so it loads unconditionally after login (not gated by view/modal).
  // Discovery is scoped to MEDIA_ADMIN_EMAIL's file specifically (see findMediaConfigFile), and
  // only that account is allowed to auto-create it if missing — otherwise a non-admin whose
  // discovery query ever failed would silently create their own stray duplicate that no one
  // else can see, which is exactly what happened before this scoping existed.
  useEffect(() => {
    if (!currentIdentity || !isInitialized) return;
    let cancelled = false;
    (async () => {
      try {
        const result = await loadMediaConfig<MediaConfig>(MEDIA_ADMIN_EMAIL);
        if (cancelled) return;
        if (result.data) {
          setAllMedia(result.data.media || []);
          setMediaDriveFileId(result.driveFileId);
          setMediaOwnerEmail(result.ownerEmail);
        } else if (currentIdentity.email === MEDIA_ADMIN_EMAIL) {
          const seeded: MediaConfig = { schemaVersion: 1, media: SEED_MEDIA };
          const newFileId = await saveMediaConfig(null, seeded, currentIdentity.email);
          if (cancelled) return;
          setAllMedia(SEED_MEDIA);
          setMediaDriveFileId(newFileId);
          setMediaOwnerEmail(currentIdentity.email);
        } else {
          // The admin hasn't created it yet — fall back to the seed list in memory only
          // (not persisted) so KPI forms still render; it'll pick up the real one once the
          // admin logs in and creates it.
          setAllMedia(SEED_MEDIA);
        }
      } catch (error) {
        console.error('Failed to load media config from Drive', error);
      } finally {
        if (!cancelled) setIsLoadingMediaConfig(false);
      }
    })();
    return () => { cancelled = true; };
  }, [currentIdentity, isInitialized]);

  const refreshMediaConfig = useCallback(async () => {
    if (!currentIdentity) return;
    try {
      const result = await loadMediaConfig<MediaConfig>(MEDIA_ADMIN_EMAIL);
      if (result.data) {
        setAllMedia(result.data.media || []);
        setMediaDriveFileId(result.driveFileId);
        setMediaOwnerEmail(result.ownerEmail);
      }
    } catch (error) {
      console.error('Failed to refresh media config from Drive', error);
    }
  }, [currentIdentity]);

  // The initial load above only runs once per sign-in, so anyone with the app already open in
  // a tab would otherwise keep seeing whatever media config was current when they loaded the
  // page — re-fetch every time the 媒体管理 modal is opened so changes the admin made since
  // then (e.g. a new fee rate) show up without a full page reload.
  useEffect(() => {
    if (isMediaModalOpen) refreshMediaConfig();
  }, [isMediaModalOpen, refreshMediaConfig]);

  const isMediaEditable = currentIdentity?.email === MEDIA_ADMIN_EMAIL;
  const activeMedia = useMemo(() => allMedia.filter(m => !m.isArchived), [allMedia]);
  const defaultKpiTargets = useMemo(() => buildDefaultKpiTargets(allMedia), [allMedia]);

  const handleCustomPeriodExport = (label: string, exportUsers: string[]) => {
    if (!customExportStartDate || !customExportEndDate) {
      alert('開始日と終了日を指定してください。');
      return;
    }
    const start = new Date(customExportStartDate + 'T00:00:00');
    const end = new Date(customExportEndDate + 'T23:59:59');
    if (start > end) {
      alert('開始日は終了日より前の日付にしてください。');
      return;
    }
    const csvContent = buildTeamProgressCsvForRange(label, exportUsers, displayedAllUsersData, allMedia, start, end);
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', `progress_${label.replace(/[^\w\-ぁ-んァ-ヶ一-龠]/g, '_')}_${customExportStartDate}_to_${customExportEndDate}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Single CSV button for the 全ユーザー/チーム別 tabs: exports whatever the dashboard is
  // currently showing — the custom period (if the period filter toggle is enabled) or the
  // regular this-month/this-week data otherwise — instead of having a separate button per mode.
  const handleExportAllUsersProgress = (label: string, exportUsers: string[]) => {
    if (dashboardPeriodOverride) {
      handleCustomPeriodExport(label, exportUsers);
      return;
    }
    const csvContent = buildTeamProgressCsv(label, exportUsers, displayedAllUsersData, allMedia, viewWeekStartDate);
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    const today = new Date().toISOString().slice(0, 10);
    link.setAttribute('download', `all_users_progress_${label.replace(/[^\w\-ぁ-んァ-ヶ一-龠]/g, '_')}_${today}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleExportTeamProgress = (teamName: string, teamUsers: string[]) => {
    if (dashboardPeriodOverride) {
      handleCustomPeriodExport(teamName, teamUsers);
      return;
    }
    const csvContent = buildTeamProgressCsv(teamName, teamUsers, displayedAllUsersData, allMedia, viewWeekStartDate);
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    const today = new Date().toISOString().slice(0, 10);
    link.setAttribute('download', `team_progress_${teamName.replace(/[^\w\-ぁ-んァ-ヶ一-龠]/g, '_')}_${today}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const persistMedia = async (updatedMedia: MediaEntry[]) => {
    setAllMedia(updatedMedia);
    if (!currentIdentity) return;
    try {
      const payload: MediaConfig = { schemaVersion: 1, media: updatedMedia };
      const newFileId = await saveMediaConfig(mediaDriveFileId, payload, currentIdentity.email);
      setMediaDriveFileId(newFileId);
      if (!mediaOwnerEmail) setMediaOwnerEmail(currentIdentity.email);
    } catch (error) {
      console.error('Failed to save media config', error);
      alert('媒体設定の保存に失敗しました。編集できるのは作成者のみです。');
    }
  };

  const slugifyMediaName = (name: string): string => {
    const base = name.toLowerCase().replace(/[^a-z0-9]+/g, '') || 'media';
    let candidate = base;
    let i = 1;
    while (allMedia.some(m => m.id === candidate)) {
      candidate = `${base}${i}`;
      i += 1;
    }
    return candidate;
  };

  const handleCreateMedia = (name: string) => {
    const newMedia: MediaEntry = {
      id: slugifyMediaName(name),
      name,
      isArchived: false,
      createdAt: new Date().toISOString(),
    };
    persistMedia([...allMedia, newMedia]);
  };

  const handleRenameMedia = (id: string, name: string) => {
    persistMedia(allMedia.map(m => (m.id === id ? { ...m, name } : m)));
  };

  const handleSetMediaFeeRate = (id: string, feeRate: number | undefined) => {
    persistMedia(allMedia.map(m => (m.id === id ? { ...m, feeRate } : m)));
  };

  const handleArchiveMedia = (id: string) => {
    persistMedia(allMedia.map(m => (m.id === id ? { ...m, isArchived: true } : m)));
  };

  const handleUnarchiveMedia = (id: string) => {
    persistMedia(allMedia.map(m => (m.id === id ? { ...m, isArchived: false } : m)));
  };

  // The signed-in user's own Drive file can lag a few seconds behind (debounced sync), so
  // always show our own in-memory copy in aggregate views instead of whatever Drive last
  // returned for us — this keeps things like a just-edited display name visible immediately.
  const displayedAllUsersData = useMemo(() => {
    if (!currentIdentity || !currentUserData) return allUsersData;
    return { ...allUsersData, [currentIdentity.email]: currentUserData };
  }, [allUsersData, currentIdentity, currentUserData]);

  // Division-switcher filter for 全ユーザー/チーム別/パイプライン(全ユーザー) — 'BCA' (the
  // default) shows everyone combined; selecting F+ or AC narrows to only members explicitly
  // assigned to that department (unassigned members are hidden once a specific one is picked).
  const isEmailInSelectedDivision = useCallback((email: string) => {
    if (selectedDivision === 'BCA') return true;
    return memberDepartments[email] === selectedDivision;
  }, [selectedDivision, memberDepartments]);

  // A ref (not just React state) tracking the teams-config file id, kept in sync with
  // teamsDriveFileId below. persistTeams reads/writes this ref directly rather than the state
  // value, and chains writes through teamsWriteQueueRef — otherwise, two team mutations fired
  // in quick succession (e.g. creating two teams back-to-back, before the first save's
  // setTeamsDriveFileId has even landed) would both see driveFileId as null and each call
  // createTeamsConfigFile, silently splitting teams across two separate Drive files with only
  // one of them ever visible again afterwards.
  const teamsDriveFileIdRef = useRef<string | null>(teamsDriveFileId);
  useEffect(() => { teamsDriveFileIdRef.current = teamsDriveFileId; }, [teamsDriveFileId]);
  const teamsWriteQueueRef = useRef<Promise<void>>(Promise.resolve());

  const persistTeamsConfig = (
    updatedTeams: Team[],
    updatedAuthorizedEditors: string[],
    updatedDepartments: Record<string, Department>
  ) => {
    setTeams(updatedTeams);
    setTeamsAuthorizedEditors(updatedAuthorizedEditors);
    setMemberDepartments(updatedDepartments);
    if (!currentIdentity) return;
    const email = currentIdentity.email;
    teamsWriteQueueRef.current = teamsWriteQueueRef.current.catch(() => {}).then(async () => {
      try {
        const payload: TeamsConfig = {
          schemaVersion: 1,
          teams: updatedTeams,
          authorizedEditorEmails: updatedAuthorizedEditors,
          memberDepartments: updatedDepartments,
        };
        const newFileId = await saveTeamsConfig(teamsDriveFileIdRef.current, payload, email);
        teamsDriveFileIdRef.current = newFileId;
        setTeamsDriveFileId(newFileId);
        setTeamsOwnerEmail(prev => prev || email);
      } catch (error) {
        console.error('Failed to save teams config', error);
        alert('チーム設定の保存に失敗しました。');
      }
    });
  };

  const persistTeams = (updatedTeams: Team[]) => persistTeamsConfig(updatedTeams, teamsAuthorizedEditors, memberDepartments);

  const handleGrantTeamsEditor = (email: string) => {
    if (teamsAuthorizedEditors.includes(email)) return;
    persistTeamsConfig(teams, [...teamsAuthorizedEditors, email], memberDepartments);
  };

  const handleRevokeTeamsEditor = (email: string) => {
    persistTeamsConfig(teams, teamsAuthorizedEditors.filter(e => e !== email), memberDepartments);
  };

  // Team editors (admin + authorized) can set anyone's department from チーム管理.
  const handleSetMemberDepartment = (email: string, department: Department | null) => {
    const updated = { ...memberDepartments };
    if (department) updated[email] = department; else delete updated[email];
    persistTeamsConfig(teams, teamsAuthorizedEditors, updated);
  };

  // Anyone signed in can set their OWN department, regardless of team-editor permission — this
  // writes to the same shared config, but only ever touches this one user's own entry in it.
  const handleSetOwnDepartment = (department: Department | null) => {
    if (!currentIdentity) return;
    handleSetMemberDepartment(currentIdentity.email, department);
  };

  const handleCreateTeam = (name: string) => {
    const newTeam: Team = {
      id: `team-${Date.now()}-${Math.random().toString(16).slice(2)}`,
      name,
      memberEmails: currentIdentity ? [currentIdentity.email] : [],
      createdBy: currentIdentity?.email || '',
      createdAt: new Date().toISOString(),
    };
    persistTeams([...teams, newTeam]);
  };

  const handleRenameTeam = (teamId: string, name: string) => {
    persistTeams(teams.map(t => (t.id === teamId ? { ...t, name } : t)));
  };

  const handleDeleteTeam = (teamId: string) => {
    persistTeams(teams.filter(t => t.id !== teamId));
    if (selectedTeamId === teamId) setSelectedTeamId(null);
  };

  const handleAddTeamMember = (teamId: string, email: string) => {
    persistTeams(
      teams.map(t => (t.id === teamId && !t.memberEmails.includes(email) ? { ...t, memberEmails: [...t.memberEmails, email] } : t))
    );
  };

  const handleRemoveTeamMember = (teamId: string, email: string) => {
    persistTeams(teams.map(t => (t.id === teamId ? { ...t, memberEmails: t.memberEmails.filter(e => e !== email) } : t)));
  };

  // Sync the current user's data to Google Drive (debounced) whenever it changes.
  // Writes through to a local cache immediately so the UI never waits on the network.
  useEffect(() => {
    if (!isInitialized || !currentIdentity || !currentUserData || isLoadingUserData) return;
    saveOwnDataDebounced(currentIdentity.email, driveFileId, currentUserData, (newFileId) => {
      setDriveFileId(newFileId);
    });
  }, [currentUserData, currentIdentity, isInitialized, isLoadingUserData]);

  // Best-effort flush of any pending debounced save when the tab is backgrounded or the page
  // is about to be torn down — the 2s debounce window would otherwise silently drop the last
  // few seconds of edits if the user closes/navigates away before it fires.
  useEffect(() => {
    const flush = () => { flushPendingSave(); };
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'hidden') flush();
    };
    document.addEventListener('visibilitychange', handleVisibilityChange);
    window.addEventListener('pagehide', flush);
    return () => {
      document.removeEventListener('visibilitychange', handleVisibilityChange);
      window.removeEventListener('pagehide', flush);
    };
  }, []);

  // A save can fail mid-session too — e.g. the Google session expires or is revoked right as
  // the debounced write fires, which happens right after the user enters a KPI or pipeline
  // entry. Surface that immediately via onSyncStatusChange (fires the moment that save's
  // success/failure is known) rather than waiting for the next periodic check, which could be
  // up to a minute late and feel disconnected from the input that actually failed. The interval
  // below is just a recovery safety net (e.g. the session becomes valid again on its own,
  // without the user touching anything) — it's not the primary detection path.
  useEffect(() => {
    if (!currentIdentity) { setHasSyncError(false); setLastSyncedAt(null); return; }
    const email = currentIdentity.email;
    setHasSyncError(hasPendingSync(email));
    setLastSyncedAt(getLastSyncedAt(email));
    const unsubscribe = onSyncStatusChange((changedEmail, hasPending) => {
      if (changedEmail !== email) return;
      setHasSyncError(hasPending);
      setLastSyncedAt(getLastSyncedAt(email));
    });
    const intervalId = setInterval(() => {
      if (hasPendingSync(email)) retryPendingSyncIfNeeded(email, driveFileId, setDriveFileId);
    }, 60000);
    return () => {
      unsubscribe();
      clearInterval(intervalId);
    };
  }, [currentIdentity, driveFileId]);


    const handleGoogleSignIn = async () => {
        setIsSigningIn(true);
        setAuthError('');
        try {
            const identity = await signIn();
            setCurrentIdentity(identity);
        } catch (error) {
            setAuthError(error instanceof Error ? error.message : 'ログインに失敗しました。');
        } finally {
            setIsSigningIn(false);
        }
    };

    const handleLogout = async () => {
        // Make sure any debounced-but-not-yet-written edit actually reaches Drive before the
        // session (and its access token) is cleared — otherwise the save either never fires or
        // fires after sign-out and silently fails, and the last few seconds of input are lost.
        await flushPendingSave();
        signOut();
        setCurrentIdentity(null);
        setCurrentUser(null);
    };

    const handleForceSync = async () => {
        if (!currentIdentity || !currentUserData) return;
        setIsForcingSync(true);
        try {
            await forceSyncNow(currentIdentity.email, driveFileId, currentUserData, setDriveFileId);
        } finally {
            setIsForcingSync(false);
        }
    };

  const handleTargetChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target as { name: KpiKey; value: string };
    setCurrentUserData(prev => prev ? ({ ...prev, kpiTargets: { ...prev.kpiTargets, [name]: value === '' ? 0 : Number(value) }}) : null);
  };

  const handleWeeklyTargetChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target as { name: KpiKey; value: string };
    setCurrentUserData(prev => prev ? ({ ...prev, weeklyKpiTargets: { ...prev.weeklyKpiTargets, [name]: value === '' ? 0 : Number(value) }}) : null);
  };
  
  const handleDailyTargetChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target as { name: KpiKey; value: string };
    setCurrentUserData(prev => prev ? ({ ...prev, dailyKpiTargets: { ...prev.dailyKpiTargets, [name]: value === '' ? 0 : Number(value) }}) : null);
  };

  const persistEntry = (date: string, newValues: KpiTotals) => {
    if (!currentUserData) return;
    const otherEntries = currentUserData.entries.filter(entry => entry.date !== date);
    const newEntry: KpiEntry = {
      id: Date.now(),
      date: date,
      values: newValues,
    };
    const updatedEntries = [...otherEntries, newEntry].sort((a, b) => a.date.localeCompare(b.date));
    const updatedData = { ...currentUserData, entries: updatedEntries };
    setCurrentUserData(updatedData);
    // Sync to Drive right away on every calendar save, instead of waiting out the usual 2s
    // debounce — a save here is a single deliberate action (unlike rapid KPI-form keystrokes),
    // so there's no batching benefit to waiting, only a window where the entry looks saved but
    // isn't yet on Drive.
    if (currentIdentity) {
      forceSyncNow(currentIdentity.email, driveFileId, updatedData, setDriveFileId);
    }
  };

  const handleSaveEntry = (date: string, newValues: KpiTotals) => {
    persistEntry(date, newValues);
    setSelectedDate(null);
  };

  // 前日/次日 navigation from within the entry modal: persists the day being left (so
  // in-progress input isn't lost) without closing, then moves selectedDate to the adjacent day.
  const handleNavigateEntryDate = (currentDate: string, currentValues: KpiTotals, offsetDays: number) => {
    persistEntry(currentDate, currentValues);
    const newDate = new Date(currentDate + 'T00:00:00');
    newDate.setDate(newDate.getDate() + offsetDays);
    const newDateStr = `${newDate.getFullYear()}-${String(newDate.getMonth() + 1).padStart(2, '0')}-${String(newDate.getDate()).padStart(2, '0')}`;
    setSelectedDate(newDateStr);
  };

  // Applies a Gmail bulk-scan result (dateISO -> mediaId -> count) across many days at once:
  // merges each date's counts into that day's existing entry (creating one if it didn't exist
  // yet) without touching any other field in that entry, then syncs once for the whole batch.
  const handleApplyBulkGmailImport = (countsByDate: Record<string, Record<string, number>>) => {
    if (!currentUserData) return;
    const entriesByDateMap = new Map<string, KpiEntry>(currentUserData.entries.map(entry => [entry.date, entry] as [string, KpiEntry]));
    let idOffset = 0;
    Object.entries(countsByDate).forEach(([dateStr, mediaCounts]) => {
      const existing = entriesByDateMap.get(dateStr);
      const values: KpiTotals = existing ? { ...existing.values } : ({} as KpiTotals);
      Object.entries(mediaCounts).forEach(([mediaId, count]) => {
        values[`${mediaId}_scoutReplies` as KpiKey] = count;
      });
      entriesByDateMap.set(dateStr, { id: existing?.id ?? Date.now() + (idOffset++), date: dateStr, values });
    });
    const updatedEntries = Array.from(entriesByDateMap.values()).sort((a, b) => a.date.localeCompare(b.date));
    const updatedData = { ...currentUserData, entries: updatedEntries };
    setCurrentUserData(updatedData);
    setIsBulkGmailModalOpen(false);
    if (currentIdentity) {
      forceSyncNow(currentIdentity.email, driveFileId, updatedData, setDriveFileId);
    }
  };

  // Same merge pattern as handleApplyBulkGmailImport, but for a single media's CSV import,
  // writing both scoutsSent and scoutReplies for that one media per date.
  const handleApplyMediaCsvImport = (mediaId: ScoutCsvMediaId, countsByDate: Record<string, ScoutCsvDayCounts>) => {
    if (!currentUserData) return;
    const entriesByDateMap = new Map<string, KpiEntry>(currentUserData.entries.map(entry => [entry.date, entry] as [string, KpiEntry]));
    let idOffset = 0;
    Object.entries(countsByDate).forEach(([dateStr, counts]) => {
      const existing = entriesByDateMap.get(dateStr);
      const values: KpiTotals = existing ? { ...existing.values } : ({} as KpiTotals);
      values[`${mediaId}_scoutsSent` as KpiKey] = counts.scoutsSent;
      values[`${mediaId}_scoutReplies` as KpiKey] = counts.scoutReplies;
      entriesByDateMap.set(dateStr, { id: existing?.id ?? Date.now() + (idOffset++), date: dateStr, values });
    });
    const updatedEntries = Array.from(entriesByDateMap.values()).sort((a, b) => a.date.localeCompare(b.date));
    const updatedData = { ...currentUserData, entries: updatedEntries };
    setCurrentUserData(updatedData);
    setIsMediaCsvModalOpen(false);
    if (currentIdentity) {
      forceSyncNow(currentIdentity.email, driveFileId, updatedData, setDriveFileId);
    }
  };

  // Mirrors googleTaskIdsByApplicationId — kept in sync with state via the effect below, but
  // also written to directly (synchronously) inside the sync queue itself, so a second sync
  // queued right after the first always sees the first one's freshly-created task IDs instead
  // of a stale pre-sync snapshot (React's own state update wouldn't land in time otherwise).
  const googleTaskIdsRef = useRef<Record<string, string>>(currentUserData?.googleTaskIdsByApplicationId || {});
  useEffect(() => {
    googleTaskIdsRef.current = currentUserData?.googleTaskIdsByApplicationId || {};
  }, [currentUserData?.googleTaskIdsByApplicationId]);
  const tasksSyncQueueRef = useRef<Promise<void>>(Promise.resolve());
  const [tasksSyncStatus, setTasksSyncStatus] = useState<'idle' | 'loading' | 'error' | 'needs-reauth'>('idle');
  const [tasksSyncMessage, setTasksSyncMessage] = useState('');

  // Google Tasks has no time-of-day field of its own (due is a date only) — so a start time,
  // when set, is prepended straight onto the title text instead, e.g. "13:30 山田太郎（株式会社
  // サンプル）".
  const buildPipelineTaskContent = (candidate: Candidate, app: CompanyApplication) => ({
    title: `${app.scheduledTime ? `${app.scheduledTime} ` : ''}${candidate.name}（${app.companyName}）`,
    notes: `ステージ: ${app.stage}${app.nextAction ? ` / 次のアクション: ${app.nextAction}` : ''}`,
    dueDateISO: app.scheduledDate!,
  });

  /**
   * Diffs one candidate's applications (prev vs next; next=null means the candidate itself was
   * removed/hidden, so every one of its tracked tasks should be deleted) against the
   * already-synced task IDs, and pushes only the create/update/delete calls actually needed.
   * Chained through tasksSyncQueueRef so two rapid edits never race into duplicate task
   * creation for the same application.
   */
  const queueCandidateTasksSync = (prevCandidate: Candidate | undefined, nextCandidate: Candidate | null) => {
    const session = getCurrentSession();
    if (!session) return;
    const accessToken = session.accessToken;
    // A hidden candidate (via the plain 非表示 toggle, or via 掘り起しリストへの追加, which also
    // hides) is archived out of active pursuit — treat it the same as a removed candidate here
    // so its applications' reminders don't linger as open Google Tasks. Un-hiding (either
    // toggle) conversely lets any still-scheduled applications recreate their tasks normally.
    const effectiveNext = nextCandidate && !nextCandidate.isHidden ? nextCandidate : null;
    const prevApps = prevCandidate?.applications || [];
    const nextApps = effectiveNext?.applications || [];
    const nextAppById = new Map(nextApps.map(a => [a.id, a]));
    const removedAppIds = prevApps.filter(a => !nextAppById.has(a.id)).map(a => a.id);
    const prevAppById = new Map(prevApps.map(a => [a.id, a]));

    tasksSyncQueueRef.current = tasksSyncQueueRef.current.catch(() => {}).then(async () => {
      const idMap = { ...googleTaskIdsRef.current };
      let changed = false;
      try {
        setTasksSyncStatus('loading');
        for (const appId of removedAppIds) {
          const taskId = idMap[appId];
          if (taskId) {
            await deletePipelineTask(accessToken, taskId);
            delete idMap[appId];
            changed = true;
          }
        }
        if (effectiveNext) {
          for (const app of nextApps) {
            const existingTaskId = idMap[app.id];
            if (!app.scheduledDate) {
              if (existingTaskId) {
                await deletePipelineTask(accessToken, existingTaskId);
                delete idMap[app.id];
                changed = true;
              }
              continue;
            }
            const prevApp = prevAppById.get(app.id);
            const isUnchanged = prevApp && existingTaskId
              && prevApp.scheduledDate === app.scheduledDate
              && prevApp.scheduledTime === app.scheduledTime
              && prevApp.companyName === app.companyName
              && prevApp.stage === app.stage
              && prevApp.nextAction === app.nextAction;
            if (isUnchanged) continue;
            const content = buildPipelineTaskContent(effectiveNext, app);
            if (existingTaskId) {
              const taskId = await updatePipelineTask(accessToken, existingTaskId, content);
              if (taskId !== existingTaskId) { idMap[app.id] = taskId; changed = true; }
            } else {
              idMap[app.id] = await createPipelineTask(accessToken, content);
              changed = true;
            }
          }
        }
        if (changed) {
          googleTaskIdsRef.current = idMap;
          setCurrentUserData(prev => (prev ? { ...prev, googleTaskIdsByApplicationId: idMap } : prev));
        }
        setTasksSyncStatus('idle');
        setTasksSyncMessage('');
      } catch (error) {
        console.error('Failed to sync pipeline entries to Google Tasks', error);
        if (error instanceof GoogleTasksPermissionError) {
          setTasksSyncStatus('needs-reauth');
          setTasksSyncMessage('Googleタスクの利用権限がまだ許可されていません。下のボタンから許可してください。');
        } else {
          setTasksSyncStatus('error');
          setTasksSyncMessage(error instanceof Error ? error.message : 'Googleタスクへの同期に失敗しました。');
        }
      }
    });
  };

  // Manual backfill/resync — loops every visible candidate as both "prev" and "next" for the
  // same diff logic above, so already-synced tasks are left alone (isUnchanged) while anything
  // scheduled before this feature existed (or before permission was granted) gets created for
  // the first time.
  const handleSyncAllTasksNow = () => {
    (currentUserData?.candidates || []).filter(c => !c.isHidden).forEach(c => queueCandidateTasksSync(c, c));
  };

  const handleReauthorizeTasks = async () => {
    setTasksSyncStatus('loading');
    setTasksSyncMessage('');
    try {
      await reauthorizeWithConsent();
      handleSyncAllTasksNow();
    } catch (err) {
      setTasksSyncStatus('error');
      setTasksSyncMessage(err instanceof Error ? err.message : 'ログインに失敗しました。');
    }
  };

  const handleSaveCandidate = (candidateData: Candidate) => {
    // candidateData may be a copy tagged with ownerEmail/ownerLabel (e.g. edited while viewing
    // an aggregate pipeline scope on one's own candidate) — those tags are presentation-only and
    // must never be written into this account's own stored candidates. See normalize()'s
    // same-shaped strip above for why.
    const { ownerEmail, ownerLabel, ...sanitized } = candidateData;
    const prevCandidate = currentUserData?.candidates.find(c => c.id === sanitized.id);
    setCurrentUserData(prevData => {
        if (!prevData) return null;
        const existing = prevData.candidates.find(c => c.id === sanitized.id);
        let updatedCandidates;
        if (existing) {
            updatedCandidates = prevData.candidates.map(c => c.id === sanitized.id ? sanitized : c);
        } else {
            updatedCandidates = [...prevData.candidates, sanitized];
        }
        return { ...prevData, candidates: updatedCandidates };
    });
    queueCandidateTasksSync(prevCandidate, sanitized);
  };

  const handleToggleCandidateVisibility = (candidateId: string) => {
      const target = currentUserData?.candidates.find(c => c.id === candidateId);
      setCurrentUserData(prevData => {
          if (!prevData) return null;
          const updatedCandidates = prevData.candidates.map(c =>
              c.id === candidateId ? { ...c, isHidden: !c.isHidden } : c
          );
          return { ...prevData, candidates: updatedCandidates };
      });
      // Hiding archives the candidate — clear any Google Tasks tied to its applications too, so
      // an inactive pipeline entry doesn't keep showing up as an open task. Un-hiding doesn't
      // recreate them automatically; the next edit that touches the candidate will.
      if (target && !target.isHidden) {
          queueCandidateTasksSync(target, null);
      }
  };

    // Extract data for child components from the single source of truth
    const { entries, kpiTargets, weeklyKpiTargets, dailyKpiTargets, candidates } = useMemo(() => ({
      entries: currentUserData?.entries || [],
      kpiTargets: { ...defaultKpiTargets, ...(currentUserData?.kpiTargets || {}) },
      weeklyKpiTargets: { ...defaultKpiTargets, ...(currentUserData?.weeklyKpiTargets || {}) },
      dailyKpiTargets: { ...defaultKpiTargets, ...(currentUserData?.dailyKpiTargets || {}) },
      candidates: currentUserData?.candidates || [],
    }), [currentUserData]);

    // Candidates shown in the pipeline view, scoped to self / all users / a team. Non-personal
    // scopes flatten every in-scope teammate's candidates and tag each with its owner, so the
    // pipeline UI can label them and disable editing on candidates that aren't the viewer's own.
    const pipelineCandidates = useMemo(() => {
      if (pipelineScope === 'personal') return candidates;
      const emailsInScope = pipelineScope === 'team'
        ? (teams.find(t => t.id === pipelineSelectedTeamId)?.memberEmails || [])
        : pipelineScope === 'user'
        ? (pipelineSelectedUserEmail ? [pipelineSelectedUserEmail] : [])
        : Object.keys(displayedAllUsersData).filter(isEmailInSelectedDivision);
      return emailsInScope.flatMap(email => {
        const resolved = resolveUserDataEntry(displayedAllUsersData, email);
        if (!resolved) return [];
        const [ownerEmail, data] = resolved;
        const ownerLabel = data.displayName || ownerEmail;
        return (data.candidates || []).map(c => ({ ...c, ownerEmail, ownerLabel }));
      });
    }, [pipelineScope, pipelineSelectedTeamId, pipelineSelectedUserEmail, teams, displayedAllUsersData, candidates, isEmailInSelectedDivision]);

    // Options for the pipeline's per-user selector, sorted by display name.
    const pipelineUserOptions = useMemo(() => {
      return Object.entries(displayedAllUsersData)
        .map(([email, data]: [string, UserData]) => ({ email, label: data.displayName || email }))
        .sort((a, b) => a.label.localeCompare(b.label, 'ja'));
    }, [displayedAllUsersData]);

    // Same list, narrowed to the header's selected division — used for the 全ユーザー tab's
    // 比較するユーザー picker, so the picker itself (not just the resulting totals) reflects
    // the current BCA/F+/AC scope. pipelineUserOptions itself stays unfiltered for contexts that
    // need every known user regardless of division (TeamsModal's department assignment, the
    // pipeline's single-user lookup).
    const divisionScopedUserOptions = useMemo(() => {
      return pipelineUserOptions.filter(u => isEmailInSelectedDivision(u.email));
    }, [pipelineUserOptions, isEmailInSelectedDivision]);

    // 比較するユーザーの選択肢をチーム単位でグルーピング — メンバーが増えるほどフラットな
    // 一覧の表示面積が広がっていくのを防ぐため、チーム→メンバーの2階層で折りたたんで表示する。
    // 複数チームに所属するメンバーはそれぞれのチームのグループに重複して現れる（チーム別タブの
    // 挙動と同じ）。どのチームにも属さないメンバーは「未所属」グループにまとめる。
    const comparisonTeamGroups = useMemo(() => {
      const labelByEmail = new Map<string, string>(divisionScopedUserOptions.map(u => [u.email, u.label]));
      const assigned = new Set<string>();
      const groups = teams
        .map(team => {
          const memberEmails: string[] = Array.from(new Set(
            team.memberEmails
              .map(email => resolveUserDataEntry(displayedAllUsersData, email)?.[0] || email)
              .filter(email => labelByEmail.has(email))
          ));
          memberEmails.forEach(email => assigned.add(email));
          return { id: team.id, name: team.name, members: memberEmails.map(email => ({ email, label: labelByEmail.get(email)! })) };
        })
        .filter(g => g.members.length > 0);
      const unassigned = divisionScopedUserOptions.filter(u => !assigned.has(u.email));
      return { groups, unassigned };
    }, [teams, divisionScopedUserOptions, displayedAllUsersData]);

    const [expandedComparisonGroups, setExpandedComparisonGroups] = useState<Record<string, boolean>>({});
    const toggleComparisonGroupExpanded = (groupId: string) => {
      setExpandedComparisonGroups(prev => ({ ...prev, [groupId]: !prev[groupId] }));
    };
    const setComparisonGroupSelected = (emails: string[], select: boolean) => {
      setComparisonUserEmails(prev => select
        ? Array.from(new Set([...prev, ...emails]))
        : prev.filter(e => !emails.includes(e)));
    };

    // The 全ユーザー tab's ad-hoc comparison selection — an empty selection means "no filter".
    const comparisonUsers = useMemo(() => {
      const inDivision = users.filter(isEmailInSelectedDivision);
      if (comparisonUserEmails.length === 0) return inDivision;
      return inDivision.filter(u => comparisonUserEmails.includes(u));
    }, [users, comparisonUserEmails, isEmailInSelectedDivision]);

    const toggleComparisonUser = (email: string) => {
      setComparisonUserEmails(prev =>
        prev.includes(email) ? prev.filter(e => e !== email) : [...prev, email]
      );
    };


  // Independent from viewDate (which drives the KPI-entry calendar grid, a different concept) —
  // lets 媒体別月次進捗/当月日次パフォーマンストレンド move to any month via 前月/次月, instead
  // of always being locked to the real current month.
  const [monthlyProgressViewDate, setMonthlyProgressViewDate] = useState(new Date());
  const isCurrentRealMonth = monthlyProgressViewDate.getFullYear() === new Date().getFullYear()
    && monthlyProgressViewDate.getMonth() === new Date().getMonth();
  const handleShiftMonthlyProgressMonth = (offset: number) => {
    setMonthlyProgressViewDate(d => new Date(d.getFullYear(), d.getMonth() + offset, 1));
  };

  const monthlyTotals = useMemo<KpiTotals>(() => {
    const year = monthlyProgressViewDate.getFullYear();
    const month = monthlyProgressViewDate.getMonth();
    const start = new Date(year, month, 1);
    const end = new Date(year, month + 1, 0, 23, 59, 59);
    return calculateTotalsForRange(entries, allMedia, start, end);
  }, [entries, allMedia, monthlyProgressViewDate]);

  const todayTotals = useMemo<KpiTotals>(() => {
    const todayStr = new Date().toLocaleDateString('sv-SE'); // YYYY-MM-DD format
    const todayEntry = entries.find(e => e.date === todayStr);
    return todayEntry ? todayEntry.values : buildAllKpiKeys(allMedia).reduce((acc, key) => {
      acc[key] = 0;
      return acc;
    }, {} as KpiTotals);
  }, [entries, allMedia]);


  const currentMonthPerformanceChartData = useMemo(() => {
    const currentMonth = monthlyProgressViewDate.getMonth();
    const currentYear = monthlyProgressViewDate.getFullYear();
    const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
    
    const labels = Array.from({ length: daysInMonth }, (_, i) => `${i + 1}`);
    const cumulativeScouts = new Array(daysInMonth).fill(0);
    const cumulativeReplies = new Array(daysInMonth).fill(0);

    entries.forEach(entry => {
        const entryDate = new Date(entry.date);
        if (entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear) {
            const day = entryDate.getDate() - 1;
            const scouts = allMedia.reduce((sum, source) => sum + (entry.values[`${source.id}_scoutsSent` as KpiKey] || 0), 0);
            const replies = allMedia.reduce((sum, source) => sum + (entry.values[`${source.id}_scoutReplies` as KpiKey] || 0), 0);

            cumulativeScouts[day] += scouts;
            cumulativeReplies[day] += replies;
        }
    });

    for(let i = 1; i < daysInMonth; i++) {
        cumulativeScouts[i] += cumulativeScouts[i - 1];
        cumulativeReplies[i] += cumulativeReplies[i - 1];
    }
    
    const replyRates = cumulativeScouts.map((scouts, i) => scouts > 0 ? (cumulativeReplies[i] / scouts) * 100 : 0);

    return {
        labels,
        datasets: [
            {
                label: '累積スカウト数',
                data: cumulativeScouts,
                borderColor: 'rgb(54, 162, 235)',
                backgroundColor: 'rgba(54, 162, 235, 0.5)',
                yAxisID: 'y-axis-count',
            },
            {
                label: '累積返信率',
                data: replyRates,
                borderColor: 'rgb(255, 99, 132)',
                backgroundColor: 'rgba(255, 99, 132, 0.5)',
                yAxisID: 'y-axis-rate',
            },
        ],
    };
  }, [entries, allMedia, monthlyProgressViewDate]);

  // Resolved to displayedAllUsersData's canonical (Google-cased) keys — memberEmails is
  // free-typed, so an exact-match filter/index against it would silently drop any member whose
  // stored casing doesn't match their actual sign-in email (see resolveUserDataEntry).
  const selectedTeamMemberEmails = useMemo(() => {
    if (!selectedTeamId) return [];
    const memberEmails = teams.find(t => t.id === selectedTeamId)?.memberEmails || [];
    const resolved = memberEmails.map(email => resolveUserDataEntry(displayedAllUsersData, email)?.[0] || email);
    return Array.from(new Set(resolved)).filter(isEmailInSelectedDivision);
  }, [teams, selectedTeamId, displayedAllUsersData, isEmailInSelectedDivision]);

  const dayOfWeekReplyRateData = useMemo(() => {
    const days = ['日', '月', '火', '水', '木', '金', '土'];
    const scoutsByDay = Array(7).fill(0);
    const repliesByDay = Array(7).fill(0);

    const allEntries = view === 'all_users_kpi'
      ? comparisonUsers.flatMap(email => displayedAllUsersData[email]?.entries || [])
      : view === 'team_kpi'
      ? Object.entries(displayedAllUsersData).filter(([email]) => selectedTeamMemberEmails.includes(email)).flatMap(([, d]: [string, UserData]) => d.entries)
      : entries;

    // 全ユーザー/チーム別タブでは月次進捗・ファネル分析と同じ期間（カスタム指定時はその範囲、未指定
    // 時は今月）に絞り込む。個人実績タブは独立した期間指定（未指定時は全期間、従来通り）。
    const periodFilteredEntries = view === 'all_users_kpi' || view === 'team_kpi'
      ? allEntries.filter(entry => {
          const entryTime = new Date(entry.date).getTime();
          if (dashboardPeriodOverride) {
            return entryTime >= dashboardPeriodOverride.start.getTime() && entryTime <= dashboardPeriodOverride.end.getTime();
          }
          const entryDate = new Date(entry.date);
          const now = new Date();
          return entryDate.getMonth() === now.getMonth() && entryDate.getFullYear() === now.getFullYear();
        })
      : personalDowPeriodOverride
      ? allEntries.filter(entry => {
          const entryTime = new Date(entry.date).getTime();
          return entryTime >= personalDowPeriodOverride.start.getTime() && entryTime <= personalDowPeriodOverride.end.getTime();
        })
      : allEntries;
    if (periodFilteredEntries.length === 0) return null;

    periodFilteredEntries.forEach(entry => {
        const dayOfWeek = new Date(entry.date).getDay();
        const scouts = allMedia.reduce((sum, source) => sum + (entry.values[`${source.id}_scoutsSent` as KpiKey] || 0), 0);
        const replies = allMedia.reduce((sum, source) => sum + (entry.values[`${source.id}_scoutReplies` as KpiKey] || 0), 0);
        scoutsByDay[dayOfWeek] += scouts;
        repliesByDay[dayOfWeek] += replies;
    });

    const rates = scoutsByDay.map((scouts, i) => scouts > 0 ? (repliesByDay[i] / scouts) * 100 : 0);

    return {
        labels: days,
        datasets: [
            {
                label: '曜日別返信率',
                data: rates,
                backgroundColor: 'rgba(75, 192, 192, 0.6)',
                borderColor: 'rgba(75, 192, 192, 1)',
                borderWidth: 1,
            }
        ]
    };
  }, [entries, view, displayedAllUsersData, selectedTeamMemberEmails, allMedia, comparisonUsers, dashboardPeriodOverride, personalDowPeriodOverride]);


  const weeklySummaryData = useMemo<WeeklyData>(() => {
      const weekStart = viewWeekStartDate.getTime();
      const weekEnd = new Date(viewWeekStartDate).setDate(viewWeekStartDate.getDate() + 6);

      const weeklyEntries = entries.filter(entry => {
          const entryTime = new Date(entry.date).getTime();
          return entryTime >= weekStart && entryTime <= weekEnd;
      });

      const allKeys = buildAllKpiKeys(activeMedia);
      const weeklyTotals = weeklyEntries.reduce((acc, entry) => {
          allKeys.forEach(key => {
              acc[key] = (acc[key] || 0) + (entry.values[key] || 0);
          });
          return acc;
      }, {} as KpiTotals);

      const mediaStats = activeMedia.map(source => {
          const sourceKey = source.id;
          return {
              source: source.name,
              id: source.id,
              scoutsSent: weeklyTotals[`${sourceKey}_scoutsSent` as KpiKey] || 0,
              scoutReplies: weeklyTotals[`${sourceKey}_scoutReplies` as KpiKey] || 0,
              effectiveReplies: weeklyTotals[`${sourceKey}_effectiveReplies` as KpiKey] || 0,
              documentsCollected: weeklyTotals[`${sourceKey}_documentsCollected` as KpiKey] || 0,
              effectiveDocumentsCollected: weeklyTotals[`${sourceKey}_effectiveDocumentsCollected` as KpiKey] || 0,
              initialInterviews: weeklyTotals[`${sourceKey}_initialInterviews` as KpiKey] || 0,
              effectiveInitialInterviews: weeklyTotals[`${sourceKey}_effectiveInitialInterviews` as KpiKey] || 0,
          };
      });

      const totalCandidatesSubmitted = weeklyTotals.candidatesSubmitted || 0;
      const totalInitialInterviews = mediaStats.reduce((sum, stat) => sum + stat.initialInterviews, 0);

      return { mediaStats, totalCandidatesSubmitted, totalInitialInterviews };
  }, [entries, viewWeekStartDate, activeMedia]);
  
  const entriesByDate = useMemo(() => {
    return new Map(entries.map(entry => [entry.date, entry.values]));
  }, [entries]);

  if (!isInitialized) {
      return <div className="loading-container">読み込み中...</div>;
  }

  if (!currentIdentity) {
    const lastEmail = getLastKnownEmail();
    const pendingOnThisDevice = lastEmail ? hasPendingSync(lastEmail) : false;
    return (
        <div className="login-container">
            <div className="login-box">
                <h1>KPI管理くん</h1>
                <p className="login-error">{authError}</p>
                <button
                    type="button"
                    className="login-button"
                    onClick={handleGoogleSignIn}
                    disabled={isSigningIn}
                >
                    {isSigningIn
                        ? 'ログイン中...'
                        : lastEmail
                        ? `${lastEmail} で続ける`
                        : 'Googleでログイン'}
                </button>
                <p style={{ marginTop: '1rem', fontSize: '0.85rem', color: '#666' }}>
                    bloom-firm.com のGoogleアカウントでログインしてください。
                </p>
                {pendingOnThisDevice && (
                    <div className="sync-error-banner" style={{ marginTop: '1rem', textAlign: 'left' }}>
                        <strong>⚠️ この端末に、まだGoogleドライブへ保存されていない入力データがあります。</strong>
                        <p style={{ margin: '0.5rem 0 0' }}>データを消さずに復旧するには：</p>
                        <ol style={{ margin: '0.25rem 0 0', paddingLeft: '1.25rem' }}>
                            <li>この<strong>同じ端末・同じブラウザ</strong>で、上のボタンから{lastEmail ? `「${lastEmail}」` : ''}で再度ログインしてください（別の端末やブラウザでログインしても、この端末に残っているデータは反映されません）。</li>
                            <li>ログインすると自動的に保存が再試行されます。この警告が消えれば保存完了です。</li>
                            <li>ログインし直すまでは、ブラウザの履歴・データ消去やシークレットモードでのアクセスは避けてください。</li>
                        </ol>
                    </div>
                )}
            </div>
        </div>
    );
  }

  if (isLoadingUserData) {
      return <div className="loading-container">Googleドライブからデータを読み込み中...</div>;
  }

  if (isLoadingMediaConfig) {
      return <div className="loading-container">媒体設定を読み込み中...</div>;
  }

  return (
    <div className="app-container">
      {legacyMigrationChoices && (
        <div className="modal-overlay" role="dialog" aria-modal="true" aria-labelledby="migration-modal-title">
          <div className="modal-content">
            <div className="modal-header">
              <h3 id="migration-modal-title">ローカルデータの引き継ぎ</h3>
            </div>
            <div className="modal-body">
              <p>このブラウザに以前のバージョンのデータが見つかりました。どのユーザーのデータを引き継ぎますか？</p>
              <ul className="user-management-list">
                {legacyMigrationChoices.map(name => (
                  <li key={name} className="user-management-item">
                    <span className="user-management-name">{name}</span>
                    <button className="submit-button" onClick={() => handleClaimLegacyData(name)}>これを引き継ぐ</button>
                  </li>
                ))}
              </ul>
            </div>
            <div className="modal-footer">
              <button type="button" className="cancel-button" onClick={() => handleClaimLegacyData(null)}>
                引き継がず新規で始める
              </button>
            </div>
          </div>
        </div>
      )}
      {isTeamsModalOpen && (
        <TeamsModal
          teams={teams}
          isEditable={isTeamsEditable}
          isAdmin={isTeamsAdmin}
          authorizedEditorEmails={teamsAuthorizedEditors}
          userOptions={pipelineUserOptions}
          memberDepartments={memberDepartments}
          onClose={() => setIsTeamsModalOpen(false)}
          onCreateTeam={handleCreateTeam}
          onRenameTeam={handleRenameTeam}
          onDeleteTeam={handleDeleteTeam}
          onAddMember={handleAddTeamMember}
          onRemoveMember={handleRemoveTeamMember}
          onGrantEditor={handleGrantTeamsEditor}
          onRevokeEditor={handleRevokeTeamsEditor}
          onSetMemberDepartment={handleSetMemberDepartment}
        />
      )}
      {isMediaModalOpen && (
        <MediaModal
          allMedia={allMedia}
          isEditable={isMediaEditable}
          onClose={() => setIsMediaModalOpen(false)}
          onCreateMedia={handleCreateMedia}
          onRenameMedia={handleRenameMedia}
          onArchiveMedia={handleArchiveMedia}
          onUnarchiveMedia={handleUnarchiveMedia}
          onSetFeeRate={handleSetMediaFeeRate}
          onRefresh={refreshMediaConfig}
        />
      )}
      {selectedDate && (
        <DateEntryModal
          date={selectedDate}
          initialValues={entriesByDate.get(selectedDate) || null}
          activeMedia={activeMedia}
          onSave={handleSaveEntry}
          onNavigate={handleNavigateEntryDate}
          onClose={() => setSelectedDate(null)}
        />
      )}
      {isBulkGmailModalOpen && (
        <BulkGmailReplyImportModal
          allMedia={allMedia}
          entriesByDate={entriesByDate}
          onApply={handleApplyBulkGmailImport}
          onClose={() => setIsBulkGmailModalOpen(false)}
        />
      )}
      {isMediaCsvModalOpen && (
        <MediaCsvImportModal
          allMedia={allMedia}
          entriesByDate={entriesByDate}
          onApply={handleApplyMediaCsvImport}
          onClose={() => setIsMediaCsvModalOpen(false)}
        />
      )}

      {hasSyncError && (
        <div className="sync-error-banner">
          <strong>⚠️ 一部の変更がまだGoogleドライブに保存されていません。</strong>
          <p style={{ margin: '0.5rem 0 0' }}>自動で再試行していますが、解決しない場合は次の手順をお試しください：</p>
          <ol style={{ margin: '0.25rem 0 0', paddingLeft: '1.25rem' }}>
            <li>この<strong>同じ端末・同じブラウザ</strong>のまま、一度ログアウトして再度ログインしてください（別の端末・別のブラウザではこのデータは見られません）。</li>
            <li>この警告が消えれば保存完了です。消えない場合はネットワーク接続を確認し、しばらく待ってから再度お試しください。</li>
          </ol>
        </div>
      )}
      <header className="app-main-header">
        <h1 className="app-title">KPI管理くん</h1>
        <div className="division-switcher" role="group" aria-label="事業部切り替え">
          <button onClick={() => setSelectedDivision('BCA')} disabled={selectedDivision === 'BCA'} title="F+とACを合わせた全体表示">BCA</button>
          <button onClick={() => setSelectedDivision('F+')} disabled={selectedDivision === 'F+'}>F+</button>
          <button onClick={() => setSelectedDivision('AC')} disabled={selectedDivision === 'AC'}>AC</button>
        </div>
        <div className="header-controls">
          <div className="view-switcher">
            <button onClick={() => setView('personal_kpi')} disabled={view === 'personal_kpi'}>個人実績</button>
            <button onClick={() => setView('all_users_kpi')} disabled={view === 'all_users_kpi'}>全ユーザー</button>
            <button onClick={() => setView('team_kpi')} disabled={view === 'team_kpi'}>チーム別</button>
            <button onClick={() => setView('pipeline')} disabled={view === 'pipeline'}>候補者パイプライン</button>
          </div>
          <div className="user-controls">
            {currentIdentity && (
              <span className="sr-only">{currentIdentity.email}</span>
            )}
            {currentIdentity && (
              isEditingDisplayName ? (
                <span style={{ display: 'flex', gap: '0.25rem', alignItems: 'center' }}>
                  <input
                    type="text"
                    value={editedDisplayName}
                    onChange={(e) => setEditedDisplayName(e.target.value)}
                    autoFocus
                    style={{ width: '8rem' }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' && editedDisplayName.trim()) {
                        handleSaveDisplayName(editedDisplayName.trim());
                        setIsEditingDisplayName(false);
                      } else if (e.key === 'Escape') {
                        setIsEditingDisplayName(false);
                      }
                    }}
                  />
                  <button
                    className="save-user-button"
                    disabled={!editedDisplayName.trim()}
                    onClick={() => {
                      handleSaveDisplayName(editedDisplayName.trim());
                      setIsEditingDisplayName(false);
                    }}
                  >
                    保存
                  </button>
                  <button className="cancel-user-button" onClick={() => setIsEditingDisplayName(false)}>キャンセル</button>
                </span>
              ) : (
                <span
                  style={{ fontSize: '0.9rem', color: '#333', cursor: 'pointer' }}
                  title="クリックして表示名を変更"
                  onClick={() => {
                    setEditedDisplayName(currentUserData?.displayName || currentIdentity.name);
                    setIsEditingDisplayName(true);
                  }}
                >
                  {currentUserData?.displayName || currentIdentity.name} ✎
                </span>
              )
            )}
            {currentIdentity && (
              <select
                className="own-department-select"
                value={memberDepartments[currentIdentity.email] || ''}
                onChange={(e) => handleSetOwnDepartment((e.target.value || null) as Department | null)}
                title="自分の所属部署"
              >
                <option value="">所属未設定</option>
                <option value="F+">Firm+</option>
                <option value="AC">AssetCareer</option>
              </select>
            )}
            <button onClick={() => setIsTeamsModalOpen(true)}>チーム管理</button>
            <button onClick={() => setIsMediaModalOpen(true)}>媒体管理</button>
            <button onClick={handleLogout} className="logout-button">ログアウト</button>
          </div>
        </div>
      </header>

      <main>
        {view === 'personal_kpi' && (
          <>
            <section aria-labelledby="daily-progress-title">
                <h2 
                  id="daily-progress-title"
                  className="section-title collapsible-header"
                  onClick={() => toggleSection('dailyProgress')}
                  onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('dailyProgress'); } }}
                  role="button" tabIndex={0} aria-expanded={sectionVisibility.dailyProgress} aria-controls="daily-progress-content"
                >
                  <span>本日の進捗</span>
                  <span className={`toggle-icon ${sectionVisibility.dailyProgress ? 'open' : ''}`}>▼</span>
                </h2>
                <div id="daily-progress-content" className={`collapsible-content ${sectionVisibility.dailyProgress ? 'open' : ''}`}>
                   <DailyProgress activeMedia={activeMedia} todayTotals={todayTotals} dailyKpiTargets={dailyKpiTargets} />
                </div>
            </section>
            
             <div className="sync-status-bar">
               <span>
                 最終同期: {lastSyncedAt ? new Date(lastSyncedAt).toLocaleString('ja-JP') : '未同期'}
               </span>
               <span style={{ display: 'flex', gap: '0.5rem' }}>
                 <button type="button" onClick={handleForceSync} disabled={isForcingSync} className="secondary-action-button">
                   {isForcingSync ? '同期中...' : '今すぐ同期'}
                 </button>
                 <button type="button" onClick={() => setIsBulkGmailModalOpen(true)} className="secondary-action-button">
                   Gmailから過去の返信を一括取得
                 </button>
                 <button type="button" onClick={() => setIsMediaCsvModalOpen(true)} className="secondary-action-button">
                   媒体CSVから実績を取り込む
                 </button>
               </span>
             </div>

             <section aria-labelledby="calendar-title">
              <h2
                id="calendar-title"
                className="section-title collapsible-header"
                onClick={() => toggleSection('calendar')}
                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('calendar'); } }}
                role="button"
                tabIndex={0}
                aria-expanded={sectionVisibility.calendar}
                aria-controls="calendar-content"
              >
                <span>実績カレンダー</span>
                <span className={`toggle-icon ${sectionVisibility.calendar ? 'open' : ''}`}>▼</span>
              </h2>
              <div id="calendar-content" className={`collapsible-content ${sectionVisibility.calendar ? 'open' : ''}`}>
                <CalendarView
                  viewDate={viewDate}
                  entriesByDate={entriesByDate}
                  allMedia={allMedia}
                  onDayClick={setSelectedDate}
                  onPrevMonth={() => setViewDate(d => new Date(d.getFullYear(), d.getMonth() - 1, 1))}
                  onNextMonth={() => setViewDate(d => new Date(d.getFullYear(), d.getMonth() + 1, 1))}
                />
              </div>
            </section>
            
            <section aria-labelledby="weekly-summary-title">
              <h2 
                id="weekly-summary-title"
                className="section-title collapsible-header"
                onClick={() => toggleSection('weeklySummary')}
                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('weeklySummary'); } }}
                role="button" tabIndex={0} aria-expanded={sectionVisibility.weeklySummary} aria-controls="weekly-summary-content"
              >
                <span>週間サマリー</span>
                <span className={`toggle-icon ${sectionVisibility.weeklySummary ? 'open' : ''}`}>▼</span>
              </h2>
              <div id="weekly-summary-content" className={`collapsible-content ${sectionVisibility.weeklySummary ? 'open' : ''}`}>
                 <WeeklySummary
                    weekStartDate={viewWeekStartDate}
                    data={weeklySummaryData}
                    weeklyKpiTargets={weeklyKpiTargets}
                    onPrevWeek={() => setViewWeekStartDate(d => new Date(d.setDate(d.getDate() - 7)))}
                    onNextWeek={() => setViewWeekStartDate(d => new Date(d.setDate(d.getDate() + 7)))}
                />
              </div>
            </section>
            
            <section aria-labelledby="media-progress-title">
                <h2 
                  id="media-progress-title"
                  className="section-title collapsible-header"
                  onClick={() => toggleSection('mediaProgress')}
                  onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('mediaProgress'); } }}
                  role="button" tabIndex={0} aria-expanded={sectionVisibility.mediaProgress} aria-controls="media-progress-content"
                >
                  <span>媒体別 月次進捗（{monthlyProgressViewDate.getFullYear()}年{monthlyProgressViewDate.getMonth() + 1}月）</span>
                  <span className={`toggle-icon ${sectionVisibility.mediaProgress ? 'open' : ''}`}>▼</span>
                </h2>
                <div id="media-progress-content" className={`collapsible-content ${sectionVisibility.mediaProgress ? 'open' : ''}`}>
                  <div className="custom-period-export-bar" onClick={(e) => e.stopPropagation()}>
                    <button type="button" onClick={() => handleShiftMonthlyProgressMonth(-1)} className="secondary-action-button month-shift-button">&lt; 前月</button>
                    <span>{monthlyProgressViewDate.getFullYear()}年{monthlyProgressViewDate.getMonth() + 1}月</span>
                    <button type="button" onClick={() => handleShiftMonthlyProgressMonth(1)} className="secondary-action-button month-shift-button">次月 &gt;</button>
                    {!isCurrentRealMonth && (
                      <button type="button" onClick={() => setMonthlyProgressViewDate(new Date())} className="secondary-action-button">今月に戻す</button>
                    )}
                  </div>
                  {!isCurrentRealMonth && (
                    <p className="gmail-scout-message">
                      目標は現在設定されている月次目標です（表示中の月の当時の目標とは異なる場合があります）。
                    </p>
                  )}
                  <div className="media-dashboard kpi-dashboard">
                     {activeMedia.map(source => (
                          <MediaKpiCard
                              key={source.id}
                              source={source}
                              monthlyTotals={monthlyTotals}
                              kpiTargets={kpiTargets}
                          />
                      ))}
                  </div>
                </div>
            </section>

             <section aria-labelledby="current-month-performance-title">
               <h2
                id="current-month-performance-title"
                className="section-title collapsible-header"
                onClick={() => toggleSection('monthlyPerformance')}
                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('monthlyPerformance'); } }}
                role="button"
                tabIndex={0}
                aria-expanded={sectionVisibility.monthlyPerformance}
                aria-controls="current-month-performance-content"
              >
                <span>日次パフォーマンストレンド（{monthlyProgressViewDate.getFullYear()}年{monthlyProgressViewDate.getMonth() + 1}月）</span>
                <span className={`toggle-icon ${sectionVisibility.monthlyPerformance ? 'open' : ''}`}>▼</span>
              </h2>
              <div id="current-month-performance-content" className={`collapsible-content ${sectionVisibility.monthlyPerformance ? 'open' : ''}`}>
                 <div className="custom-period-export-bar" onClick={(e) => e.stopPropagation()}>
                   <button type="button" onClick={() => handleShiftMonthlyProgressMonth(-1)} className="secondary-action-button month-shift-button">&lt; 前月</button>
                   <span>{monthlyProgressViewDate.getFullYear()}年{monthlyProgressViewDate.getMonth() + 1}月</span>
                   <button type="button" onClick={() => handleShiftMonthlyProgressMonth(1)} className="secondary-action-button month-shift-button">次月 &gt;</button>
                   {!isCurrentRealMonth && (
                     <button type="button" onClick={() => setMonthlyProgressViewDate(new Date())} className="secondary-action-button">今月に戻す</button>
                   )}
                 </div>
                 <CurrentMonthPerformanceChart data={currentMonthPerformanceChartData} />
              </div>
            </section>
            
             <section aria-labelledby="month-over-month-performance-title">
               <h2
                id="month-over-month-performance-title"
                className="section-title collapsible-header"
                onClick={() => toggleSection('monthOverMonthPerformance')}
                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('monthOverMonthPerformance'); } }}
                role="button"
                tabIndex={0}
                aria-expanded={sectionVisibility.monthOverMonthPerformance}
                aria-controls="month-over-month-performance-content"
              >
                <span>月別パフォーマンストレンド</span>
                <span className={`toggle-icon ${sectionVisibility.monthOverMonthPerformance ? 'open' : ''}`}>▼</span>
              </h2>
              <div id="month-over-month-performance-content" className={`collapsible-content ${sectionVisibility.monthOverMonthPerformance ? 'open' : ''}`}>
                 <MonthlyTrendChart perUserEntries={[{ label: currentUserData?.displayName || currentIdentity?.name || '自分', entries }]} allMedia={allMedia} />
              </div>
            </section>
            
            <section aria-labelledby="day-of-week-rate-title">
               <h2
                id="day-of-week-rate-title"
                className="section-title collapsible-header"
                onClick={() => toggleSection('dayOfWeekRate')}
                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('dayOfWeekRate'); } }}
                role="button"
                tabIndex={0}
                aria-expanded={sectionVisibility.dayOfWeekRate}
                aria-controls="day-of-week-rate-content"
              >
                <span>曜日別 累積返信率（{personalDowPeriodOverride ? '指定期間' : '全期間'}）</span>
                <span className={`toggle-icon ${sectionVisibility.dayOfWeekRate ? 'open' : ''}`}>▼</span>
              </h2>
              <div id="day-of-week-rate-content" className={`collapsible-content ${sectionVisibility.dayOfWeekRate ? 'open' : ''}`}>
                <div className="custom-period-export-bar" onClick={(e) => e.stopPropagation()}>
                  <span>表示期間（未入力の場合は全期間）:</span>
                  <input type="date" value={personalDowStartDate} onChange={(e) => setPersonalDowStartDate(e.target.value)} aria-label="開始日" />
                  <span>〜</span>
                  <input type="date" value={personalDowEndDate} onChange={(e) => setPersonalDowEndDate(e.target.value)} aria-label="終了日" />
                  <button onClick={handleTogglePersonalDowPeriod} className="secondary-action-button">
                    {personalDowPeriodOverride ? '全期間表示に戻す' : '期間で絞り込みを有効にする'}
                  </button>
                </div>
                {dayOfWeekReplyRateData ? (
                    <DayOfWeekReplyRateChart data={dayOfWeekReplyRateData} />
                ) : (
                    <p className="no-data-message">実績データがありません。</p>
                )}
              </div>
            </section>
            
            <section aria-labelledby="custom-report-title">
               <h2
                id="custom-report-title"
                className="section-title collapsible-header"
                onClick={() => toggleSection('customPeriodReport')}
                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('customPeriodReport'); } }}
                role="button"
                tabIndex={0}
                aria-expanded={sectionVisibility.customPeriodReport}
                aria-controls="custom-report-content"
              >
                <span>カスタム期間レポート</span>
                <span className={`toggle-icon ${sectionVisibility.customPeriodReport ? 'open' : ''}`}>▼</span>
              </h2>
              <div id="custom-report-content" className={`collapsible-content ${sectionVisibility.customPeriodReport ? 'open' : ''}`}>
                <CustomPeriodReport entries={entries} allMedia={allMedia} />
              </div>
            </section>
            
            <section aria-labelledby="target-settings-title">
                <h2
                    id="target-settings-title"
                    className="section-title collapsible-header"
                    onClick={() => toggleSection('monthlyTargetSettings')}
                    onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('monthlyTargetSettings'); } }}
                    role="button"
                    tabIndex={0}
                    aria-expanded={sectionVisibility.monthlyTargetSettings}
                    aria-controls="target-settings-content"
                >
                    <span>月次目標設定</span>
                    <span className={`toggle-icon ${sectionVisibility.monthlyTargetSettings ? 'open' : ''}`}>▼</span>
                </h2>
                <div id="target-settings-content" className={`collapsible-content ${sectionVisibility.monthlyTargetSettings ? 'open' : ''}`}>
                    <form className="modal-body" style={{padding:0}}>
                        <fieldset className="general-kpi-fieldset">
                          <legend className="sr-only">全体実績 目標</legend>
                          {(Object.keys(GENERAL_KPIS) as (keyof typeof GENERAL_KPIS)[]).map(key => (
                            <div key={key} className="form-group">
                              <label htmlFor={`target-${key}`}>{GENERAL_KPIS[key].label}</label>
                              <input
                                type="number"
                                id={`target-${key}`}
                                name={key}
                                value={kpiTargets[key] || ''}
                                onChange={handleTargetChange}
                                min="0"
                                placeholder="0"
                              />
                            </div>
                          ))}
                        </fieldset>

                         <div className="media-kpi-section">
                           <h3 className="sub-section-title">媒体別実績 目標</h3>
                           <div className="media-kpi-grid">
                             {activeMedia.map(source => {
                                 const sourceKey = source.id;
                                 const fields: {key: KpiKey, label: string}[] = [
                                     {key: `${sourceKey}_scoutsSent`, label: 'スカウト数'},
                                     {key: `${sourceKey}_scoutReplies`, label: 'スカウト返信数'},
                                     {key: `${sourceKey}_effectiveReplies`, label: '有効返信数'},
                                     {key: `${sourceKey}_documentsCollected`, label: '書類回収数'},
                                     {key: `${sourceKey}_effectiveDocumentsCollected`, label: '有効書類回収数'},
                                     {key: `${sourceKey}_initialInterviews`, label: '初回面談数'},
                                     {key: `${sourceKey}_effectiveInitialInterviews`, label: '初回有効面談数'},
                                 ];
                                return (
                                 <fieldset key={source.id} className="media-fieldset">
                                     <legend>{source.name}</legend>
                                     <div className="inputs-wrapper">
                                     {fields.map(field => (
                                          <div key={field.key} className="form-group">
                                             <label htmlFor={`target-${field.key}`}>{field.label}</label>
                                             <input
                                                 type="number"
                                                 id={`target-${field.key}`}
                                                 name={field.key}
                                                 value={kpiTargets[field.key] || ''}
                                                 onChange={handleTargetChange}
                                                 min="0"
                                                 placeholder="0"
                                             />
                                         </div>
                                     ))}
                                     </div>
                                 </fieldset>
                                 )
                             })}
                           </div>
                         </div>
                    </form>
                </div>
            </section>
            
            <section aria-labelledby="weekly-target-settings-title">
                <h2
                    id="weekly-target-settings-title"
                    className="section-title collapsible-header"
                    onClick={() => toggleSection('weeklyTargetSettings')}
                    onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('weeklyTargetSettings'); } }}
                    role="button"
                    tabIndex={0}
                    aria-expanded={sectionVisibility.weeklyTargetSettings}
                    aria-controls="weekly-target-settings-content"
                >
                    <span>週次目標設定</span>
                    <span className={`toggle-icon ${sectionVisibility.weeklyTargetSettings ? 'open' : ''}`}>▼</span>
                </h2>
                <div id="weekly-target-settings-content" className={`collapsible-content ${sectionVisibility.weeklyTargetSettings ? 'open' : ''}`}>
                    <form className="modal-body" style={{padding:0}}>
                         <div className="media-kpi-section">
                           <h3 className="sub-section-title">媒体別実績 週間目標</h3>
                           <div className="media-kpi-grid">
                             {activeMedia.map(source => {
                                 const sourceKey = source.id;
                                  const fields: {key: KpiKey, label: string}[] = [
                                     {key: `${sourceKey}_scoutsSent`, label: 'スカウト数'},
                                     {key: `${sourceKey}_scoutReplies`, label: 'スカウト返信数'},
                                     {key: `${sourceKey}_effectiveReplies`, label: '有効返信数'},
                                     {key: `${sourceKey}_documentsCollected`, label: '書類回収数'},
                                     {key: `${sourceKey}_effectiveDocumentsCollected`, label: '有効書類回収数'},
                                     {key: `${sourceKey}_initialInterviews`, label: '初回面談数'},
                                     {key: `${sourceKey}_effectiveInitialInterviews`, label: '初回有効面談数'},
                                 ];
                                return (
                                 <fieldset key={source.id} className="media-fieldset">
                                     <legend>{source.name}</legend>
                                     <div className="inputs-wrapper">
                                     {fields.map(field => (
                                          <div key={field.key} className="form-group">
                                             <label htmlFor={`weekly-target-${field.key}`}>{field.label}</label>
                                             <input
                                                 type="number"
                                                 id={`weekly-target-${field.key}`}
                                                 name={field.key}
                                                 value={weeklyKpiTargets[field.key] || ''}
                                                 onChange={handleWeeklyTargetChange}
                                                 min="0"
                                                 placeholder="0"
                                             />
                                         </div>
                                     ))}
                                     </div>
                                 </fieldset>
                                 )
                             })}
                           </div>
                         </div>
                    </form>
                </div>
            </section>
            
            <section aria-labelledby="daily-target-settings-title">
                <h2
                    id="daily-target-settings-title"
                    className="section-title collapsible-header"
                    onClick={() => toggleSection('dailyTargetSettings')}
                    onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSection('dailyTargetSettings'); } }}
                    role="button"
                    tabIndex={0}
                    aria-expanded={sectionVisibility.dailyTargetSettings}
                    aria-controls="daily-target-settings-content"
                >
                    <span>日次目標設定</span>
                     <span className={`toggle-icon ${sectionVisibility.dailyTargetSettings ? 'open' : ''}`}>▼</span>
                </h2>
                <div id="daily-target-settings-content" className={`collapsible-content ${sectionVisibility.dailyTargetSettings ? 'open' : ''}`}>
                    <form className="modal-body" style={{padding:0}}>
                         <div className="media-kpi-section">
                           <h3 className="sub-section-title">媒体別実績 日次目標</h3>
                           <div className="media-kpi-grid">
                             {activeMedia.map(source => {
                                 const sourceKey = source.id;
                                 const fields: {key: KpiKey, label: string}[] = [
                                     {key: `${sourceKey}_scoutsSent`, label: 'スカウト数'},
                                     {key: `${sourceKey}_scoutReplies`, label: 'スカウト返信数'},
                                     {key: `${sourceKey}_effectiveReplies`, label: '有効返信数'},
                                 ];
                                return (
                                 <fieldset key={source.id} className="media-fieldset">
                                     <legend>{source.name}</legend>
                                     <div className="inputs-wrapper">
                                     {fields.map(field => (
                                          <div key={field.key} className="form-group">
                                             <label htmlFor={`daily-target-${field.key}`}>{field.label}</label>
                                             <input
                                                 type="number"
                                                 id={`daily-target-${field.key}`}
                                                 name={field.key}
                                                 value={dailyKpiTargets[field.key] || ''}
                                                 onChange={handleDailyTargetChange}
                                                 min="0"
                                                 placeholder="0"
                                             />
                                         </div>
                                     ))}
                                     </div>
                                 </fieldset>
                                 )
                             })}
                           </div>
                         </div>
                    </form>
                </div>
            </section>
          </>
        )}
        {view === 'all_users_kpi' && (
          isLoadingAllUsers ? (
            <div className="loading-container">チームメンバーのデータをGoogleドライブから読み込み中...</div>
          ) : (
            <>
              <div className="comparison-user-selector">
                <div className="comparison-user-selector-header">
                  <span>比較するユーザーを選択（未選択の場合は全員を表示）</span>
                  <div style={{ display: 'flex', gap: '0.5rem' }}>
                    <button onClick={() => setComparisonUserEmails(divisionScopedUserOptions.map(u => u.email))} className="secondary-action-button">全て選択</button>
                    <button onClick={() => setComparisonUserEmails([])} className="secondary-action-button">選択をクリア</button>
                  </div>
                </div>
                <div className="comparison-team-groups">
                  {comparisonTeamGroups.groups.map(group => {
                    const emails = group.members.map(m => m.email);
                    const selectedCount = emails.filter(e => comparisonUserEmails.includes(e)).length;
                    const allSelected = selectedCount === emails.length;
                    const isExpanded = !!expandedComparisonGroups[group.id];
                    return (
                      <div key={group.id} className="comparison-team-group">
                        <div className="comparison-team-group-header">
                          <button
                            type="button"
                            className="comparison-team-toggle"
                            onClick={() => toggleComparisonGroupExpanded(group.id)}
                            aria-expanded={isExpanded}
                            aria-label={isExpanded ? 'メンバーを折りたたむ' : 'メンバーを表示'}
                          >
                            <span className={`toggle-icon ${isExpanded ? 'open' : ''}`}>▶</span>
                          </button>
                          <label className="comparison-user-checkbox comparison-team-checkbox">
                            <input
                              type="checkbox"
                              checked={allSelected}
                              onChange={() => setComparisonGroupSelected(emails, !allSelected)}
                            />
                            {group.name}
                            <small style={{ color: 'var(--text-muted-color)' }}> ({selectedCount}/{emails.length})</small>
                          </label>
                        </div>
                        {isExpanded && (
                          <div className="comparison-user-checkbox-list nested">
                            {group.members.map(u => (
                              <label key={u.email} className="comparison-user-checkbox">
                                <input
                                  type="checkbox"
                                  checked={comparisonUserEmails.includes(u.email)}
                                  onChange={() => toggleComparisonUser(u.email)}
                                />
                                {u.label}
                              </label>
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })}
                  {comparisonTeamGroups.unassigned.length > 0 && (() => {
                    const emails = comparisonTeamGroups.unassigned.map(u => u.email);
                    const selectedCount = emails.filter(e => comparisonUserEmails.includes(e)).length;
                    const allSelected = selectedCount === emails.length;
                    const isExpanded = !!expandedComparisonGroups['__unassigned__'];
                    return (
                      <div className="comparison-team-group">
                        <div className="comparison-team-group-header">
                          <button
                            type="button"
                            className="comparison-team-toggle"
                            onClick={() => toggleComparisonGroupExpanded('__unassigned__')}
                            aria-expanded={isExpanded}
                            aria-label={isExpanded ? 'メンバーを折りたたむ' : 'メンバーを表示'}
                          >
                            <span className={`toggle-icon ${isExpanded ? 'open' : ''}`}>▶</span>
                          </button>
                          <label className="comparison-user-checkbox comparison-team-checkbox">
                            <input
                              type="checkbox"
                              checked={allSelected}
                              onChange={() => setComparisonGroupSelected(emails, !allSelected)}
                            />
                            未所属
                            <small style={{ color: 'var(--text-muted-color)' }}> ({selectedCount}/{emails.length})</small>
                          </label>
                        </div>
                        {isExpanded && (
                          <div className="comparison-user-checkbox-list nested">
                            {comparisonTeamGroups.unassigned.map(u => (
                              <label key={u.email} className="comparison-user-checkbox">
                                <input
                                  type="checkbox"
                                  checked={comparisonUserEmails.includes(u.email)}
                                  onChange={() => toggleComparisonUser(u.email)}
                                />
                                {u.label}
                              </label>
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })()}
                </div>
              </div>
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.5rem', marginBottom: '0.5rem' }}>
                <button
                  onClick={() => {
                    const label = comparisonUserEmails.length > 0 ? `選択ユーザー${comparisonUsers.length}名` : '全ユーザー';
                    handleExportAllUsersProgress(label, comparisonUsers);
                  }}
                  className="export-button"
                >
                  CSV出力
                </button>
                <button onClick={() => fetchAllUsersData()}>更新</button>
              </div>
              <div className="custom-period-export-bar">
                <span>表示・出力期間（未入力の場合は今月）:</span>
                <button onClick={() => handleShiftDashboardMonth(-1)} className="secondary-action-button month-shift-button">&lt; 前月</button>
                <input type="date" value={customExportStartDate} onChange={(e) => setCustomExportStartDate(e.target.value)} aria-label="開始日" />
                <span>〜</span>
                <input type="date" value={customExportEndDate} onChange={(e) => setCustomExportEndDate(e.target.value)} aria-label="終了日" />
                <button onClick={() => handleShiftDashboardMonth(1)} className="secondary-action-button month-shift-button">次月 &gt;</button>
                <button onClick={handleTogglePeriodFilter} className="secondary-action-button">
                  {dashboardPeriodOverride ? '今月表示に戻す' : '期間で絞り込みを有効にする'}
                </button>
              </div>
              <AllUsersDashboard
                  users={comparisonUsers}
                  allUsersData={displayedAllUsersData}
                  allMedia={allMedia}
                  dayOfWeekReplyRateData={dayOfWeekReplyRateData}
                  weekStartDate={viewWeekStartDate}
                  onPrevWeek={() => setViewWeekStartDate(d => new Date(d.setDate(d.getDate() - 7)))}
                  onNextWeek={() => setViewWeekStartDate(d => new Date(d.setDate(d.getDate() + 7)))}
                  visibility={{ progress: sectionVisibility.allUsersProgress, dowRate: sectionVisibility.allUsersDayOfWeekRate, weeklySummary: sectionVisibility.allUsersWeeklySummary, memberWeeklySummary: sectionVisibility.allUsersMemberWeeklySummary, grossProfit: sectionVisibility.allUsersGrossProfit, monthlyTrend: sectionVisibility.allUsersMonthlyTrend }}
                  toggleSection={toggleSection}
                  showGrossProfit={false}
                  periodOverride={dashboardPeriodOverride}
              />
            </>
          )
        )}
        {view === 'team_kpi' && (
          <div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end', marginBottom: '1.5rem' }}>
              <div className="form-group" style={{ maxWidth: '300px', marginBottom: 0 }}>
                <label htmlFor="team-select">チームを選択</label>
                <select
                  id="team-select"
                  value={selectedTeamId || ''}
                  onChange={(e) => setSelectedTeamId(e.target.value || null)}
                >
                  <option value="">選択してください</option>
                  {teams.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
                </select>
              </div>
              <div style={{ display: 'flex', gap: '0.5rem' }}>
                <button
                  onClick={() => {
                    const teamName = teams.find(t => t.id === selectedTeamId)?.name || 'チーム';
                    const teamUsers = selectedTeamMemberEmails.filter(email => displayedAllUsersData[email]);
                    handleExportTeamProgress(teamName, teamUsers);
                  }}
                  disabled={!selectedTeamId}
                  className="export-button"
                >
                  CSV出力
                </button>
                <button onClick={() => fetchAllUsersData()}>更新</button>
              </div>
            </div>
            <div className="custom-period-export-bar">
              <span>表示・出力期間（未入力の場合は今月）:</span>
              <button onClick={() => handleShiftDashboardMonth(-1)} className="secondary-action-button month-shift-button">&lt; 前月</button>
              <input type="date" value={customExportStartDate} onChange={(e) => setCustomExportStartDate(e.target.value)} aria-label="開始日" />
              <span>〜</span>
              <input type="date" value={customExportEndDate} onChange={(e) => setCustomExportEndDate(e.target.value)} aria-label="終了日" />
              <button onClick={() => handleShiftDashboardMonth(1)} className="secondary-action-button month-shift-button">次月 &gt;</button>
              <button onClick={handleTogglePeriodFilter} className="secondary-action-button">
                {dashboardPeriodOverride ? '今月表示に戻す' : '期間で絞り込みを有効にする'}
              </button>
            </div>
            {!selectedTeamId ? (
              <p className="no-data-message">チームを選択してください。チームがまだない場合は「チーム管理」から作成してください。</p>
            ) : isLoadingAllUsers ? (
              <div className="loading-container">チームメンバーのデータをGoogleドライブから読み込み中...</div>
            ) : (
              <AllUsersDashboard
                  users={selectedTeamMemberEmails.filter(email => displayedAllUsersData[email])}
                  allUsersData={displayedAllUsersData}
                  allMedia={allMedia}
                  dayOfWeekReplyRateData={dayOfWeekReplyRateData}
                  weekStartDate={viewWeekStartDate}
                  onPrevWeek={() => setViewWeekStartDate(d => new Date(d.setDate(d.getDate() - 7)))}
                  onNextWeek={() => setViewWeekStartDate(d => new Date(d.setDate(d.getDate() + 7)))}
                  visibility={{ progress: sectionVisibility.allUsersProgress, dowRate: sectionVisibility.allUsersDayOfWeekRate, weeklySummary: sectionVisibility.allUsersWeeklySummary, memberWeeklySummary: sectionVisibility.allUsersMemberWeeklySummary, grossProfit: sectionVisibility.allUsersGrossProfit, monthlyTrend: sectionVisibility.allUsersMonthlyTrend }}
                  toggleSection={toggleSection}
                  periodOverride={dashboardPeriodOverride}
              />
            )}
          </div>
        )}
        {view === 'pipeline' && (
          <>
            <div className="google-tasks-sync-bar">
              <span className="google-tasks-sync-label">選考予定はGoogleタスクに自動で同期されます</span>
              <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: '0.5rem' }}>
                <button
                  type="button"
                  onClick={handleSyncAllTasksNow}
                  disabled={tasksSyncStatus === 'loading'}
                  className="secondary-action-button"
                >
                  {tasksSyncStatus === 'loading' ? '同期中...' : '今すぐ同期'}
                </button>
                {tasksSyncStatus === 'needs-reauth' && (
                  <button type="button" onClick={handleReauthorizeTasks} className="secondary-action-button">
                    Googleタスクの利用を許可
                  </button>
                )}
                {tasksSyncMessage && (
                  <span className={`google-tasks-sync-message ${tasksSyncStatus === 'error' || tasksSyncStatus === 'needs-reauth' ? 'is-error' : ''}`}>
                    {tasksSyncMessage}
                  </span>
                )}
              </div>
            </div>
            <CandidatePipelineView
                candidates={pipelineCandidates}
                allMedia={allMedia}
                onSave={handleSaveCandidate}
                onToggleVisibility={handleToggleCandidateVisibility}
                currentUserEmail={currentIdentity?.email || ''}
                scope={pipelineScope}
                onScopeChange={setPipelineScope}
                teams={teams}
                selectedTeamId={pipelineSelectedTeamId}
                onSelectedTeamIdChange={setPipelineSelectedTeamId}
                userOptions={pipelineUserOptions}
                selectedUserEmail={pipelineSelectedUserEmail}
                onSelectedUserEmailChange={setPipelineSelectedUserEmail}
                isLoadingAggregate={isLoadingAllUsers}
            />
          </>
        )}
      </main>
    </div>
  );
};

const root = ReactDOM.createRoot(document.getElementById('root') as HTMLElement);
root.render(<React.StrictMode><App /></React.StrictMode>);