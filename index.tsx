
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
import { signIn, signOut, getCurrentSession, getLastKnownEmail, GoogleIdentity } from './services/googleAuth';
import { loadOwnData, saveOwnDataDebounced, flushPendingSave, forceSyncNow, hasPendingSync, retryPendingSyncIfNeeded, onSyncStatusChange, getLastSyncedAt, readLegacyAppData, loadAllTeammatesData, loadTeamsConfig, saveTeamsConfig, readLocalCache, loadMediaConfig, saveMediaConfig, readMediaConfigCache } from './services/dataSync';
import { searchInterviewLogsByName, exportGoogleDocAsText, InterviewLogFile } from './services/googleDrive';

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
const PIPELINE_STAGES = ['打診', '書類選考', 'カジュアル面談', '1次面接', '2次面接', '最終面接', '内定', '内定承諾', 'お見送り', '選考辞退'] as const;
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
}

interface Team {
  id: string;
  name: string;
  memberEmails: string[];
  createdBy: string;
  createdAt: string;
}

interface TeamsConfig {
  schemaVersion: number;
  teams: Team[];
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

const MonthOverMonthPerformanceChart: React.FC<{ entries: KpiEntry[]; allMedia: MediaEntry[] }> = ({ entries, allMedia }) => {
    const { chartData, maxReplyRate } = useMemo(() => {
        const monthlyData: Record<string, KpiTotals> = {};
        const allKeys = buildAllKpiKeys(allMedia);

        entries.forEach(entry => {
            const month = entry.date.substring(0, 7); // YYYY-MM
            if (!monthlyData[month]) {
                monthlyData[month] = allKeys.reduce((acc, key) => {
                  acc[key] = 0;
                  return acc;
                }, {} as KpiTotals);
            }

            (Object.keys(entry.values) as KpiKey[]).forEach(key => {
                monthlyData[month][key] += entry.values[key] || 0;
            });
        });

        const sortedMonths = Object.keys(monthlyData).sort();

        const labels = sortedMonths;
        const scoutRepliesData = sortedMonths.map(month => getTotalFromLump(monthlyData[month], '_scoutReplies', allMedia));
        const docScreeningPassedData = sortedMonths.map(month => monthlyData[month].documentScreeningPassed || 0);
        const firstInterviewPassedData = sortedMonths.map(month => monthlyData[month].firstInterviewPassed || 0);
        const secondInterviewPassedData = sortedMonths.map(month => monthlyData[month].secondInterviewPassed || 0);
        const offersExtendedData = sortedMonths.map(month => monthlyData[month].offersExtended || 0);
        const placementsData = sortedMonths.map(month => monthlyData[month].placements || 0);
        const replyRateData = sortedMonths.map(month => {
            const scoutsSent = getTotalFromLump(monthlyData[month], '_scoutsSent', allMedia);
            const scoutReplies = getTotalFromLump(monthlyData[month], '_scoutReplies', allMedia);
            return scoutsSent > 0 ? (scoutReplies / scoutsSent) * 100 : 0;
        });

        const calculateAverage = (data: number[]) => data.length > 0 ? data.reduce((a, b) => a + b, 0) / data.length : 0;
        const avgReplyRate = calculateAverage(replyRateData);
        const avgScoutReplies = calculateAverage(scoutRepliesData);
        const avgPlacements = calculateAverage(placementsData);
        
        const currentMaxReplyRate = Math.max(...replyRateData, 0);

        const data = {
            labels,
            datasets: [
                {
                    label: '返信率',
                    data: replyRateData,
                    borderColor: 'rgb(255, 99, 132)',
                    yAxisID: 'y-axis-rate',
                    tension: 0.2,
                    fill: false,
                },
                {
                    label: 'スカウト返信数',
                    data: scoutRepliesData,
                    borderColor: '#a9def9',
                    yAxisID: 'y-axis-count',
                    tension: 0.2,
                    fill: false,
                },
                {
                    label: '書類選考通過数',
                    data: docScreeningPassedData,
                    borderColor: '#72b6e8',
                    yAxisID: 'y-axis-count',
                    tension: 0.2,
                    fill: false,
                },
                {
                    label: '1次面接通過数',
                    data: firstInterviewPassedData,
                    borderColor: '#3c8abe',
                    yAxisID: 'y-axis-count',
                    tension: 0.2,
                    fill: false,
                },
                {
                    label: '2次面接通過数',
                    data: secondInterviewPassedData,
                    borderColor: '#1e5f94',
                    yAxisID: 'y-axis-count',
                    tension: 0.2,
                    fill: false,
                },
                {
                    label: '内定数',
                    data: offersExtendedData,
                    borderColor: '#0b3a61',
                    yAxisID: 'y-axis-count',
                    tension: 0.2,
                    fill: false,
                },
                {
                    label: '内定承諾数',
                    data: placementsData,
                    borderColor: '#28a745',
                    yAxisID: 'y-axis-count',
                    tension: 0.2,
                    fill: false,
                },
                {
                    label: '返信率 (平均)',
                    data: new Array(labels.length).fill(avgReplyRate),
                    borderColor: 'rgba(255, 99, 132, 0.5)',
                    yAxisID: 'y-axis-rate',
                    borderDash: [5, 5],
                    pointRadius: 0,
                    fill: false,
                    tension: 0,
                    borderWidth: 2,
                },
                {
                    label: 'スカウト返信数 (平均)',
                    data: new Array(labels.length).fill(avgScoutReplies),
                    borderColor: 'rgba(169, 222, 249, 0.8)',
                    yAxisID: 'y-axis-count',
                    borderDash: [5, 5],
                    pointRadius: 0,
                    fill: false,
                    tension: 0,
                    borderWidth: 2,
                },
                {
                    label: '内定承諾数 (平均)',
                    data: new Array(labels.length).fill(avgPlacements),
                    borderColor: 'rgba(40, 167, 69, 0.7)',
                    yAxisID: 'y-axis-count',
                    borderDash: [5, 5],
                    pointRadius: 0,
                    fill: false,
                    tension: 0,
                    borderWidth: 2,
                },
            ],
        };
        return { chartData: data, maxReplyRate: currentMaxReplyRate };
    }, [entries, allMedia]);

    const options = useMemo(() => ({
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
            mode: 'index' as const,
            intersect: false,
        },
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
                    drawOnChartArea: false,
                },
                beginAtZero: true,
                max: maxReplyRate > 0 ? Math.ceil(maxReplyRate * 1.2) : 10,
            },
            x: {
               grid: {
                  display: false
               }
            }
        },
    }), [maxReplyRate]);

    if (entries.length === 0) {
        return <p className="no-data-message">実績データがありません。</p>;
    }

    return (
        <div className="chart-container">
            <Line options={options} data={chartData as any} />
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
        <div className="weekly-summary-totals">
          <h4>週間進捗</h4>
          <div className="total-item">
            <span>候補者推薦数:</span>
            <strong>{data.totalCandidatesSubmitted}</strong>
          </div>
          <div className="total-item">
            <span>初回面談数:</span>
            <strong>{data.totalInitialInterviews}</strong>
          </div>
        </div>
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
  onClose: () => void;
}> = ({ date, initialValues, activeMedia, onSave, onClose }) => {
  const [entryValues, setEntryValues] = useState<{ [key in KpiKey]?: number }>(
    initialValues || {}
  );

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

  const isFormEmpty = Object.values(entryValues).every(val => val === undefined || val === 0 || val === null);
  const isSaveDisabled = isFormEmpty && (!initialValues || Object.values(initialValues).every(v => v === 0));
  const canClear = initialValues && Object.values(initialValues).some(v => v > 0);

  return (
    <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="modal-title">
      <div className="modal-content" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h3 id="modal-title">{formattedDate} の実績入力</h3>
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
            <div className="media-kpi-grid">
              {activeMedia.map(source => {
                const sourceKey = source.id;
                const scoutsKey = `${sourceKey}_scoutsSent` as KpiKey;
                const repliesKey = `${sourceKey}_scoutReplies` as KpiKey;
                const effectiveRepliesKey = `${sourceKey}_effectiveReplies` as KpiKey;
                const documentsCollectedKey = `${sourceKey}_documentsCollected` as KpiKey;
                const effectiveDocumentsCollectedKey = `${sourceKey}_effectiveDocumentsCollected` as KpiKey;
                const interviewsKey = `${sourceKey}_initialInterviews` as KpiKey;
                const effectiveInterviewsKey = `${sourceKey}_effectiveInitialInterviews` as KpiKey;
                return (
                  <fieldset key={source.id} className="media-fieldset">
                    <legend>{source.name}</legend>
                    <div className="inputs-wrapper">
                      <div className="form-group">
                        <label htmlFor={`modal-${scoutsKey}`}>スカウト数</label>
                        <input
                          type="number"
                          id={`modal-${scoutsKey}`}
                          name={scoutsKey}
                          value={entryValues[scoutsKey] ?? ''}
                          onChange={handleInputChange}
                          min="0"
                          placeholder="0"
                          aria-label={`${source.name} スカウト数を入力`}
                        />
                      </div>
                      <div className="form-group">
                        <label htmlFor={`modal-${repliesKey}`}>スカウト返信数</label>
                        <input
                          type="number"
                          id={`modal-${repliesKey}`}
                          name={repliesKey}
                          value={entryValues[repliesKey] ?? ''}
                          onChange={handleInputChange}
                          min="0"
                          placeholder="0"
                          aria-label={`${source.name} スカウト返信数を入力`}
                        />
                      </div>
                      <div className="form-group">
                        <label htmlFor={`modal-${effectiveRepliesKey}`}>有効返信数</label>
                        <input
                          type="number"
                          id={`modal-${effectiveRepliesKey}`}
                          name={effectiveRepliesKey}
                          value={entryValues[effectiveRepliesKey] ?? ''}
                          onChange={handleInputChange}
                          min="0"
                          placeholder="0"
                          aria-label={`${source.name} 有効返信数を入力`}
                        />
                      </div>
                      <div className="form-group">
                        <label htmlFor={`modal-${documentsCollectedKey}`}>書類回収数</label>
                        <input
                          type="number"
                          id={`modal-${documentsCollectedKey}`}
                          name={documentsCollectedKey}
                          value={entryValues[documentsCollectedKey] ?? ''}
                          onChange={handleInputChange}
                          min="0"
                          placeholder="0"
                          aria-label={`${source.name} 書類回収数を入力`}
                        />
                      </div>
                      <div className="form-group">
                        <label htmlFor={`modal-${effectiveDocumentsCollectedKey}`}>有効書類回収数</label>
                        <input
                          type="number"
                          id={`modal-${effectiveDocumentsCollectedKey}`}
                          name={effectiveDocumentsCollectedKey}
                          value={entryValues[effectiveDocumentsCollectedKey] ?? ''}
                          onChange={handleInputChange}
                          min="0"
                          placeholder="0"
                          aria-label={`${source.name} 有効書類回収数を入力`}
                        />
                      </div>
                       <div className="form-group">
                        <label htmlFor={`modal-${interviewsKey}`}>初回面談数</label>
                        <input
                          type="number"
                          id={`modal-${interviewsKey}`}
                          name={interviewsKey}
                          value={entryValues[interviewsKey] ?? ''}
                          onChange={handleInputChange}
                          min="0"
                          placeholder="0"
                          aria-label={`${source.name} 初回面談数を入力`}
                        />
                      </div>
                      <div className="form-group">
                        <label htmlFor={`modal-${effectiveInterviewsKey}`}>初回有効面談数</label>
                        <input
                          type="number"
                          id={`modal-${effectiveInterviewsKey}`}
                          name={effectiveInterviewsKey}
                          value={entryValues[effectiveInterviewsKey] ?? ''}
                          onChange={handleInputChange}
                          min="0"
                          placeholder="0"
                          aria-label={`${source.name} 初回有効面談数を入力`}
                        />
                      </div>
                    </div>
                  </fieldset>
                )
              })}
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
    ownerEmail: string | null;
    userOptions: { email: string; label: string }[];
    onClose: () => void;
    onCreateTeam: (name: string) => void;
    onRenameTeam: (teamId: string, name: string) => void;
    onDeleteTeam: (teamId: string) => void;
    onAddMember: (teamId: string, email: string) => void;
    onRemoveMember: (teamId: string, email: string) => void;
}> = ({ teams, isEditable, ownerEmail, userOptions, onClose, onCreateTeam, onRenameTeam, onDeleteTeam, onAddMember, onRemoveMember }) => {
    const [newTeamName, setNewTeamName] = useState('');
    const [editingTeamId, setEditingTeamId] = useState<string | null>(null);
    const [editedName, setEditedName] = useState('');
    const [memberInputs, setMemberInputs] = useState<Record<string, string>>({});

    const labelByEmail = useMemo(() => new Map(userOptions.map(u => [u.email, u.label])), [userOptions]);

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
        const email = (emailOverride ?? memberInputs[teamId] ?? '').trim();
        if (!email) return;
        if (!email.toLowerCase().endsWith('@bloom-firm.com')) {
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
                    {!isEditable && (
                        <p className="no-data-message">
                            チーム設定の編集は作成者（{ownerEmail || '不明'}）のみ可能です。閲覧のみできます。
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
    const resumeInputRef = useRef<HTMLInputElement>(null);
    const audioInputRef = useRef<HTMLInputElement>(null);

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

    const handleUseInterviewLog = async (file: InterviewLogFile) => {
        setIsSummarizingInterviewLog(true);
        try {
            const text = await exportGoogleDocAsText(file.id);
            const ai = new GoogleGenAI({ apiKey: process.env.API_KEY as string });
            const response = await ai.models.generateContent({
                model: 'gemini-2.5-flash',
                contents: `以下はGoogle Meetの面談議事録です。この内容をシンプルに要約してください。\n\n---\n${text}`,
            });
            const dateLabel = new Date(file.modifiedTime).toLocaleDateString('ja-JP');
            const entry = `--- 面談ログ「${file.name}」(${dateLabel}) より ---\n${response.text.trim()}`;
            setCandidate(prev => ({
                ...prev,
                interviewSummary: prev.interviewSummary ? `${prev.interviewSummary}\n\n${entry}` : entry,
            }));
            setInterviewLogResults(null);
        } catch (error) {
            console.error('Error summarizing interview log:', error);
            alert('面談ログの要約生成中にエラーが発生しました。');
        } finally {
            setIsSummarizingInterviewLog(false);
        }
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
                          <input
                            id={`scheduledDate-${app.id}`}
                            type="date"
                            value={app.scheduledDate || ''}
                            onChange={e => handleApplicationChange(index, 'scheduledDate', e.target.value)}
                            aria-label={`選考予定日 ${index + 1}`}
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

              { (resumeDragActive || audioDragActive) && <div id="drag-file-element" onDragEnter={handleResumeDrag} onDragLeave={handleResumeDrag} onDragOver={handleResumeDrag} onDrop={handleResumeDrop}></div>}
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
                        <input
                            type="date"
                            id="scheduledDate"
                            name="scheduledDate"
                            value={application.scheduledDate || ''}
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
        'カジュアル面談': 3,
        '1次面接': 4,
        '2次面接': 5,
        '最終面接': 6,
        '内定': 7,
        '内定承諾': 8,
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
        PIPELINE_STAGES.map(stage => [stage, { stage, count: 0, estimableCount: 0, revenue: 0, cost: 0, profit: 0 }])
    );

    pickBestApplicationPerCandidate(candidates).forEach(({ candidate, application }) => {
        const bucket = totalsByStage.get(application.stage)!;
        bucket.count++;
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
    const stageTotals = useMemo(() => computeGrossProfitByStage(candidates, allMedia), [candidates, allMedia]);

    const grandTotal = useMemo(() => {
        return stageTotals
            .filter(s => s.stage !== 'お見送り' && s.stage !== '選考辞退')
            .reduce((acc, s) => ({
                count: acc.count + s.count,
                estimableCount: acc.estimableCount + s.estimableCount,
                revenue: acc.revenue + s.revenue,
                cost: acc.cost + s.cost,
                profit: acc.profit + s.profit,
            }), { count: 0, estimableCount: 0, revenue: 0, cost: 0, profit: 0 });
    }, [stageTotals]);

    return (
        <div className="gross-profit-summary">
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
                    ※ お見送り・選考辞退を除く{grandTotal.count}件中、想定年収とfee料率が両方入力済みの{grandTotal.estimableCount}件のみを集計しています（残り{grandTotal.count - grandTotal.estimableCount}件は未入力のため対象外）。
                </p>
            )}
            <div className="detail-application-grid">
                {stageTotals.filter(s => s.count > 0).map(s => (
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
                    </div>
                ))}
            </div>
        </div>
    );
};


interface PipelineCalendarEvent {
    candidateName: string;
    companyName: string;
    stage: PipelineStage;
    ownerLabel?: string;
}

/**
 * Month calendar of every visible application's manually-set scheduledDate. Reads whatever
 * `candidates` list the caller passes in, so it's automatically scoped by the pipeline's
 * existing 自分/全ユーザー/チーム/ユーザー別 switcher — no separate scope control needed here.
 * Clicking a day lets the signed-in user schedule (or edit) one of their own candidates'
 * applications for that date, via onDayClick.
 */
const PipelineCalendarView: React.FC<{
    candidates: Candidate[];
    viewDate: Date;
    onPrevMonth: () => void;
    onNextMonth: () => void;
    onDayClick: (dateStr: string) => void;
}> = ({ candidates, viewDate, onPrevMonth, onNextMonth, onDayClick }) => {
    const year = viewDate.getFullYear();
    const month = viewDate.getMonth();

    const eventsByDate = useMemo(() => {
        const map = new Map<string, PipelineCalendarEvent[]>();
        candidates.filter(c => !c.isHidden).forEach(c => {
            c.applications.filter(app => !app.isHidden && app.scheduledDate).forEach(app => {
                const list = map.get(app.scheduledDate!) || [];
                list.push({ candidateName: c.name, companyName: app.companyName, stage: app.stage, ownerLabel: c.ownerLabel });
                map.set(app.scheduledDate!, list);
            });
        });
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
                        {events.map((ev, idx) => (
                            <div
                                key={idx}
                                className="pipeline-calendar-event"
                                style={{ '--badge-color': STAGE_COLOR_MAP[ev.stage] } as React.CSSProperties}
                                title={`${ev.candidateName} / ${ev.companyName} / ${ev.stage}${ev.ownerLabel ? ` (${ev.ownerLabel})` : ''}`}
                            >
                                <span className="pipeline-calendar-event-stage">{STAGE_SHORT_LABELS[ev.stage]}</span>
                                {ev.candidateName} - {ev.companyName}
                            </div>
                        ))}
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
    const [showHidden, setShowHidden] = useState(false);
    const [selectedStageFilters, setSelectedStageFilters] = useState<PipelineStage[]>([]);
    const [sortConfig, setSortConfig] = useState<{ key: keyof Candidate; direction: 'asc' | 'desc' } | null>({ key: 'createdAt', direction: 'desc'});
    const [expandedCandidateId, setExpandedCandidateId] = useState<string | null>(null);
    const [isReportVisible, setIsReportVisible] = useState(true);
    const [isCalendarVisible, setIsCalendarVisible] = useState(true);
    const [isCompanyPipelineVisible, setIsCompanyPipelineVisible] = useState(true);
    const [isGrossProfitVisible, setIsGrossProfitVisible] = useState(true);
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

    const handleExportCSV = () => {
        const dataToExport = candidates.filter(c => !c.isHidden);
        
        if (dataToExport.length === 0) {
            alert('エクスポート対象の候補者がいません。');
            return;
        }

        const headers = [
            '氏名', '担当者', '現職企業名', '最終学歴', '現年収(万円)', '希望年収(万円)',
            '集客媒体', '他エージェント使用状況', '登録日', '概要',
            '応募企業名', '進捗状況', '次アクション', '内定確度', '入社確度'
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
                escapeCSV(candidate.source),
                candidate.usingOtherAgents ? 'あり' : 'なし',
                new Date(candidate.createdAt).toLocaleDateString('ja-JP'),
                escapeCSV(candidate.summary),
            ];

            const emptyCommonData = Array(commonData.length).fill('');
            const visibleApps = candidate.applications.filter(app => !app.isHidden);

            if (visibleApps.length > 0) {
                visibleApps.forEach((app, index) => {
                    const applicationData = [
                        escapeCSV(app.companyName),
                        escapeCSV(app.stage),
                        escapeCSV(app.nextAction),
                        escapeCSV(app.offerConfidence || ''),
                        escapeCSV(app.acceptanceConfidence || ''),
                    ];
                    if (index === 0) {
                        rows.push([...commonData, ...applicationData]);
                    } else {
                        rows.push([...emptyCommonData, ...applicationData]);
                    }
                });
            } else {
                 rows.push([...commonData, '', '', '', '', '']);
            }
        });
        
        // --- Pipeline Summary Calculation ---
        const STAGE_WEIGHTS: Record<PipelineStage, number> = {
            '打診': 1, '書類選考': 2, 'カジュアル面談': 3, '1次面接': 4, '2次面接': 5,
            '最終面接': 6, '内定': 7, '内定承諾': 8, 'お見送り': 0, '選考辞退': 0,
        };
        const stageCounts = PIPELINE_STAGES.reduce((acc, stage) => {
            acc[stage] = 0;
            return acc;
        }, {} as Record<PipelineStage, number>);

        dataToExport.forEach(candidate => {
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
            if (mostAdvancedStage) {
                stageCounts[mostAdvancedStage]++;
            }
        });

        const summaryRows: string[] = [];
        summaryRows.push('');
        summaryRows.push('パイプラインサマリー');
        summaryRows.push('ステージ,候補者数');
        PIPELINE_STAGES.forEach(stage => {
            summaryRows.push(`${escapeCSV(stage)},${stageCounts[stage]}`);
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
            const matchesVisibility = showHidden ? c.isHidden === true : !c.isHidden;
            const matchesStage = selectedStageFilters.length === 0 || c.applications.some(
                app => !app.isHidden && selectedStageFilters.includes(app.stage)
            );
            return matchesSearch && matchesVisibility && matchesStage;
        });
    }, [candidates, searchTerm, showHidden, selectedStageFilters]);

    const sortedCandidates = useMemo(() => {
        let sortableItems = [...filteredCandidates];
        if (sortConfig !== null) {
            sortableItems.sort((a, b) => {
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

    const requestSort = (key: keyof Candidate) => {
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


    const getSortIndicator = (key: keyof Candidate) => {
        if (!sortConfig || sortConfig.key !== key) return '';
        return sortConfig.direction === 'asc' ? ' ▲' : ' ▼';
    };

    const sortOptions: { key: keyof Candidate, label: string }[] = [
      { key: 'createdAt', label: '登録日' },
      { key: 'name', label: '氏名' },
      { key: 'currentCompany', label: '現職企業名' },
      { key: 'currentSalary', label: '現職年収' },
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
                        onPrevMonth={() => setCalendarViewDate(d => new Date(d.getFullYear(), d.getMonth() - 1, 1))}
                        onNextMonth={() => setCalendarViewDate(d => new Date(d.getFullYear(), d.getMonth() + 1, 1))}
                        onDayClick={(dateStr) => setScheduleModalDate(dateStr)}
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
                 <div className="single-checkbox">
                    <input
                        type="checkbox"
                        id="show-hidden-candidates"
                        checked={showHidden}
                        onChange={e => setShowHidden(e.target.checked)}
                    />
                    <label htmlFor="show-hidden-candidates">非表示の候補者を表示</label>
                </div>
            </div>

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
                                            <button onClick={() => handleOpenApplicationModal(c, null)} className="add-selection-button">+ 選考追加</button>
                                            <button onClick={() => handleEdit(c)} className="edit-user-button">編集</button>
                                            <button onClick={() => onToggleVisibility(c.id)} className={showHidden ? "secondary-action-button" : "delete-user-button"}>
                                                {showHidden ? '再表示' : '非表示'}
                                            </button>
                                        </div>
                                    )}
                                </div>
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
                                                              <span>{app.scheduledDate ? new Date(app.scheduledDate + 'T00:00:00').toLocaleDateString('ja-JP') : '未設定'}</span>
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

// The pipeline funnel starts from 候補者推薦数 onward (GENERAL_KPIS); 返信数/面談数 are earlier,
// media-scoped sourcing-side steps (aggregated across all media via getTotalFromLump) that
// happen before a candidate is ever submitted to a client.
const FUNNEL_STAGES: FunnelStageDef[] = [
  { key: 'scoutReplies', label: '返信数', getValue: (totals, allMedia) => getTotalFromLump(totals, '_scoutReplies', allMedia) },
  { key: 'initialInterviews', label: '面談数', getValue: (totals, allMedia) => getTotalFromLump(totals, '_initialInterviews', allMedia) },
  ...(Object.keys(GENERAL_KPIS) as Array<keyof typeof GENERAL_KPIS>).map(key => ({
    key: key as string,
    label: GENERAL_KPIS[key].label,
    getValue: (totals: KpiTotals) => totals[key] || 0,
  })),
];

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
}> = ({ users, allUsersData, allMedia }) => {
  const [isVisible, setIsVisible] = useState(true);
  const [aiSuggestion, setAiSuggestion] = useState('');
  const [isGeneratingSuggestion, setIsGeneratingSuggestion] = useState(false);

  const { totalStageValues, perUserStageValues } = useMemo(() => {
    const perUser: Record<string, number[]> = {};
    const totals = FUNNEL_STAGES.map(() => 0);
    users.forEach(email => {
      const data = allUsersData[email];
      if (!data) return;
      const monthlyTotals = calculateMonthlyTotals(data.entries || [], allMedia);
      const values = FUNNEL_STAGES.map(stage => stage.getValue(monthlyTotals, allMedia));
      perUser[email] = values;
      values.forEach((v, i) => { totals[i] += v; });
    });
    return { totalStageValues: totals, perUserStageValues: perUser };
  }, [users, allUsersData, allMedia]);

  const totalConversionRates = useMemo(() => computeConversionRates(totalStageValues), [totalStageValues]);
  const bottleneckIndex = useMemo(() => findBottleneckIndex(totalConversionRates), [totalConversionRates]);

  const handleGenerateSuggestion = async () => {
    setIsGeneratingSuggestion(true);
    setAiSuggestion('');
    try {
      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY as string });
      const lines: string[] = ['【全体ファネル実績（今月・全ユーザー合計）】'];
      FUNNEL_STAGES.forEach((stage, i) => {
        const rate = totalConversionRates[i];
        lines.push(`${stage.label}: ${totalStageValues[i]}件${rate !== null ? `（前段階からの歩留まり: ${rate.toFixed(1)}%）` : ''}`);
      });
      lines.push('', '【ユーザー別実績（今月）】');
      users.forEach(email => {
        const data = allUsersData[email];
        if (!data) return;
        const label = data.displayName || email;
        const values = perUserStageValues[email] || FUNNEL_STAGES.map(() => 0);
        const rates = computeConversionRates(values);
        const parts = FUNNEL_STAGES.map((stage, i) =>
          `${stage.label} ${values[i]}件${rates[i] !== null ? `(${rates[i]!.toFixed(1)}%)` : ''}`
        );
        lines.push(`${label}: ${parts.join(', ')}`);
      });

      const prompt = `以下は採用エージェントの求人紹介パイプラインにおける、今月の全体ファネル実績とユーザー別実績です。各ステップの歩留まり（前段階からの通過率）を踏まえて、ボトルネックとなっている工程を指摘し、改善のための具体的な施策を日本語で3〜5点、簡潔に提案してください。\n\n${lines.join('\n')}`;

      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: prompt,
      });
      setAiSuggestion(response.text.trim());
    } catch (error) {
      console.error('Error generating improvement suggestion:', error);
      alert('AIによる改善提案の生成中にエラーが発生しました。');
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
        <h3 className="sub-section-title">全体ファネル（全ユーザー合計・今月）</h3>
        <div className="all-users-table-container">
          <table className="all-users-table">
            <thead>
              <tr><th>指標</th><th>件数</th><th>前段階からの歩留まり</th></tr>
            </thead>
            <tbody>
              {FUNNEL_STAGES.map((stage, i) => (
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

        <h3 className="sub-section-title" style={{ marginTop: '1.5rem' }}>ユーザー別ファネル比較（今月の歩留まり %）</h3>
        <div className="all-users-table-container">
          <table className="all-users-table">
            <thead>
              <tr>
                <th>ユーザー</th>
                {FUNNEL_STAGES.slice(1).map(stage => <th key={stage.key}>{stage.label}</th>)}
              </tr>
            </thead>
            <tbody>
              {users.map(email => {
                const data = allUsersData[email];
                if (!data) return null;
                const label = data.displayName || email;
                const values = perUserStageValues[email] || FUNNEL_STAGES.map(() => 0);
                const rates = computeConversionRates(values);
                const userBottleneckIndex = findBottleneckIndex(rates);
                return (
                  <tr key={email}>
                    <td>{label}</td>
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

        <div style={{ marginTop: '1.5rem' }}>
          <button type="button" onClick={handleGenerateSuggestion} disabled={isGeneratingSuggestion} className="submit-button">
            {isGeneratingSuggestion ? 'AIが分析中...' : 'AIに改善提案をもらう'}
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

const AllUsersDashboard: React.FC<{
  users: string[];
  allUsersData: Record<string, UserData>;
  allMedia: MediaEntry[];
  dayOfWeekReplyRateData: any | null;
  weekStartDate: Date;
  onPrevWeek: () => void;
  onNextWeek: () => void;
  visibility: { progress: boolean; dowRate: boolean; weeklySummary: boolean; memberWeeklySummary: boolean; grossProfit: boolean };
  toggleSection: (key: 'allUsersProgress' | 'allUsersDayOfWeekRate' | 'allUsersWeeklySummary' | 'allUsersMemberWeeklySummary' | 'allUsersGrossProfit') => void;
  showGrossProfit?: boolean;
}> = ({ users, allUsersData, allMedia, dayOfWeekReplyRateData, weekStartDate, onPrevWeek, onNextWeek, visibility, toggleSection, showGrossProfit = true }) => {
  const activeMedia = useMemo(() => allMedia.filter(m => !m.isArchived), [allMedia]);
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
          <span>全ユーザーの月次進捗</span>
          <span className={`toggle-icon ${visibility.progress ? 'open' : ''}`}>▼</span>
        </h2>
        <div id="all-users-progress-content" className={`collapsible-content ${visibility.progress ? 'open' : ''}`}>
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
                {users.map(user => {
                  const userData = allUsersData[user];
                  const displayName = userData?.displayName || user;
                  if (!userData) {
                    return <tr key={user}><td colSpan={Object.keys(GENERAL_KPIS).length + 7}>{displayName}のデータがありません。</td></tr>;
                  }

                  const monthlyTotals = calculateMonthlyTotals(userData.entries || [], allMedia);
                  const kpiTargets = { ...buildDefaultKpiTargets(allMedia), ...(userData.kpiTargets || {}) };

                  const sent = getTotalFromLump(monthlyTotals, '_scoutsSent', allMedia);
                  const replies = getTotalFromLump(monthlyTotals, '_scoutReplies', allMedia);
                  const effectiveReplies = getTotalFromLump(monthlyTotals, '_effectiveReplies', allMedia);

                  const replyRate = sent > 0 ? (replies / sent) * 100 : 0;
                  const effectiveReplyRate = replies > 0 ? (effectiveReplies / replies) * 100 : 0;

                  // Actuals (numerator) include archived media — their historical performance
                  // still counts — but targets (denominator) must only sum activeMedia: the
                  // 月次目標設定 form only ever lets anyone edit targets for active media, so
                  // summing allMedia would silently add in stale/default target values for
                  // archived media nobody can see or edit, making this total never match what
                  // was actually set in the settings form.
                  const documentsCollected = getTotalFromLump(monthlyTotals, '_documentsCollected', allMedia);
                  const documentsCollectedTarget = getTotalFromLump(kpiTargets, '_documentsCollected', activeMedia);
                  const documentsCollectedProgress = documentsCollectedTarget > 0 ? Math.min((documentsCollected / documentsCollectedTarget) * 100, 100) : 0;

                  const effectiveDocumentsCollected = getTotalFromLump(monthlyTotals, '_effectiveDocumentsCollected', allMedia);
                  const effectiveDocumentsCollectedTarget = getTotalFromLump(kpiTargets, '_effectiveDocumentsCollected', activeMedia);
                  const effectiveDocumentsCollectedProgress = effectiveDocumentsCollectedTarget > 0 ? Math.min((effectiveDocumentsCollected / effectiveDocumentsCollectedTarget) * 100, 100) : 0;

                  const initialInterviews = getTotalFromLump(monthlyTotals, '_initialInterviews', allMedia);
                  const effectiveInitialInterviews = getTotalFromLump(monthlyTotals, '_effectiveInitialInterviews', allMedia);
                  const effectiveInterviewRate = initialInterviews > 0 ? (effectiveInitialInterviews / initialInterviews) * 100 : 0;

                  const initialInterviewsTarget = getTotalFromLump(kpiTargets, '_initialInterviews', activeMedia);
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
                          <span>{documentsCollected} / {documentsCollectedTarget}</span>
                          <div className="mini-progress-bar">
                              <div className="progress-bar-fill" style={{ width: `${documentsCollectedProgress}%` }}></div>
                          </div>
                      </td>
                       <td className="progress-cell">
                          <span>{effectiveDocumentsCollected} / {effectiveDocumentsCollectedTarget}</span>
                          <div className="mini-progress-bar">
                              <div className="progress-bar-fill" style={{ width: `${effectiveDocumentsCollectedProgress}%`, backgroundColor: 'var(--info-color)' }}></div>
                          </div>
                      </td>
                      <td className="progress-cell">
                          <span>{initialInterviews} / {initialInterviewsTarget}</span>
                          <div className="mini-progress-bar">
                              <div className="progress-bar-fill" style={{ width: `${initialInterviewsProgress}%` }}></div>
                          </div>
                      </td>
                      <td className="progress-cell">
                        <span>{effectiveInterviewRate.toFixed(1)}%</span>
                        <div className="mini-progress-bar">
                          <div className="progress-bar-fill" style={{ width: `${Math.min(effectiveInterviewRate, 100)}%`, backgroundColor: 'var(--info-color)' }}></div>
                        </div>
                        <small>({effectiveInitialInterviews}/{initialInterviews})</small>
                      </td>
                      {(Object.keys(GENERAL_KPIS) as Array<keyof typeof GENERAL_KPIS>).map(key => {
                          const value = monthlyTotals[key] || 0;
                          const target = kpiTargets[key] || 0;
                          const progress = target > 0 ? Math.min((value / target) * 100, 100) : 0;
                          return (
                              <td key={key} className="progress-cell">
                                  <span>{value} / {target}</span>
                                  <div className="mini-progress-bar">
                                      <div className="progress-bar-fill" style={{ width: `${progress}%` }}></div>
                                  </div>
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

      <FunnelAnalysisSection users={users} allUsersData={allUsersData} allMedia={allMedia} />

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
            <span>全ユーザー 曜日別 累積返信率</span>
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
  | 'allUsersProgress' | 'allUsersDayOfWeekRate' | 'allUsersWeeklySummary' | 'allUsersMemberWeeklySummary' | 'allUsersGrossProfit';


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

  // Teams state
  const [teams, setTeams] = useState<Team[]>([]);
  const [teamsDriveFileId, setTeamsDriveFileId] = useState<string | null>(null);
  const [teamsOwnerEmail, setTeamsOwnerEmail] = useState<string | null>(null);
  const [isTeamsModalOpen, setIsTeamsModalOpen] = useState(false);
  const [selectedTeamId, setSelectedTeamId] = useState<string | null>(null);
  // Empty = no filter (show everyone) on the 全ユーザー tab; otherwise an ad-hoc selection of
  // specific users to compare, independent of the formal Team groupings.
  const [comparisonUserEmails, setComparisonUserEmails] = useState<string[]>([]);
  const [customExportStartDate, setCustomExportStartDate] = useState('');
  const [customExportEndDate, setCustomExportEndDate] = useState('');
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
    allUsersProgress: true,
    allUsersDayOfWeekRate: true,
    allUsersWeeklySummary: true,
    allUsersMemberWeeklySummary: true,
    allUsersGrossProfit: true,
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
      candidates: d.candidates || [],
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

  // Load the shared teams config when opening the team-filtered view, the pipeline's team
  // scope, or the Teams management modal.
  useEffect(() => {
    if (!currentIdentity || !isInitialized) return;
    const needsTeams = view === 'team_kpi' || isTeamsModalOpen || (view === 'pipeline' && pipelineScope === 'team');
    if (!needsTeams) return;
    let cancelled = false;
    (async () => {
      try {
        const result = await loadTeamsConfig<TeamsConfig>();
        if (cancelled) return;
        setTeams(result.data?.teams || []);
        setTeamsDriveFileId(result.driveFileId);
        setTeamsOwnerEmail(result.ownerEmail);
      } catch (error) {
        console.error('Failed to load teams config from Drive', error);
      }
    })();
    return () => { cancelled = true; };
  }, [currentIdentity, isInitialized, view, isTeamsModalOpen, pipelineScope]);

  const isTeamsEditable = !teamsOwnerEmail || teamsOwnerEmail === currentIdentity?.email;

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

  const persistTeams = async (updatedTeams: Team[]) => {
    setTeams(updatedTeams);
    if (!currentIdentity) return;
    try {
      const payload: TeamsConfig = { schemaVersion: 1, teams: updatedTeams };
      const newFileId = await saveTeamsConfig(teamsDriveFileId, payload, currentIdentity.email);
      setTeamsDriveFileId(newFileId);
      if (!teamsOwnerEmail) setTeamsOwnerEmail(currentIdentity.email);
    } catch (error) {
      console.error('Failed to save teams config', error);
      alert('チーム設定の保存に失敗しました。編集できるのは作成者のみです。');
    }
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

  const handleSaveEntry = (date: string, newValues: KpiTotals) => {
    setCurrentUserData(prevData => {
      if (!prevData) return null;
      const otherEntries = prevData.entries.filter(entry => entry.date !== date);
      const newEntry: KpiEntry = {
        id: Date.now(),
        date: date,
        values: newValues,
      };
      const updatedEntries = [...otherEntries, newEntry].sort((a,b) => a.date.localeCompare(b.date));
      return { ...prevData, entries: updatedEntries };
    });
    setSelectedDate(null);
  };

  const handleSaveCandidate = (candidateData: Candidate) => {
    setCurrentUserData(prevData => {
        if (!prevData) return null;
        const existing = prevData.candidates.find(c => c.id === candidateData.id);
        let updatedCandidates;
        if (existing) {
            updatedCandidates = prevData.candidates.map(c => c.id === candidateData.id ? candidateData : c);
        } else {
            updatedCandidates = [...prevData.candidates, candidateData];
        }
        return { ...prevData, candidates: updatedCandidates };
    });
  };

  const handleToggleCandidateVisibility = (candidateId: string) => {
      setCurrentUserData(prevData => {
          if (!prevData) return null;
          const updatedCandidates = prevData.candidates.map(c => 
              c.id === candidateId ? { ...c, isHidden: !c.isHidden } : c
          );
          return { ...prevData, candidates: updatedCandidates };
      });
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
        : Object.keys(displayedAllUsersData);
      return emailsInScope.flatMap(email => {
        const data = displayedAllUsersData[email];
        if (!data) return [];
        const ownerLabel = data.displayName || email;
        return (data.candidates || []).map(c => ({ ...c, ownerEmail: email, ownerLabel }));
      });
    }, [pipelineScope, pipelineSelectedTeamId, pipelineSelectedUserEmail, teams, displayedAllUsersData, candidates]);

    // Options for the pipeline's per-user selector, sorted by display name.
    const pipelineUserOptions = useMemo(() => {
      return Object.entries(displayedAllUsersData)
        .map(([email, data]: [string, UserData]) => ({ email, label: data.displayName || email }))
        .sort((a, b) => a.label.localeCompare(b.label, 'ja'));
    }, [displayedAllUsersData]);

    // The 全ユーザー tab's ad-hoc comparison selection — an empty selection means "no filter".
    const comparisonUsers = useMemo(() => {
      if (comparisonUserEmails.length === 0) return users;
      return users.filter(u => comparisonUserEmails.includes(u));
    }, [users, comparisonUserEmails]);

    const toggleComparisonUser = (email: string) => {
      setComparisonUserEmails(prev =>
        prev.includes(email) ? prev.filter(e => e !== email) : [...prev, email]
      );
    };


  const monthlyTotals = useMemo<KpiTotals>(() => {
    return calculateMonthlyTotals(entries, allMedia);
  }, [entries, allMedia]);

  const todayTotals = useMemo<KpiTotals>(() => {
    const todayStr = new Date().toLocaleDateString('sv-SE'); // YYYY-MM-DD format
    const todayEntry = entries.find(e => e.date === todayStr);
    return todayEntry ? todayEntry.values : buildAllKpiKeys(allMedia).reduce((acc, key) => {
      acc[key] = 0;
      return acc;
    }, {} as KpiTotals);
  }, [entries, allMedia]);


  const currentMonthPerformanceChartData = useMemo(() => {
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
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
  }, [entries, allMedia]);

  const selectedTeamMemberEmails = useMemo(() => {
    if (!selectedTeamId) return [];
    return teams.find(t => t.id === selectedTeamId)?.memberEmails || [];
  }, [teams, selectedTeamId]);

  const dayOfWeekReplyRateData = useMemo(() => {
    const days = ['日', '月', '火', '水', '木', '金', '土'];
    const scoutsByDay = Array(7).fill(0);
    const repliesByDay = Array(7).fill(0);

    const allEntries = view === 'all_users_kpi'
      ? comparisonUsers.flatMap(email => displayedAllUsersData[email]?.entries || [])
      : view === 'team_kpi'
      ? Object.entries(displayedAllUsersData).filter(([email]) => selectedTeamMemberEmails.includes(email)).flatMap(([, d]: [string, UserData]) => d.entries)
      : entries;
    if (allEntries.length === 0) return null;

    allEntries.forEach(entry => {
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
  }, [entries, view, displayedAllUsersData, selectedTeamMemberEmails, allMedia, comparisonUsers]);


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
          ownerEmail={teamsOwnerEmail}
          userOptions={pipelineUserOptions}
          onClose={() => setIsTeamsModalOpen(false)}
          onCreateTeam={handleCreateTeam}
          onRenameTeam={handleRenameTeam}
          onDeleteTeam={handleDeleteTeam}
          onAddMember={handleAddTeamMember}
          onRemoveMember={handleRemoveTeamMember}
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
          onClose={() => setSelectedDate(null)}
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
               <button type="button" onClick={handleForceSync} disabled={isForcingSync} className="secondary-action-button">
                 {isForcingSync ? '同期中...' : '今すぐ同期'}
               </button>
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
                  <span>媒体別 月次進捗</span>
                  <span className={`toggle-icon ${sectionVisibility.mediaProgress ? 'open' : ''}`}>▼</span>
                </h2>
                <div id="media-progress-content" className={`collapsible-content ${sectionVisibility.mediaProgress ? 'open' : ''}`}>
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
                <span>当月日次パフォーマンストレンド</span>
                <span className={`toggle-icon ${sectionVisibility.monthlyPerformance ? 'open' : ''}`}>▼</span>
              </h2>
              <div id="current-month-performance-content" className={`collapsible-content ${sectionVisibility.monthlyPerformance ? 'open' : ''}`}>
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
                 <MonthOverMonthPerformanceChart entries={entries} allMedia={allMedia} />
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
                <span>曜日別 累積返信率</span>
                <span className={`toggle-icon ${sectionVisibility.dayOfWeekRate ? 'open' : ''}`}>▼</span>
              </h2>
              <div id="day-of-week-rate-content" className={`collapsible-content ${sectionVisibility.dayOfWeekRate ? 'open' : ''}`}>
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
                    <button onClick={() => setComparisonUserEmails(pipelineUserOptions.map(u => u.email))} className="secondary-action-button">全て選択</button>
                    <button onClick={() => setComparisonUserEmails([])} className="secondary-action-button">選択をクリア</button>
                  </div>
                </div>
                <div className="comparison-user-checkbox-list">
                  {pipelineUserOptions.map(u => (
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
              </div>
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.5rem', marginBottom: '0.5rem' }}>
                <button
                  onClick={() => {
                    const label = comparisonUserEmails.length > 0 ? `選択ユーザー${comparisonUsers.length}名` : '全ユーザー';
                    const csvContent = buildTeamProgressCsv(label, comparisonUsers, displayedAllUsersData, allMedia, viewWeekStartDate);
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
                  }}
                  className="export-button"
                >
                  CSV出力
                </button>
                <button onClick={() => fetchAllUsersData()}>更新</button>
              </div>
              <div className="custom-period-export-bar">
                <span>カスタム期間で出力:</span>
                <input type="date" value={customExportStartDate} onChange={(e) => setCustomExportStartDate(e.target.value)} aria-label="開始日" />
                <span>〜</span>
                <input type="date" value={customExportEndDate} onChange={(e) => setCustomExportEndDate(e.target.value)} aria-label="終了日" />
                <button
                  onClick={() => {
                    const label = comparisonUserEmails.length > 0 ? `選択ユーザー${comparisonUsers.length}名` : '全ユーザー';
                    handleCustomPeriodExport(label, comparisonUsers);
                  }}
                  className="export-button"
                >
                  出力
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
                  visibility={{ progress: sectionVisibility.allUsersProgress, dowRate: sectionVisibility.allUsersDayOfWeekRate, weeklySummary: sectionVisibility.allUsersWeeklySummary, memberWeeklySummary: sectionVisibility.allUsersMemberWeeklySummary, grossProfit: sectionVisibility.allUsersGrossProfit }}
                  toggleSection={toggleSection}
                  showGrossProfit={false}
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
              <span>カスタム期間で出力:</span>
              <input type="date" value={customExportStartDate} onChange={(e) => setCustomExportStartDate(e.target.value)} aria-label="開始日" />
              <span>〜</span>
              <input type="date" value={customExportEndDate} onChange={(e) => setCustomExportEndDate(e.target.value)} aria-label="終了日" />
              <button
                onClick={() => {
                  const teamName = teams.find(t => t.id === selectedTeamId)?.name || 'チーム';
                  const teamUsers = selectedTeamMemberEmails.filter(email => displayedAllUsersData[email]);
                  handleCustomPeriodExport(teamName, teamUsers);
                }}
                disabled={!selectedTeamId}
                className="export-button"
              >
                出力
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
                  visibility={{ progress: sectionVisibility.allUsersProgress, dowRate: sectionVisibility.allUsersDayOfWeekRate, weeklySummary: sectionVisibility.allUsersWeeklySummary, memberWeeklySummary: sectionVisibility.allUsersMemberWeeklySummary, grossProfit: sectionVisibility.allUsersGrossProfit }}
                  toggleSection={toggleSection}
              />
            )}
          </div>
        )}
        {view === 'pipeline' && (
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
        )}
      </main>
    </div>
  );
};

const root = ReactDOM.createRoot(document.getElementById('root') as HTMLElement);
root.render(<React.StrictMode><App /></React.StrictMode>);