
import React, { useState, useEffect, useMemo, useRef } from 'react';
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
import { signIn, signOut, getCurrentSession, GoogleIdentity } from './services/googleAuth';
import { loadOwnData, saveOwnDataDebounced, readLegacyAppData, loadAllTeammatesData, loadTeamsConfig, saveTeamsConfig } from './services/dataSync';

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

const GENERAL_KPIS = {
  candidatesSubmitted: { label: '候補者推薦数', target: 30 },
  documentScreeningPassed: { label: '書類選考通過数', target: 25 },
  firstInterviewPassed: { label: '1次面接通過数', target: 15 },
  secondInterviewPassed: { label: '2次面接通過数', target: 10 },
  finalInterviewPassed: { label: '最終面接合格数', target: 7 },
  offersExtended: { label: '内定数', target: 5 },
  placements: { label: '内定承諾数', target: 3 },
};

const MEDIA_SOURCES = ['RDS', 'Doda', 'Liiga', 'BIZ', 'Linkedin', 'AMBI', 'Green'] as const;

type MediaSource = typeof MEDIA_SOURCES[number];
type MediaKpiKey = `${Lowercase<MediaSource>}_scoutsSent` | `${Lowercase<MediaSource>}_scoutReplies` | `${Lowercase<MediaSource>}_effectiveReplies` | `${Lowercase<MediaSource>}_documentsCollected` | `${Lowercase<MediaSource>}_effectiveDocumentsCollected` | `${Lowercase<MediaSource>}_initialInterviews` | `${Lowercase<MediaSource>}_effectiveInitialInterviews`;

// FIX: Correctly type MEDIA_KPIS to ensure KpiKey is a union of string literals, not a broad string type.
const MEDIA_KPIS = MEDIA_SOURCES.reduce((acc, source) => {
  const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
  acc[`${sourceKey}_scoutsSent`] = { label: `${source} スカウト数` };
  acc[`${sourceKey}_scoutReplies`] = { label: `${source} スカウト返信数` };
  acc[`${sourceKey}_effectiveReplies`] = { label: `${source} 有効返信数` };
  acc[`${sourceKey}_documentsCollected`] = { label: `${source} 書類回収数` };
  acc[`${sourceKey}_effectiveDocumentsCollected`] = { label: `${source} 有効書類回収数` };
  acc[`${sourceKey}_initialInterviews`] = { label: `${source} 初回面談数` };
  acc[`${sourceKey}_effectiveInitialInterviews`] = { label: `${source} 初回有効面談数` };
  return acc;
}, {} as Record<MediaKpiKey, { label: string }>);


const ALL_KPI_DEFINITIONS = { ...GENERAL_KPIS, ...MEDIA_KPIS };
type KpiKey = keyof typeof ALL_KPI_DEFINITIONS;


const DEFAULT_KPI_TARGETS = Object.keys(ALL_KPI_DEFINITIONS).reduce((acc, key) => {
    const kpiKey = key as KpiKey;
    if (kpiKey in GENERAL_KPIS) {
        acc[kpiKey] = GENERAL_KPIS[kpiKey as keyof typeof GENERAL_KPIS].target;
    } else {
        if (kpiKey.endsWith('_scoutsSent')) {
            acc[kpiKey] = 200;
        } else if (kpiKey.endsWith('_scoutReplies')) {
            acc[kpiKey] = 20;
        } else if (kpiKey.endsWith('_effectiveReplies')) {
            acc[kpiKey] = 5;
        } else if (kpiKey.endsWith('_documentsCollected')) {
            acc[kpiKey] = 6;
        } else if (kpiKey.endsWith('_effectiveDocumentsCollected')) {
            acc[kpiKey] = 4;
        } else if (kpiKey.endsWith('_initialInterviews')) {
            acc[kpiKey] = 8;
        } else if (kpiKey.endsWith('_effectiveInitialInterviews')) {
            acc[kpiKey] = 7;
        }
    }
    return acc;
}, {} as Record<KpiKey, number>);


interface KpiEntry {
  id: number;
  date: string;
  values: { [key in KpiKey]: number };
}

type KpiTotals = { [key in KpiKey]: number };

const calculateMonthlyTotals = (entries: KpiEntry[]): KpiTotals => {
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();

    const totals = Object.keys(ALL_KPI_DEFINITIONS).reduce((acc, key) => {
      acc[key as KpiKey] = 0;
      return acc;
    }, {} as KpiTotals);

    entries.forEach(entry => {
      const entryDate = new Date(entry.date);
      if (entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear) {
        (Object.keys(ALL_KPI_DEFINITIONS) as KpiKey[]).forEach(key => {
          totals[key] += entry.values[key] || 0;
        });
      }
    });

    return totals;
};


// Types for Weekly Summary
interface WeeklyMediaStats {
  source: MediaSource;
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

interface CompanyApplication {
  id: string;
  companyName: string;
  stage: PipelineStage;
  nextAction: string;
  isHidden?: boolean;
}

interface Candidate {
  id: string;
  name: string;
  salary: number; // in JPY万
  currentSalary: number; // in JPY万
  currentCompany: string;
  education: string;
  source: MediaSource | 'Other' | '';
  usingOtherAgents: boolean;
  applications: CompanyApplication[];
  summary: string;
  resumeFile?: { name: string; }; // for backward compatibility
  resumeFiles?: { name: string; }[];
  interviewAudioFile?: { name: string; } | null;
  interviewSummary?: string;
  createdAt: string; // ISO string
  isHidden?: boolean;
}


// --- Multi-user data structures ---
interface UserData {
  entries: KpiEntry[];
  kpiTargets: Record<KpiKey, number>;
  weeklyKpiTargets: Record<KpiKey, number>;
  dailyKpiTargets: Record<KpiKey, number>;
  candidates: Candidate[];
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

const MonthOverMonthPerformanceChart: React.FC<{ entries: KpiEntry[] }> = ({ entries }) => {
    const { chartData, maxReplyRate } = useMemo(() => {
        const monthlyData: Record<string, KpiTotals> = {};

        entries.forEach(entry => {
            const month = entry.date.substring(0, 7); // YYYY-MM
            if (!monthlyData[month]) {
                monthlyData[month] = Object.keys(ALL_KPI_DEFINITIONS).reduce((acc, key) => {
                  acc[key as KpiKey] = 0;
                  return acc;
                }, {} as KpiTotals);
            }
            
            (Object.keys(entry.values) as KpiKey[]).forEach(key => {
                monthlyData[month][key] += entry.values[key] || 0;
            });
        });

        const sortedMonths = Object.keys(monthlyData).sort();
        
        const labels = sortedMonths;
        const scoutRepliesData = sortedMonths.map(month => getTotalFromLump(monthlyData[month], '_scoutReplies'));
        const docScreeningPassedData = sortedMonths.map(month => monthlyData[month].documentScreeningPassed || 0);
        const firstInterviewPassedData = sortedMonths.map(month => monthlyData[month].firstInterviewPassed || 0);
        const secondInterviewPassedData = sortedMonths.map(month => monthlyData[month].secondInterviewPassed || 0);
        const offersExtendedData = sortedMonths.map(month => monthlyData[month].offersExtended || 0);
        const placementsData = sortedMonths.map(month => monthlyData[month].placements || 0);
        const replyRateData = sortedMonths.map(month => {
            const scoutsSent = getTotalFromLump(monthlyData[month], '_scoutsSent');
            const scoutReplies = getTotalFromLump(monthlyData[month], '_scoutReplies');
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
    }, [entries]);
    
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
            {data.mediaStats.map(({ source, scoutsSent, scoutReplies, effectiveReplies, documentsCollected, effectiveDocumentsCollected, initialInterviews, effectiveInitialInterviews }) => {
              const replyRate = scoutsSent > 0 ? (scoutReplies / scoutsSent) * 100 : 0;
              const effectiveReplyRate = scoutReplies > 0 ? (effectiveReplies / scoutReplies) * 100 : 0;
              const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
              
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
    </div>
  );
};


const MediaKpiCard: React.FC<{
    source: MediaSource;
    monthlyTotals: KpiTotals;
    kpiTargets: Record<KpiKey, number>;
}> = ({ source, monthlyTotals, kpiTargets }) => {
    const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
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
        <div className="media-kpi-card" aria-label={`${source}の進捗`}>
            <h3>{source}</h3>
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
  onSave: (date: string, values: KpiTotals) => void;
  onClose: () => void;
}> = ({ date, initialValues, onSave, onClose }) => {
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
  
  const handleClear = () => {
      if (window.confirm('この日の実績をすべてクリアします。よろしいですか？')) {
          const emptyValues = Object.keys(ALL_KPI_DEFINITIONS).reduce((acc, key) => {
              acc[key as KpiKey] = 0;
              return acc;
          }, {} as KpiTotals);
          onSave(date, emptyValues);
      }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const valuesWithDefaults = Object.keys(ALL_KPI_DEFINITIONS).reduce((acc, key) => {
      acc[key as KpiKey] = entryValues[key as KpiKey] || 0;
      return acc;
    }, {} as KpiTotals);
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
              {MEDIA_SOURCES.map(source => {
                const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
                const scoutsKey = `${sourceKey}_scoutsSent` as KpiKey;
                const repliesKey = `${sourceKey}_scoutReplies` as KpiKey;
                const effectiveRepliesKey = `${sourceKey}_effectiveReplies` as KpiKey;
                const documentsCollectedKey = `${sourceKey}_documentsCollected` as KpiKey;
                const effectiveDocumentsCollectedKey = `${sourceKey}_effectiveDocumentsCollected` as KpiKey;
                const interviewsKey = `${sourceKey}_initialInterviews` as KpiKey;
                const effectiveInterviewsKey = `${sourceKey}_effectiveInitialInterviews` as KpiKey;
                return (
                  <fieldset key={source} className="media-fieldset">
                    <legend>{source}</legend>
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
                          aria-label={`${source} スカウト数を入力`}
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
                          aria-label={`${source} スカウト返信数を入力`}
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
                          aria-label={`${source} 有効返信数を入力`}
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
                          aria-label={`${source} 書類回収数を入力`}
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
                          aria-label={`${source} 有効書類回収数を入力`}
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
                          aria-label={`${source} 初回面談数を入力`}
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
                          aria-label={`${source} 初回有効面談数を入力`}
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


const getTotalFromLump = (lump: { [key: string]: number | undefined }, kpiSuffix: string): number => {
    if (!lump) return 0;
    return MEDIA_SOURCES.reduce((acc, source) => {
        const sourceKey = source.toLowerCase();
        const kpiKey = `${sourceKey}${kpiSuffix}`;
        return acc + (lump[kpiKey] || 0);
    }, 0);
};

const CalendarView: React.FC<{
  viewDate: Date;
  entriesByDate: Map<string, KpiTotals>;
  onDayClick: (date: string) => void;
  onPrevMonth: () => void;
  onNextMonth: () => void;
}> = ({ viewDate, entriesByDate, onDayClick, onPrevMonth, onNextMonth }) => {
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
      interviewed: data ? getTotalFromLump(data, '_initialInterviews') : 0,
      effectiveInterviewed: data ? getTotalFromLump(data, '_effectiveInitialInterviews') : 0,
      submitted: data?.candidatesSubmitted || 0,
      collected: data ? getTotalFromLump(data, '_documentsCollected') : 0,
      effectiveCollected: data ? getTotalFromLump(data, '_effectiveDocumentsCollected') : 0,
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
    onClose: () => void;
    onCreateTeam: (name: string) => void;
    onRenameTeam: (teamId: string, name: string) => void;
    onDeleteTeam: (teamId: string) => void;
    onAddMember: (teamId: string, email: string) => void;
    onRemoveMember: (teamId: string, email: string) => void;
}> = ({ teams, isEditable, ownerEmail, onClose, onCreateTeam, onRenameTeam, onDeleteTeam, onAddMember, onRemoveMember }) => {
    const [newTeamName, setNewTeamName] = useState('');
    const [editingTeamId, setEditingTeamId] = useState<string | null>(null);
    const [editedName, setEditedName] = useState('');
    const [memberInputs, setMemberInputs] = useState<Record<string, string>>({});

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

    const handleAddMember = (teamId: string) => {
        const email = (memberInputs[teamId] || '').trim();
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
                                                <span>{email}</span>
                                                {isEditable && (
                                                    <button onClick={() => onRemoveMember(team.id, email)} className="delete-user-button">削除</button>
                                                )}
                                            </li>
                                        ))}
                                    </ul>
                                    {isEditable && (
                                        <div style={{ display: 'flex', gap: '0.5rem', marginTop: '0.5rem' }}>
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


const DailyProgressCard: React.FC<{
    source: MediaSource;
    todayTotals: KpiTotals;
    dailyKpiTargets: Record<KpiKey, number>;
}> = ({ source, todayTotals, dailyKpiTargets }) => {
    const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
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
            <h3>{source}</h3>
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
    todayTotals: KpiTotals;
    dailyKpiTargets: Record<KpiKey, number>;
}> = ({ todayTotals, dailyKpiTargets }) => {
    const hasTargets = MEDIA_SOURCES.some(source => {
         const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
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
            {MEDIA_SOURCES.map(source => (
                <DailyProgressCard
                    key={source}
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
}> = ({ onSave, onClose, initialData }) => {
    const defaultCandidate: Candidate = {
        id: initialData?.id || `candidate_${Date.now()}`,
        name: '',
        salary: 0,
        currentSalary: 0,
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
    const resumeInputRef = useRef<HTMLInputElement>(null);
    const audioInputRef = useRef<HTMLInputElement>(null);

    const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
        const { name, value } = e.target;
        setCandidate(prev => ({ ...prev, [name]: (name === 'salary' || name === 'currentSalary') ? Number(value) : value }));
    };

    const handleApplicationChange = (index: number, field: keyof CompanyApplication, value: string) => {
        const newApplications = [...candidate.applications];
        newApplications[index] = { ...newApplications[index], [field]: value };
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
                text: 'この面談の音声データを要約してください。候補者の強み、弱み、懸念事項、そして特筆すべきスキルや経験について、箇条書きで簡潔にまとめてください。',
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
                    <label htmlFor="source">集客媒体</label>
                    <select id="source" name="source" value={candidate.source} onChange={handleChange}>
                        <option value="">選択してください</option>
                        {MEDIA_SOURCES.map(s => <option key={s} value={s}>{s}</option>)}
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
        id: '', companyName: '', stage: '打診', nextAction: ''
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
                    nextAction: ''
                });
            }
        }
    }, [isOpen, initialData]);

    if (!isOpen) return null;

    const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
        const { name, value } = e.target;
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

    const modalTitle = `${candidateName}さん - ${initialData ? '選考情報を編集' : '選考情報を追加'}`;

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
                </form>
                <div className="modal-footer">
                    <button type="button" onClick={onClose} className="cancel-button">キャンセル</button>
                    <button type="submit" form="application-form" className="submit-button">保存</button>
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

    const stageColorMap: Record<PipelineStage, string> = {
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
                <div key={stage} className="pipeline-stage-card" style={{'--badge-color': stageColorMap[stage]} as React.CSSProperties}>
                    <span className="stage-name">{stage}</span>
                    <span className="stage-count">{stageCounts[stage]}</span>
                </div>
            ))}
        </div>
    );
};


const SourceEffectivenessReport: React.FC<{ candidates: Candidate[] }> = ({ candidates }) => {
    const reportData = useMemo(() => {
        const visibleCandidates = candidates.filter(c => !c.isHidden);
        const sources = [...MEDIA_SOURCES, 'Other', ''];
        
        const stats = sources.map(source => {
            const sourceName = source || '未設定';
            const sourceCandidates = visibleCandidates.filter(c => (c.source || '未設定') === sourceName);
            const total = sourceCandidates.length;

            if (total === 0) return null;

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
        }).filter((stat): stat is NonNullable<typeof stat> => stat !== null)
          .sort((a, b) => b.placements - a.placements || b.total - a.total);

        return stats;
    }, [candidates]);

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
    onSave: (candidate: Candidate) => void;
    onToggleVisibility: (candidateId: string) => void;
}> = ({ candidates, onSave, onToggleVisibility }) => {
    const [isModalOpen, setIsModalOpen] = useState(false);
    const [editingCandidate, setEditingCandidate] = useState<Candidate | null>(null);
    const [searchTerm, setSearchTerm] = useState('');
    const [showHidden, setShowHidden] = useState(false);
    const [sortConfig, setSortConfig] = useState<{ key: keyof Candidate; direction: 'asc' | 'desc' } | null>({ key: 'createdAt', direction: 'desc'});
    const [expandedCandidateId, setExpandedCandidateId] = useState<string | null>(null);
    const [isReportVisible, setIsReportVisible] = useState(true);
    const [showHiddenApps, setShowHiddenApps] = useState(false);
    
    // State for the new Application Modal
    const [isApplicationModalOpen, setIsApplicationModalOpen] = useState(false);
    const [applicationModalData, setApplicationModalData] = useState<{
        candidate: Candidate | null,
        application: CompanyApplication | null
    }>({ candidate: null, application: null });

    const handleAdd = () => {
        setEditingCandidate(null);
        setIsModalOpen(true);
    };
    
    const handleEdit = (candidate: Candidate) => {
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
        setApplicationModalData({ candidate, application });
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
            '氏名', '現職企業名', '最終学歴', '現年収(万円)', '希望年収(万円)', 
            '集客媒体', '他エージェント使用状況', '登録日', '概要',
            '応募企業名', '進捗状況', '次アクション'
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
                    ];
                    if (index === 0) {
                        rows.push([...commonData, ...applicationData]);
                    } else {
                        rows.push([...emptyCommonData, ...applicationData]);
                    }
                });
            } else {
                 rows.push([...commonData, '', '', '']);
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
            return matchesSearch && matchesVisibility;
        });
    }, [candidates, searchTerm, showHidden]);

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
    
    const getSortIndicator = (key: keyof Candidate) => {
        if (!sortConfig || sortConfig.key !== key) return '';
        return sortConfig.direction === 'asc' ? ' ▲' : ' ▼';
    };

    const stageColorMap: Record<PipelineStage, string> = {
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

            <PipelineDashboard candidates={candidates} />

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
                    <SourceEffectivenessReport candidates={candidates} />
                </div>
            </div>

            <div className="add-candidate-action-bar">
                <button onClick={handleAdd} className="add-candidate-large-button">
                    + 新規候補者を追加
                </button>
            </div>
            
             <div className="pipeline-list-controls">
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

                    return (
                        <div key={c.id} className="candidate-list-item">
                            <div className="candidate-card-main">
                                <div className="candidate-card-header">
                                    <h3
                                      onClick={() => handleEdit(c)}
                                      className="candidate-name-clickable"
                                      role="button"
                                      tabIndex={0}
                                      onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleEdit(c); }}}
                                      title={`${c.name}を編集`}
                                    >
                                      {c.name}
                                    </h3>
                                    {c.usingOtherAgents && <span className="other-agent-tag">他エージェント利用中</span>}
                                    <div className="candidate-card-actions">
                                        <button onClick={() => handleOpenApplicationModal(c, null)} className="add-selection-button">+ 選考追加</button>
                                        <button onClick={() => handleEdit(c)} className="edit-user-button">編集</button>
                                        <button onClick={() => onToggleVisibility(c.id)} className={showHidden ? "secondary-action-button" : "delete-user-button"}>
                                            {showHidden ? '再表示' : '非表示'}
                                        </button>
                                    </div>
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
                                                        className="status-badge" 
                                                        style={{'--badge-color': stageColorMap[app.stage]} as React.CSSProperties}
                                                        title={`${app.companyName}: ${app.stage}`}
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
                                                          <div className="detail-card-actions">
                                                            <button onClick={() => handleOpenApplicationModal(c, app)} className="edit-user-button">編集</button>
                                                            <button onClick={() => handleToggleApplicationVisibility(c, app.id)} className={app.isHidden ? "secondary-action-button" : "delete-user-button"}>
                                                                {app.isHidden ? '再表示' : '非表示'}
                                                            </button>
                                                          </div>
                                                      </div>
                                                      <div className="detail-card-body">
                                                          <div className="detail-card-item">
                                                              <span>進捗状況:</span>
                                                              <span 
                                                                  className="status-badge" 
                                                                  style={{'--badge-color': stageColorMap[app.stage]} as React.CSSProperties}
                                                              >
                                                                  {app.stage}
                                                              </span>
                                                          </div>
                                                          <div className="detail-card-item">
                                                              <span>次アクション:</span>
                                                              <span>{app.nextAction || '未設定'}</span>
                                                          </div>
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
        </section>
    );
};


const AllUsersDashboard: React.FC<{
  users: string[];
  allUsersData: Record<string, UserData>;
  dayOfWeekReplyRateData: any | null;
  visibility: { progress: boolean; dowRate: boolean };
  toggleSection: (key: 'allUsersProgress' | 'allUsersDayOfWeekRate') => void;
}> = ({ users, allUsersData, dayOfWeekReplyRateData, visibility, toggleSection }) => {
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
                  if (!userData) {
                    return <tr key={user}><td colSpan={Object.keys(GENERAL_KPIS).length + 7}>{user}のデータがありません。</td></tr>;
                  }

                  const monthlyTotals = calculateMonthlyTotals(userData.entries || []);
                  const kpiTargets = { ...DEFAULT_KPI_TARGETS, ...(userData.kpiTargets || {}) };
                  
                  const sent = MEDIA_SOURCES.reduce((acc, source) => {
                      const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
                      const sentKey = `${sourceKey}_scoutsSent` as KpiKey;
                      return acc + (monthlyTotals[sentKey] || 0);
                  }, 0);
                  const replies = MEDIA_SOURCES.reduce((acc, source) => {
                      const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
                      const replyKey = `${sourceKey}_scoutReplies` as KpiKey;
                      return acc + (monthlyTotals[replyKey] || 0);
                  }, 0);
                  const effectiveReplies = MEDIA_SOURCES.reduce((acc, source) => {
                      const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
                      const effectiveReplyKey = `${sourceKey}_effectiveReplies` as KpiKey;
                      return acc + (monthlyTotals[effectiveReplyKey] || 0);
                  }, 0);
                  
                  const replyRate = sent > 0 ? (replies / sent) * 100 : 0;
                  const effectiveReplyRate = replies > 0 ? (effectiveReplies / replies) * 100 : 0;

                  const documentsCollected = getTotalFromLump(monthlyTotals, '_documentsCollected');
                  const documentsCollectedTarget = getTotalFromLump(kpiTargets, '_documentsCollected');
                  const documentsCollectedProgress = documentsCollectedTarget > 0 ? Math.min((documentsCollected / documentsCollectedTarget) * 100, 100) : 0;
                  
                  const effectiveDocumentsCollected = getTotalFromLump(monthlyTotals, '_effectiveDocumentsCollected');
                  const effectiveDocumentsCollectedTarget = getTotalFromLump(kpiTargets, '_effectiveDocumentsCollected');
                  const effectiveDocumentsCollectedProgress = effectiveDocumentsCollectedTarget > 0 ? Math.min((effectiveDocumentsCollected / effectiveDocumentsCollectedTarget) * 100, 100) : 0;

                  const initialInterviews = getTotalFromLump(monthlyTotals, '_initialInterviews');
                  const effectiveInitialInterviews = getTotalFromLump(monthlyTotals, '_effectiveInitialInterviews');
                  const effectiveInterviewRate = initialInterviews > 0 ? (effectiveInitialInterviews / initialInterviews) * 100 : 0;

                  const initialInterviewsTarget = getTotalFromLump(kpiTargets, '_initialInterviews');
                  const initialInterviewsProgress = initialInterviewsTarget > 0 ? Math.min((initialInterviews / initialInterviewsTarget) * 100, 100) : 0;

                  return (
                    <tr key={user}>
                      <td>{user}</td>
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
}

interface ReportData {
  generalTotals: KpiTotals;
  mediaStats: Array<{
    source: MediaSource;
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


const CustomPeriodReport: React.FC<CustomPeriodReportProps> = ({ entries }) => {
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

    const totals = Object.keys(ALL_KPI_DEFINITIONS).reduce((acc, key) => {
      acc[key as KpiKey] = 0;
      return acc;
    }, {} as KpiTotals);

    filteredEntries.forEach(entry => {
      (Object.keys(ALL_KPI_DEFINITIONS) as KpiKey[]).forEach(key => {
        totals[key] += entry.values[key] || 0;
      });
    });

    const mediaStats = MEDIA_SOURCES.map(source => {
      const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
      const scoutsKey = `${sourceKey}_scoutsSent` as KpiKey;
      const repliesKey = `${sourceKey}_scoutReplies` as KpiKey;
      const effectiveRepliesKey = `${sourceKey}_effectiveReplies` as KpiKey;
      const documentsCollectedKey = `${sourceKey}_documentsCollected` as KpiKey;
      const effectiveDocumentsCollectedKey = `${sourceKey}_effectiveDocumentsCollected` as KpiKey;
      const interviewsKey = `${sourceKey}_initialInterviews` as KpiKey;
      const effectiveInterviewsKey = `${sourceKey}_effectiveInitialInterviews` as KpiKey;
      return {
        source,
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
  | 'allUsersProgress' | 'allUsersDayOfWeekRate';


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

  // --- Consolidated user data state ---
  const [currentUserData, setCurrentUserData] = useState<UserData | null>(null);
  
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
  });

  const toggleSection = (sectionKey: SectionVisibilityKeys) => {
    setSectionVisibility(prev => ({
        ...prev,
        [sectionKey]: !prev[sectionKey]
    }));
  };
  
  // Restore Google session (if any)
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


  // Load the signed-in user's data from Google Drive (Drive is the source of truth;
  // the local cache inside loadOwnData is only a fallback for offline/error cases).
  useEffect(() => {
    if (!currentIdentity) {
      setCurrentUserData(null);
      setDriveFileId(null);
      return;
    }
    let cancelled = false;
    setIsLoadingUserData(true);
    (async () => {
      const email = currentIdentity.email;
      const result = await loadOwnData<UserData>(email);
      if (cancelled) return;

      if (result.data) {
        const d = result.data;
        setCurrentUserData({
          entries: d.entries || [],
          candidates: d.candidates || [],
          kpiTargets: { ...DEFAULT_KPI_TARGETS, ...(d.kpiTargets || {}) },
          weeklyKpiTargets: { ...DEFAULT_KPI_TARGETS, ...(d.weeklyKpiTargets || {}) },
          dailyKpiTargets: { ...DEFAULT_KPI_TARGETS, ...(d.dailyKpiTargets || {}) },
        });
        setDriveFileId(result.driveFileId);
      } else {
        // Brand-new signed-in user: offer to claim any pre-Google-login local data.
        const legacy = readLegacyAppData();
        const legacyNames = (legacy?.users || []).filter(name => legacy?.userData?.[name]);
        if (legacyNames.length > 0) {
          setLegacyMigrationChoices(legacyNames);
        }
        setCurrentUserData({
          entries: [],
          candidates: [],
          kpiTargets: DEFAULT_KPI_TARGETS,
          weeklyKpiTargets: DEFAULT_KPI_TARGETS,
          dailyKpiTargets: DEFAULT_KPI_TARGETS,
        });
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
      kpiTargets: { ...DEFAULT_KPI_TARGETS, ...(legacyUserData.kpiTargets || {}) },
      weeklyKpiTargets: { ...DEFAULT_KPI_TARGETS, ...(legacyUserData.weeklyKpiTargets || {}) },
      dailyKpiTargets: { ...DEFAULT_KPI_TARGETS, ...(legacyUserData.dailyKpiTargets || {}) },
    });
  };


  // Load every teammate's Drive-shared data (domain-wide, cross-device) when switching to
  // the all-users or team-filtered views.
  useEffect(() => {
    if ((view !== 'all_users_kpi' && view !== 'team_kpi') || !isInitialized || !currentIdentity) return;
    let cancelled = false;
    setIsLoadingAllUsers(true);
    (async () => {
      try {
        const teammates = await loadAllTeammatesData<UserData>();
        if (cancelled) return;
        const merged: Record<string, UserData> = {};
        teammates.forEach(({ email, data }) => {
          merged[email] = {
            entries: data.entries || [],
            candidates: data.candidates || [],
            kpiTargets: { ...DEFAULT_KPI_TARGETS, ...(data.kpiTargets || {}) },
            weeklyKpiTargets: { ...DEFAULT_KPI_TARGETS, ...(data.weeklyKpiTargets || {}) },
            dailyKpiTargets: { ...DEFAULT_KPI_TARGETS, ...(data.dailyKpiTargets || {}) },
          };
        });
        setAllUsersData(merged);
        setUsers(Object.keys(merged));
      } catch (error) {
        console.error("Failed to load teammates' data from Drive", error);
        setAllUsersData({});
      } finally {
        if (!cancelled) setIsLoadingAllUsers(false);
      }
    })();
    return () => { cancelled = true; };
  }, [view, isInitialized, currentIdentity]);

  // Load the shared teams config when opening the team-filtered view or the Teams management modal.
  useEffect(() => {
    if (!currentIdentity || !isInitialized) return;
    if (view !== 'team_kpi' && !isTeamsModalOpen) return;
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
  }, [currentIdentity, isInitialized, view, isTeamsModalOpen]);

  const isTeamsEditable = !teamsOwnerEmail || teamsOwnerEmail === currentIdentity?.email;

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

    const handleLogout = () => {
        signOut();
        setCurrentIdentity(null);
        setCurrentUser(null);
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
      kpiTargets: { ...DEFAULT_KPI_TARGETS, ...(currentUserData?.kpiTargets || {}) },
      weeklyKpiTargets: { ...DEFAULT_KPI_TARGETS, ...(currentUserData?.weeklyKpiTargets || {}) },
      dailyKpiTargets: { ...DEFAULT_KPI_TARGETS, ...(currentUserData?.dailyKpiTargets || {}) },
      candidates: currentUserData?.candidates || [],
    }), [currentUserData]);


  const monthlyTotals = useMemo<KpiTotals>(() => {
    return calculateMonthlyTotals(entries);
  }, [entries]);

  const todayTotals = useMemo<KpiTotals>(() => {
    const todayStr = new Date().toLocaleDateString('sv-SE'); // YYYY-MM-DD format
    const todayEntry = entries.find(e => e.date === todayStr);
    return todayEntry ? todayEntry.values : Object.keys(ALL_KPI_DEFINITIONS).reduce((acc, key) => {
      acc[key as KpiKey] = 0;
      return acc;
    }, {} as KpiTotals);
  }, [entries]);


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
            const scouts = MEDIA_SOURCES.reduce((sum, source) => sum + (entry.values[`${source.toLowerCase()}_scoutsSent` as KpiKey] || 0), 0);
            const replies = MEDIA_SOURCES.reduce((sum, source) => sum + (entry.values[`${source.toLowerCase()}_scoutReplies` as KpiKey] || 0), 0);
            
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
  }, [entries]);

  const selectedTeamMemberEmails = useMemo(() => {
    if (!selectedTeamId) return [];
    return teams.find(t => t.id === selectedTeamId)?.memberEmails || [];
  }, [teams, selectedTeamId]);

  const dayOfWeekReplyRateData = useMemo(() => {
    const days = ['日', '月', '火', '水', '木', '金', '土'];
    const scoutsByDay = Array(7).fill(0);
    const repliesByDay = Array(7).fill(0);

    const allEntries = view === 'all_users_kpi'
      ? Object.values(allUsersData).flatMap(d => d.entries)
      : view === 'team_kpi'
      ? Object.entries(allUsersData).filter(([email]) => selectedTeamMemberEmails.includes(email)).flatMap(([, d]: [string, UserData]) => d.entries)
      : entries;
    if (allEntries.length === 0) return null;

    allEntries.forEach(entry => {
        const dayOfWeek = new Date(entry.date).getDay();
        const scouts = MEDIA_SOURCES.reduce((sum, source) => sum + (entry.values[`${source.toLowerCase()}_scoutsSent` as KpiKey] || 0), 0);
        const replies = MEDIA_SOURCES.reduce((sum, source) => sum + (entry.values[`${source.toLowerCase()}_scoutReplies` as KpiKey] || 0), 0);
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
  }, [entries, view, allUsersData, selectedTeamMemberEmails]);


  const weeklySummaryData = useMemo<WeeklyData>(() => {
      const weekStart = viewWeekStartDate.getTime();
      const weekEnd = new Date(viewWeekStartDate).setDate(viewWeekStartDate.getDate() + 6);

      const weeklyEntries = entries.filter(entry => {
          const entryTime = new Date(entry.date).getTime();
          return entryTime >= weekStart && entryTime <= weekEnd;
      });

      const weeklyTotals = weeklyEntries.reduce((acc, entry) => {
          (Object.keys(ALL_KPI_DEFINITIONS) as KpiKey[]).forEach(key => {
              acc[key] = (acc[key] || 0) + (entry.values[key] || 0);
          });
          return acc;
      }, {} as KpiTotals);

      const mediaStats = MEDIA_SOURCES.map(source => {
          const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
          return {
              source,
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
  }, [entries, viewWeekStartDate]);
  
  const entriesByDate = useMemo(() => {
    return new Map(entries.map(entry => [entry.date, entry.values]));
  }, [entries]);

  if (!isInitialized) {
      return <div className="loading-container">読み込み中...</div>;
  }

  if (!currentIdentity) {
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
                    {isSigningIn ? 'ログイン中...' : 'Googleでログイン'}
                </button>
                <p style={{ marginTop: '1rem', fontSize: '0.85rem', color: '#666' }}>
                    bloom-firm.com のGoogleアカウントでログインしてください。
                </p>
            </div>
        </div>
    );
  }

  if (isLoadingUserData) {
      return <div className="loading-container">Googleドライブからデータを読み込み中...</div>;
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
          onClose={() => setIsTeamsModalOpen(false)}
          onCreateTeam={handleCreateTeam}
          onRenameTeam={handleRenameTeam}
          onDeleteTeam={handleDeleteTeam}
          onAddMember={handleAddTeamMember}
          onRemoveMember={handleRemoveTeamMember}
        />
      )}
      {selectedDate && (
        <DateEntryModal
          date={selectedDate}
          initialValues={entriesByDate.get(selectedDate) || null}
          onSave={handleSaveEntry}
          onClose={() => setSelectedDate(null)}
        />
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
              <span style={{ fontSize: '0.9rem', color: '#333' }}>{currentIdentity.name}</span>
            )}
            <button onClick={() => setIsTeamsModalOpen(true)}>チーム管理</button>
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
                   <DailyProgress todayTotals={todayTotals} dailyKpiTargets={dailyKpiTargets} />
                </div>
            </section>
            
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
                     {MEDIA_SOURCES.map(source => (
                          <MediaKpiCard
                              key={source}
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
                 <MonthOverMonthPerformanceChart entries={entries} />
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
                <CustomPeriodReport entries={entries} />
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
                             {MEDIA_SOURCES.map(source => {
                                 const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
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
                                 <fieldset key={source} className="media-fieldset">
                                     <legend>{source}</legend>
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
                             {MEDIA_SOURCES.map(source => {
                                 const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
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
                                 <fieldset key={source} className="media-fieldset">
                                     <legend>{source}</legend>
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
                             {MEDIA_SOURCES.map(source => {
                                 const sourceKey = source.toLowerCase() as Lowercase<MediaSource>;
                                 const fields: {key: KpiKey, label: string}[] = [
                                     {key: `${sourceKey}_scoutsSent`, label: 'スカウト数'},
                                     {key: `${sourceKey}_scoutReplies`, label: 'スカウト返信数'},
                                     {key: `${sourceKey}_effectiveReplies`, label: '有効返信数'},
                                 ];
                                return (
                                 <fieldset key={source} className="media-fieldset">
                                     <legend>{source}</legend>
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
            <AllUsersDashboard
                users={users}
                allUsersData={allUsersData}
                dayOfWeekReplyRateData={dayOfWeekReplyRateData}
                visibility={{ progress: sectionVisibility.allUsersProgress, dowRate: sectionVisibility.allUsersDayOfWeekRate }}
                toggleSection={toggleSection}
            />
          )
        )}
        {view === 'team_kpi' && (
          <div>
            <div className="form-group" style={{ maxWidth: '300px', marginBottom: '1.5rem' }}>
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
            {!selectedTeamId ? (
              <p className="no-data-message">チームを選択してください。チームがまだない場合は「チーム管理」から作成してください。</p>
            ) : isLoadingAllUsers ? (
              <div className="loading-container">チームメンバーのデータをGoogleドライブから読み込み中...</div>
            ) : (
              <AllUsersDashboard
                  users={selectedTeamMemberEmails.filter(email => allUsersData[email])}
                  allUsersData={allUsersData}
                  dayOfWeekReplyRateData={dayOfWeekReplyRateData}
                  visibility={{ progress: sectionVisibility.allUsersProgress, dowRate: sectionVisibility.allUsersDayOfWeekRate }}
                  toggleSection={toggleSection}
              />
            )}
          </div>
        )}
        {view === 'pipeline' && (
            <CandidatePipelineView 
                candidates={candidates}
                onSave={handleSaveCandidate}
                onToggleVisibility={handleToggleCandidateVisibility}
            />
        )}
      </main>
    </div>
  );
};

const root = ReactDOM.createRoot(document.getElementById('root') as HTMLElement);
root.render(<React.StrictMode><App /></React.StrictMode>);