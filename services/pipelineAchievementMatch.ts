// Scans Gmail for emails reporting a pipeline stage-pass (書類選考通過/1次面接通過/etc.) or a
// recommendation/submission (推薦) for one of this user's own pipeline candidates. Unlike
// gmailScout.ts's scout-reply detection, these emails come from varied ATS systems or
// individual client-company staff with no consistent subject/sender template — so matching
// can't rely on a fixed keyword. Instead: a cheap client-side substring pre-filter (does the
// candidate's name literally appear in the email?) bounds which emails are worth a Gemini call,
// then Gemini reads the actual content and decides whether it really describes an event for
// that candidate/company and which one.

import { GoogleGenAI, Type } from '@google/genai';
import { fetchFullMessagesInRange, GmailPermissionError } from './gmailScout';

export { GmailPermissionError };

// Mirrors index.tsx's GENERAL_KPIS keys (minus the KPIs that aren't pipeline-stage events, e.g.
// scouting-related ones) — must stay in sync with that object's key names, since these are the
// exact field names written into KpiEntry.values.
export const PIPELINE_EVENT_KPI_KEYS = [
  'candidatesSubmitted',
  'documentScreeningPassed',
  'firstInterviewPassed',
  'secondInterviewPassed',
  'finalInterviewPassed',
  'offersExtended',
  'placements',
] as const;
export type PipelineEventKpiKey = typeof PIPELINE_EVENT_KPI_KEYS[number];

const KPI_KEY_DESCRIPTIONS: Record<PipelineEventKpiKey, string> = {
  candidatesSubmitted: '候補者をクライアント企業に推薦・提出したことを示す内容（書類選考の結果が出るより前の「推薦した／提出した」という内容）',
  documentScreeningPassed: '書類選考を通過したことを示す内容',
  firstInterviewPassed: '1次面接を通過したことを示す内容',
  secondInterviewPassed: '2次面接を通過したことを示す内容（3次以降の中間面接に進んだ場合も含める）',
  finalInterviewPassed: '最終面接に合格したことを示す内容',
  offersExtended: '内定が出たことを示す内容',
  placements: '候補者が内定を承諾したことを示す内容',
};

export interface PipelineMatchCandidate {
  candidateId: string;
  candidateName: string;
  // Company name -> stage-pass KPI keys not yet recorded for that specific application.
  // candidatesSubmitted is intentionally excluded here (see recommendationPending) since it's
  // a candidate-level event, not tied to any one company.
  pendingKpiKeysByCompany: Record<string, PipelineEventKpiKey[]>;
  // Whether candidatesSubmitted (推薦) is still pending for this candidate overall — true means
  // it hasn't been recorded for ANY of their applications yet.
  recommendationPending: boolean;
}

export interface DetectedPipelineAchievement {
  candidateId: string;
  candidateName: string;
  companyName: string | null;
  kpiKey: PipelineEventKpiKey;
  dateISO: string;
  messageId: string;
  subject: string;
  note: string; // Gemini's short quote/reason — shown to the user to sanity-check before applying
}

interface RawMatch {
  candidateName: string;
  companyName: string;
  kpiKey: string;
  note: string;
}

async function classifyEmailForCandidates(
  subject: string,
  body: string,
  matchedCandidates: PipelineMatchCandidate[]
): Promise<RawMatch[]> {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY as string });
  const candidateListText = matchedCandidates
    .map(c => {
      const companies = Object.keys(c.pendingKpiKeysByCompany);
      return `- 候補者名: ${c.candidateName} / 応募先企業: ${companies.length > 0 ? companies.join('、') : '（登録なし）'}`;
    })
    .join('\n');
  const kpiListText = PIPELINE_EVENT_KPI_KEYS.map(k => `- ${k}: ${KPI_KEY_DESCRIPTIONS[k]}`).join('\n');

  const prompt = `以下は人材紹介の候補者に関するメールです。本文を読み、次の候補者リストのいずれかについて、選考結果・推薦に関する具体的な出来事が書かれているかを判定してください。

候補者リスト:
${candidateListText}

判定するイベントの種類（kpiKeyの値として使う文字列）:
${kpiListText}

メール件名: ${subject}
メール本文:
${body.slice(0, 6000)}

候補者リストにある候補者名・企業名と一致する内容が本文に明確に書かれている場合のみ matches に含めてください。憶測や一般的な内容だけでは含めないでください。candidateNameとcompanyNameは候補者リストにある表記と完全に一致する文字列を使ってください（対応する企業が特定できない場合はcompanyNameを空文字にしてください）。該当がなければmatchesを空配列にしてください。`;

  const response = await ai.models.generateContent({
    model: 'gemini-2.5-flash',
    contents: prompt,
    config: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          matches: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                candidateName: { type: Type.STRING, description: '候補者リストにある候補者名と完全一致する文字列' },
                companyName: { type: Type.STRING, description: '候補者リストにあるその候補者の応募先企業名と完全一致する文字列。特定できない場合は空文字。' },
                kpiKey: { type: Type.STRING, enum: PIPELINE_EVENT_KPI_KEYS as unknown as string[], description: '該当するイベントの種類' },
                note: { type: Type.STRING, description: 'この判定の根拠となる本文中の内容の要約（日本語、50文字程度）' },
              },
              required: ['candidateName', 'companyName', 'kpiKey', 'note'],
            },
          },
        },
        required: ['matches'],
      },
    },
  });

  try {
    const parsed = JSON.parse(response.text.trim());
    return Array.isArray(parsed.matches) ? parsed.matches : [];
  } catch {
    return [];
  }
}

/**
 * Scans every message in [startDateISO, endDateISOInclusive] for stage-pass/recommendation
 * events involving any of `candidates`. Only messages whose subject+body literally contain a
 * candidate's name are sent to Gemini at all — bounds the number of Gemini calls to roughly the
 * number of genuinely name-matching emails, not the whole inbox. Results are already filtered
 * down to events that are still pending (per each candidate's pendingKpiKeysByCompany /
 * recommendationPending) — callers don't need to re-check for already-recorded duplicates.
 */
export async function detectPipelineAchievements(
  accessToken: string,
  startDateISO: string,
  endDateISOInclusive: string,
  candidates: PipelineMatchCandidate[],
  onProgress?: (done: number, total: number) => void
): Promise<DetectedPipelineAchievement[]> {
  const messages = await fetchFullMessagesInRange(accessToken, startDateISO, endDateISOInclusive);
  const relevant = messages
    .map(msg => ({
      msg,
      matched: candidates.filter(c => c.candidateName.trim() && (msg.subject + '\n' + msg.body).includes(c.candidateName.trim())),
    }))
    .filter(x => x.matched.length > 0);

  const results: DetectedPipelineAchievement[] = [];
  const total = relevant.length;
  let done = 0;
  onProgress?.(0, total);

  // Gemini calls are far more expensive than the Gmail fetches above, so this stays modest even
  // though gmailScout's own internal fetch used a higher concurrency cap.
  const CONCURRENCY = 3;
  let nextIndex = 0;
  async function worker() {
    while (true) {
      const i = nextIndex++;
      if (i >= relevant.length) return;
      const { msg, matched } = relevant[i];
      try {
        const rawMatches = await classifyEmailForCandidates(msg.subject, msg.body, matched);
        rawMatches.forEach(m => {
          const candidate = matched.find(c => c.candidateName === m.candidateName);
          if (!candidate) return;
          if (!(PIPELINE_EVENT_KPI_KEYS as readonly string[]).includes(m.kpiKey)) return;
          const kpiKey = m.kpiKey as PipelineEventKpiKey;
          if (kpiKey === 'candidatesSubmitted') {
            if (!candidate.recommendationPending) return; // already recorded — drop
            results.push({
              candidateId: candidate.candidateId,
              candidateName: candidate.candidateName,
              companyName: m.companyName && candidate.pendingKpiKeysByCompany[m.companyName] ? m.companyName : null,
              kpiKey, dateISO: msg.dateISO, messageId: msg.id, subject: msg.subject, note: m.note || '',
            });
            return;
          }
          const pendingForCompany = m.companyName ? candidate.pendingKpiKeysByCompany[m.companyName] : undefined;
          if (!pendingForCompany || !pendingForCompany.includes(kpiKey)) return; // not a real/pending company+stage combo
          results.push({
            candidateId: candidate.candidateId,
            candidateName: candidate.candidateName,
            companyName: m.companyName,
            kpiKey, dateISO: msg.dateISO, messageId: msg.id, subject: msg.subject, note: m.note || '',
          });
        });
      } catch (error) {
        // A single email's misclassification/timeout shouldn't abort the whole scan.
        console.error('Failed to classify one email for pipeline achievements', error);
      }
      done++;
      onProgress?.(done, total);
    }
  }
  await Promise.all(Array.from({ length: Math.max(1, Math.min(CONCURRENCY, relevant.length)) }, () => worker()));

  // Each per-email classification only checks against state from BEFORE this scan started, so
  // two different emails within the same scan (e.g. an ATS's advance notice + confirmation
  // about the same pass, or two companies' emails both mentioning an as-yet-unrecorded
  // recommendation) can both surface the same underlying event. Collapse those down to one —
  // per (candidate, company, stage) for stage-passes, per candidate only for candidatesSubmitted
  // (the user's explicit requirement: a recommendation counts once per candidate, however many
  // companies it appears across) — keeping the earliest dated occurrence.
  const seenRecommendation = new Set<string>();
  const seenStagePass = new Set<string>();
  const deduped: DetectedPipelineAchievement[] = [];
  [...results].sort((a, b) => a.dateISO.localeCompare(b.dateISO)).forEach(r => {
    if (r.kpiKey === 'candidatesSubmitted') {
      if (seenRecommendation.has(r.candidateId)) return;
      seenRecommendation.add(r.candidateId);
    } else {
      const key = `${r.candidateId}|${r.companyName}|${r.kpiKey}`;
      if (seenStagePass.has(key)) return;
      seenStagePass.add(key);
    }
    deduped.push(r);
  });
  return deduped;
}
