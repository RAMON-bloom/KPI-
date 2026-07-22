// Scans Gmail for emails reporting a pipeline stage-pass (書類選考通過/1次面接通過/etc.) or a
// recommendation/submission (推薦) for one of this user's own pipeline candidates.
//
// Originally this classified each candidate-name-matching email with Gemini, but that was too
// slow for routine use (one AI round-trip per matching email). Rebuilt as pure keyword/regex
// rules after reviewing real production emails (bloom's actual inbox) — the notification
// landscape turned out to be more tractable than expected:
//   - Recommendations (推薦) almost always go through bloom's own internal ATS-integration
//     mailbox (consultant-group@bloom-firm.com) with a small number of fixed templates (HERP
//     Hire's "〜さんを〜に推薦しました", リクナビHRTech's "新規候補者推薦 〜様", HRMOS's "〜への
//     〜様の紹介が完了しました") — not one-per-client-company free text at all.
//   - Stage-pass emails DO vary per client (each uses their own ATS or writes by hand), but the
//     actual result is almost always spelled out as "{stage}" + "通過"/"合格" within a short
//     span of each other (e.g. "一次面接通過と二次面接のご案内", "書類選考結果のご連絡...書類選考
//     の結果、通過でございます"), or a structured "【結果】\n合格" block near a stage mention.
//   - Rejections use a small, consistent vocabulary (不合格/見送り/添いかねる/叶わず/辞退) that's
//     cheap to detect and must gate out everything else, since many ATSes reuse the exact same
//     subject line for both a pass and a fail (e.g. "選考結果のご連絡") — the verdict is only in
//     the body.
// A generic "進んでいただきたく" pass phrase with no named stage falls back to whatever stage the
// candidate's application is CURRENTLY at in our own pipeline data — the email doesn't have to
// spell out the stage name if we already know where that application stood.

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

export interface PipelineMatchCandidate {
  candidateId: string;
  candidateName: string;
  // Company name -> stage-pass KPI keys not yet recorded for that specific application.
  // candidatesSubmitted is intentionally excluded here (see recommendationPending) since it's
  // a candidate-level event, not tied to any one company.
  pendingKpiKeysByCompany: Record<string, PipelineEventKpiKey[]>;
  // This application's CURRENT stage in our own pipeline (PipelineStage value, e.g. '1次面接')
  // — used only as a fallback when an email confirms a pass without naming which stage.
  currentStageByCompany: Record<string, string>;
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
  note: string; // the matched snippet — shown to the user to sanity-check before applying
}

// Only the newest message in a thread is classified — Gmail bodies include the full quoted
// reply chain below it, and matching against old quoted text (a different stage from weeks
// ago) is a real source of false positives. The newest content is always at the top.
const BODY_INSPECT_LENGTH = 3000;

const NEGATIVE_PATTERN = /不合格|お?見送(?:り|らせて)|添いかねる|叶わず|ご縁がなかった|辞退させて|選考を辞退|貴意に添いかねる/;

const RECOMMENDATION_PATTERN = /推薦しました|新規候補者推薦|候補者が推薦されました|の紹介が完了しました|のご推薦ありがとうございます/;

// Ordered latest-stage-first: an email naming two stages (e.g. "一次面接合格・最終面接のご案内"
// — a pass plus the next step being scheduled) must be attributed to the stage that actually
// has a pass/合格 marker next to it, not misread as the later stage just because it's ALSO
// mentioned somewhere in the same email.
const STAGE_PATTERNS: { kpiKey: PipelineEventKpiKey; re: RegExp }[] = [
  { kpiKey: 'finalInterviewPassed', re: /最終(?:面接|選考)/g },
  { kpiKey: 'secondInterviewPassed', re: /(?:二次|2次)(?:面接|選考)/g },
  { kpiKey: 'firstInterviewPassed', re: /(?:一次|1次)(?:面接|選考)/g },
  { kpiKey: 'documentScreeningPassed', re: /書類選考|書類通過/g },
];
const RESULT_WORD_SOURCE = '通過|合格';
const RESULT_WORD_PATTERN = /通過|合格/g;
// Real notification templates often put the verdict a little further down, in its own
// structured block ("【結果】\n合格"), rather than in the same sentence as the stage name. \s
// (not [^\n]) since the marker and the verdict are typically on two separate lines.
const RESULT_BLOCK_PATTERN = new RegExp(`【結果】\\s{0,10}(?:${RESULT_WORD_SOURCE})`);
const MAX_ADJACENCY_GAP = 15;

/**
 * For each 通過/合格 occurrence, finds the NEAREST stage-keyword occurrence (either side, same
 * line) and attributes the result to THAT stage only. This "nearest wins" rule is what makes
 * "一次面接通過と二次面接のご案内" resolve to 一次面接 (immediately adjacent to 通過) and not
 * 二次面接 (mentioned right after, but further from the actual verdict word) — a plain
 * "stage...within 15 chars...result" test in either direction can't tell those apart, since
 * both stages sit within the window; only comparing which one is literally closer can.
 */
function findExplicitStageResult(text: string): { kpiKey: PipelineEventKpiKey; note: string } | null {
  const blockMatch = text.match(RESULT_BLOCK_PATTERN);
  if (blockMatch && blockMatch.index !== undefined) {
    // A structured "【結果】\n合格" block doesn't name its own stage — attribute it to whichever
    // stage keyword's LAST occurrence appears closest before the block.
    const before = text.slice(Math.max(0, blockMatch.index - 200), blockMatch.index);
    let best: { kpiKey: PipelineEventKpiKey; idx: number } | null = null;
    for (const { kpiKey, re } of STAGE_PATTERNS) {
      const occurrences = [...before.matchAll(re)];
      if (occurrences.length === 0) continue;
      const lastIdx = occurrences[occurrences.length - 1].index!;
      if (!best || lastIdx > best.idx) best = { kpiKey, idx: lastIdx };
    }
    if (best) return { kpiKey: best.kpiKey, note: blockMatch[0] };
  }

  const stageOccurrences = STAGE_PATTERNS.flatMap(({ kpiKey, re }) =>
    [...text.matchAll(re)].map(m => ({ kpiKey, index: m.index!, end: m.index! + m[0].length }))
  );
  for (const rm of text.matchAll(RESULT_WORD_PATTERN)) {
    const resultStart = rm.index!;
    const resultEnd = resultStart + rm[0].length;
    let nearest: { kpiKey: PipelineEventKpiKey; gap: number; note: string } | null = null;
    for (const sm of stageOccurrences) {
      const gap = sm.index >= resultEnd ? sm.index - resultEnd : resultStart >= sm.end ? resultStart - sm.end : null;
      if (gap === null || gap > MAX_ADJACENCY_GAP) continue;
      const start = Math.min(sm.index, resultStart);
      const end = Math.max(sm.end, resultEnd);
      if (text.slice(start, end).includes('\n')) continue;
      if (!nearest || gap < nearest.gap) nearest = { kpiKey: sm.kpiKey, gap, note: text.slice(start, end) };
    }
    if (nearest) return { kpiKey: nearest.kpiKey, note: nearest.note };
  }
  return null;
}

// Confirms a pass happened but doesn't name the stage (e.g. "面接選考結果のご連絡"..."選考の結果、
// ぜひ次の選考に進んでいただきたく存じます", or "1次面接へ進めさせていただきたいと判断いたしまし
// た") — falls back to the application's CURRENT pipeline stage (see currentStageByCompany) to
// infer which stage was just passed.
const GENERIC_PASS_PATTERN = /ぜひ次の(?:選考|ステップ)に進んで|次の選考に進んで|選考を進めさせていただきたく|次の面接に進んで|(?:面接|選考)へ進め|進んでいただきたく|進めさせていただきたい/;
const GENERIC_FALLBACK_BY_STAGE: Record<string, PipelineEventKpiKey> = {
  '書類選考': 'documentScreeningPassed',
  '1次面接': 'firstInterviewPassed',
  '2次面接': 'secondInterviewPassed',
  '最終面接': 'finalInterviewPassed',
};

function classifyText(subjectAndBody: string): { kpiKey: PipelineEventKpiKey | null; note: string; isGeneric: boolean } | null {
  if (NEGATIVE_PATTERN.test(subjectAndBody)) return null;
  if (RECOMMENDATION_PATTERN.test(subjectAndBody)) {
    const match = subjectAndBody.match(RECOMMENDATION_PATTERN)!;
    return { kpiKey: 'candidatesSubmitted', note: match[0], isGeneric: false };
  }
  const offerOrPlacement = subjectAndBody.match(/内定承諾|内定を?承諾|入社(?:を)?決意/)
    ? 'placements'
    : subjectAndBody.match(/内定(?:が)?(?:決定|確定|出た|通知)|内定のご連絡|正式に内定/)
    ? 'offersExtended'
    : null;
  if (offerOrPlacement) {
    const re = offerOrPlacement === 'placements' ? /内定承諾|内定を?承諾|入社(?:を)?決意/ : /内定(?:が)?(?:決定|確定|出た|通知)|内定のご連絡|正式に内定/;
    return { kpiKey: offerOrPlacement, note: subjectAndBody.match(re)![0], isGeneric: false };
  }
  const explicit = findExplicitStageResult(subjectAndBody);
  if (explicit) return { ...explicit, isGeneric: false };
  if (GENERIC_PASS_PATTERN.test(subjectAndBody)) {
    const match = subjectAndBody.match(GENERIC_PASS_PATTERN)!;
    return { kpiKey: null, note: match[0], isGeneric: true };
  }
  return null;
}

/**
 * Scans every message in [startDateISO, endDateISOInclusive] for stage-pass/recommendation
 * events involving any of `candidates`. Only messages whose subject+body literally contain a
 * candidate's name are classified at all — same cheap pre-filter as before, now just gating a
 * synchronous regex check instead of a Gemini call. Results are already filtered down to events
 * that are still pending (per each candidate's pendingKpiKeysByCompany / recommendationPending)
 * — callers don't need to re-check for already-recorded duplicates.
 */
export async function detectPipelineAchievements(
  accessToken: string,
  startDateISO: string,
  endDateISOInclusive: string,
  candidates: PipelineMatchCandidate[],
  onProgress?: (done: number, total: number) => void
): Promise<DetectedPipelineAchievement[]> {
  const messages = await fetchFullMessagesInRange(accessToken, startDateISO, endDateISOInclusive);
  const results: DetectedPipelineAchievement[] = [];
  const total = messages.length;

  messages.forEach((msg, i) => {
    onProgress?.(i, total);
    const matched = candidates.filter(c => c.candidateName.trim() && (msg.subject + '\n' + msg.body).includes(c.candidateName.trim()));
    if (matched.length === 0) return;
    const inspectText = `${msg.subject}\n${msg.body.slice(0, BODY_INSPECT_LENGTH)}`;
    const classified = classifyText(inspectText);
    if (!classified) return;

    matched.forEach(candidate => {
      if (classified.kpiKey === 'candidatesSubmitted' && !classified.isGeneric) {
        if (!candidate.recommendationPending) return;
        const companyName = Object.keys(candidate.pendingKpiKeysByCompany).find(c => inspectText.includes(c)) || null;
        results.push({
          candidateId: candidate.candidateId, candidateName: candidate.candidateName, companyName,
          kpiKey: 'candidatesSubmitted', dateISO: msg.dateISO, messageId: msg.id, subject: msg.subject, note: classified.note,
        });
        return;
      }

      // Resolve which of this candidate's companies the email is about — only the ones
      // literally named in the subject/body qualify; if that's ambiguous (none or several
      // match, when the candidate has more than one), skip rather than guess.
      const companies = Object.keys(candidate.pendingKpiKeysByCompany);
      const mentionedCompanies = companies.filter(c => inspectText.includes(c));
      const companyName = mentionedCompanies.length === 1 ? mentionedCompanies[0]
        : mentionedCompanies.length === 0 && companies.length === 1 ? companies[0]
        : null;
      if (!companyName) return;

      const kpiKey = classified.isGeneric
        ? GENERIC_FALLBACK_BY_STAGE[candidate.currentStageByCompany[companyName]]
        : classified.kpiKey;
      if (!kpiKey) return;
      const pending = candidate.pendingKpiKeysByCompany[companyName];
      if (!pending || !pending.includes(kpiKey)) return;

      results.push({
        candidateId: candidate.candidateId, candidateName: candidate.candidateName, companyName,
        kpiKey, dateISO: msg.dateISO, messageId: msg.id, subject: msg.subject, note: classified.note,
      });
    });
  });
  onProgress?.(total, total);

  // Two different emails about the same underlying event (an ATS's advance notice + a
  // confirmation, or two companies' emails both mentioning an as-yet-unrecorded recommendation)
  // can both surface the same event within one scan. Collapse down to one — per (candidate,
  // company, stage) for stage-passes, per candidate only for candidatesSubmitted (a
  // recommendation counts once per candidate, however many companies it appears across) —
  // keeping the earliest dated occurrence.
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
