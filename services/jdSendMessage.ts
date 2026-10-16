// 「求人票出力」で作ったPDFを候補者に送るときの文面（メール／メッセージ）。
//
// 文面は各ユーザーが自分用に書き換えられる（UserData.jdSendTemplates、2026-10-16）。差し込み項目は
// JD_SEND_TEMPLATE_FIELDS。{企業情報} には、担当者が任意で付けた企業概要・企業URL（Web検索で調べて
// 画面で直せる）が入る。付けない場合は空になり、前後の空行は詰める。

import { GoogleGenAI } from '@google/genai';

export type JdSendChannel = 'email' | 'message';
export type JdSendTemplates = Partial<Record<JdSendChannel, string>>;

export const JD_SEND_CHANNEL_LABELS: Record<JdSendChannel, string> = {
  email: 'メール',
  message: 'メッセージ（LINE・チャット）',
};

export const JD_SEND_TEMPLATE_FIELDS = [
  { token: '{候補者名}', label: '候補者名（「様」付き）' },
  { token: '{担当者名}', label: 'あなたの表示名' },
  { token: '{件数}', label: '求人の件数' },
  { token: '{求人一覧}', label: '企業名・ポジションの一覧' },
  { token: '{企業情報}', label: '付けた企業概要・URL（無ければ空）' },
] as const;

/** 既定の文面（差し込み項目入り）。編集画面の初期値と、未設定時の文面に使う。 */
export const defaultJdSendTemplate = (channel: JdSendChannel): string => {
  if (channel === 'message') {
    return [
      '{候補者名}',
      'お疲れさまです、bloomの{担当者名}です。',
      'ご紹介したい求人の求人票をお送りします！',
      '',
      '{求人一覧}',
      '',
      '{企業情報}',
      '',
      '気になる求人があれば、お気軽にお返事ください。',
    ].join('\n');
  }
  return [
    '件名：【bloom】求人票送付のご案内（{件数}件）',
    '',
    '{候補者名}',
    '',
    'お世話になっております。bloomの{担当者名}でございます。',
    'ご紹介したい求人の求人票をお送りいたします。添付のPDFをご確認ください。',
    '',
    '■ご紹介求人',
    '{求人一覧}',
    '',
    '{企業情報}',
    '',
    'ご興味のある求人がございましたら、お気軽にお知らせください。',
    'ご不明な点も、何なりとご質問ください。',
    '',
    '何卒よろしくお願いいたします。',
  ].join('\n');
};

export interface JdSendJob {
  company: string;
  position: string;
}

export interface JdSendCompanyInfo {
  company: string;
  overview: string; // 空なら付けない
  url: string; // 空なら付けない
}

export interface JdSendParams {
  candidateName: string;
  senderName: string;
  jobs: JdSendJob[];
  companyInfos: JdSendCompanyInfo[];
}

const jobList = (channel: JdSendChannel, jobs: JdSendJob[]): string =>
  jobs.map((j, i) => `${channel === 'email' ? `${i + 1}. ` : '・'}${j.company}／${j.position || '企業共通'}`).join('\n');

const companyInfoBlock = (channel: JdSendChannel, infos: JdSendCompanyInfo[]): string => {
  const blocks = infos
    .map(c => ({ ...c, overview: c.overview.trim(), url: c.url.trim() }))
    .filter(c => c.overview || c.url)
    .map(c => [`【${c.company}】`, c.overview, c.url ? (channel === 'email' ? `URL：${c.url}` : c.url) : ''].filter(Boolean).join('\n'));
  if (blocks.length === 0) return '';
  return [channel === 'email' ? '■企業情報' : '▼企業情報', blocks.join('\n\n')].join('\n');
};

/** テンプレートに差し込み項目を入れて文面にする。{企業情報} が無いテンプレートで企業情報を付けた場合は求人一覧の後ろ（無ければ末尾）に入れる。 */
export const fillJdSendTemplate = (channel: JdSendChannel, template: string, params: JdSendParams): string => {
  const info = companyInfoBlock(channel, params.companyInfos);
  let t = template;
  if (info && !t.includes('{企業情報}')) {
    t = t.includes('{求人一覧}') ? t.replace('{求人一覧}', '{求人一覧}\n\n{企業情報}') : `${t.trimEnd()}\n\n{企業情報}`;
  }
  const name = params.candidateName.trim() ? `${params.candidateName.trim()}様` : '';
  return t
    .split('{候補者名}').join(name)
    .split('{担当者名}').join(params.senderName.trim())
    .split('{件数}').join(String(params.jobs.length))
    .split('{求人一覧}').join(jobList(channel, params.jobs))
    .split('{企業情報}').join(info)
    // 候補者名・企業情報が無い時に残る空行を詰める
    .replace(/^\s*\n/, '')
    .replace(/\n{3,}/g, '\n\n')
    .trimEnd();
};

export const buildJdSendText = (channel: JdSendChannel, params: JdSendParams, templates?: JdSendTemplates): string => {
  const custom = templates?.[channel];
  return fillJdSendTemplate(channel, custom?.trim() ? custom : defaultJdSendTemplate(channel), params);
};

/** メールの文面を、先頭の「件名：」行と本文に分ける（件名行が無ければ件名は空）。 */
export const splitEmailSubject = (text: string): { subject: string; body: string } => {
  const m = text.match(/^\s*件名[：:]\s*(.*)\n?/);
  if (!m) return { subject: '', body: text };
  return { subject: m[1].trim(), body: text.slice(m[0].length).replace(/^\s*\n/, '') };
};

// ---- 企業概要・URLのWeb検索（Gemini + Google検索） ----

const MODEL = 'gemini-2.5-flash';
const newAi = () => new GoogleGenAI({ apiKey: process.env.API_KEY as string });

/**
 * 企業の公式サイトURLと、候補者向けの短い企業概要をWeb検索で調べる。hint（求人票の事業内容・本社所在地など）は
 * 同名の別企業と取り違えないための手がかり。結果は画面で担当者が確認・修正してから使う前提。
 */
export const researchCompanyBrief = async (company: string, hint: string): Promise<{ overview: string; url: string }> => {
  const prompt = [
    `人材紹介会社のキャリアアドバイザーが、候補者に求人を紹介するメールに添える企業情報として、「${company}」をWeb検索で調べてください。`,
    hint ? `同名の別企業と混同しないための手がかり（求人票より）:\n${hint}` : '',
    '次の2行だけを、この形式のまま日本語で出力してください。前置き・補足・箇条書き・マークダウンは不要です。',
    'URL: 公式サイトのトップページのURL（確認できなければ空のまま）',
    '概要: 候補者向けに、事業内容・主なサービス・特徴を2〜3文、120字程度で（宣伝的な誇張はしない。確認できない数字は書かない）',
  ].filter(Boolean).join('\n');
  const response = await newAi().models.generateContent({
    model: MODEL,
    contents: prompt,
    config: { tools: [{ googleSearch: {} }] },
  });
  const text = (response.text || '').trim();
  if (!text) throw new Error('Web検索の結果が空でした。');
  const url = (text.match(/URL\s*[:：]\s*(https?:\/\/\S+)/i)?.[1] || '').replace(/[)）。、]+$/, '');
  const overview = (text.match(/概要\s*[:：]\s*([\s\S]+)/)?.[1] || '').replace(/\*\*/g, '').trim();
  if (!overview && !url) throw new Error('企業情報を読み取れませんでした。');
  return { overview, url };
};
