// 紹介先企業カオスマップの型定義とシード値。Reactに依存しない純粋なデータ定義。
// 難易度(S〜D)は固定5値なので設定エンティティ化しない。横軸カテゴリは可変長にして、
// 将来ユーザーが列構成を変える場合もDrive上のconfigを書き換えるだけで対応できるようにする。

export type ChaosMapDifficulty = 'S' | 'A' | 'B' | 'C' | 'D';
export type ChaosMapKind = 'firm' | 'it';

export const CHAOS_MAP_DIFFICULTIES: ChaosMapDifficulty[] = ['S', 'A', 'B', 'C', 'D'];

export const CHAOS_MAP_KIND_LABELS: Record<ChaosMapKind, string> = {
  firm: 'Firm系',
  it: 'IT系',
};

export interface ChaosMapBadge {
  id: string;
  emoji: string;
  label: string;
}

export interface ChaosMapCompany {
  id: string;
  mapKind: ChaosMapKind;
  categoryId: string;
  name: string;
  difficulty: ChaosMapDifficulty;
  isNg: boolean;
  badgeIds: string[];
  memo?: string;
  createdBy: string;
  createdAt: string;
  updatedBy?: string;
  updatedAt?: string;
}

export interface ChaosMapCategoryDef {
  id: string;
  mapKind: ChaosMapKind;
  label: string;
  order: number;
}

export interface ChaosMapConfig {
  schemaVersion: number;
  categories: ChaosMapCategoryDef[];
  badgeCatalog: ChaosMapBadge[];
  companies: ChaosMapCompany[];
}

export function generateChaosMapId(prefix: string): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

// 参考資料「Firm+ map20250219_03 紹介先マップ.pdf」の2枚の列構成をそのままシードにする。
// 元PDFの「セールス」列内の縦積みサブグループ（大手事業会社/M&A仲介/コンサルファーム/SaaS）は
// MVPでは4本の独立列として展開する（ネスト表示は後続フェーズ）。
export const SEED_CHAOS_MAP_CATEGORIES: ChaosMapCategoryDef[] = [
  { id: 'firm_pe_fund', mapKind: 'firm', label: 'PEファンド/事業投資', order: 0 },
  { id: 'firm_strategy', mapKind: 'firm', label: '戦略ファーム', order: 1 },
  { id: 'firm_general', mapKind: 'firm', label: '総合ファーム', order: 2 },
  { id: 'firm_biz_dev', mapKind: 'firm', label: '事業開発/経営企画', order: 3 },
  { id: 'firm_marketing', mapKind: 'firm', label: 'マーケティング', order: 4 },
  { id: 'firm_sales_enterprise', mapKind: 'firm', label: 'セールス(大手事業会社)', order: 5 },
  { id: 'firm_sales_ma', mapKind: 'firm', label: 'セールス(M&A仲介)', order: 6 },
  { id: 'firm_sales_consulting', mapKind: 'firm', label: 'セールス(コンサルファーム)', order: 7 },
  { id: 'firm_sales_saas', mapKind: 'firm', label: 'セールス(SaaS)', order: 8 },
  { id: 'it_it_consulting', mapKind: 'it', label: 'ITコン', order: 0 },
  { id: 'it_ai_consulting', mapKind: 'it', label: 'AIコン', order: 1 },
  { id: 'it_saas_pdm', mapKind: 'it', label: 'SaaS(PDM/SE)', order: 2 },
  { id: 'it_sier', mapKind: 'it', label: 'SIer', order: 3 },
  { id: 'it_ses', mapKind: 'it', label: 'SES', order: 4 },
  { id: 'it_johosys', mapKind: 'it', label: '情シス(事業会社)', order: 5 },
];

export const SEED_CHAOS_MAP_BADGES: ChaosMapBadge[] = [
  { id: 'retirement65', emoji: '🌿', label: '定年が65歳' },
  { id: 'hiring50s', emoji: '🔥', label: '50代でも採用可能性有' },
];

export const SEED_CHAOS_MAP_CONFIG: ChaosMapConfig = {
  schemaVersion: 1,
  categories: SEED_CHAOS_MAP_CATEGORIES,
  badgeCatalog: SEED_CHAOS_MAP_BADGES,
  companies: [],
};
