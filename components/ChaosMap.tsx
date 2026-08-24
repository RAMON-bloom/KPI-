import React, { useMemo, useState } from 'react';
import {
  ChaosMapBadge,
  ChaosMapCategoryDef,
  ChaosMapCompany,
  ChaosMapConfig,
  ChaosMapDifficulty,
  ChaosMapKind,
  CHAOS_MAP_DIFFICULTIES,
  CHAOS_MAP_KIND_LABELS,
} from '../services/chaosMapData';

interface CompanyDraft {
  name: string;
  categoryId: string;
  difficulty: ChaosMapDifficulty;
  isNg: boolean;
  badgeIds: string[];
  memo: string;
}

function emptyDraft(categoryId: string, difficulty: ChaosMapDifficulty): CompanyDraft {
  return { name: '', categoryId, difficulty, isNg: false, badgeIds: [], memo: '' };
}

function draftFromCompany(company: ChaosMapCompany): CompanyDraft {
  return {
    name: company.name,
    categoryId: company.categoryId,
    difficulty: company.difficulty,
    isNg: company.isNg,
    badgeIds: company.badgeIds || [],
    memo: company.memo || '',
  };
}

interface CompanyModalProps {
  mapKind: ChaosMapKind;
  categories: ChaosMapCategoryDef[];
  badgeCatalog: ChaosMapBadge[];
  initialDraft: CompanyDraft;
  isEditing: boolean;
  onSave: (draft: CompanyDraft) => void;
  onDelete?: () => void;
  onClose: () => void;
}

const CompanyModal: React.FC<CompanyModalProps> = ({ mapKind, categories, badgeCatalog, initialDraft, isEditing, onSave, onDelete, onClose }) => {
  const [draft, setDraft] = useState<CompanyDraft>(initialDraft);

  const toggleBadge = (badgeId: string) => {
    setDraft(prev => ({
      ...prev,
      badgeIds: prev.badgeIds.includes(badgeId) ? prev.badgeIds.filter(id => id !== badgeId) : [...prev.badgeIds, badgeId],
    }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!draft.name.trim()) return;
    onSave({ ...draft, name: draft.name.trim(), memo: draft.memo.trim() });
  };

  return (
    <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="chaos-map-company-modal-title">
      <div className="modal-content" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h3 id="chaos-map-company-modal-title">{isEditing ? '企業を編集' : `企業を追加（${CHAOS_MAP_KIND_LABELS[mapKind]}）`}</h3>
          <button onClick={onClose} className="close-button" aria-label="閉じる">&times;</button>
        </div>
        <form onSubmit={handleSubmit} className="modal-body">
          <div className="form-group">
            <label htmlFor="chaos-map-company-name">企業名</label>
            <input
              id="chaos-map-company-name"
              type="text"
              value={draft.name}
              onChange={e => setDraft(prev => ({ ...prev, name: e.target.value }))}
              placeholder="例: 〇〇コンサルティング株式会社"
              autoFocus
            />
          </div>
          <div className="form-group">
            <label htmlFor="chaos-map-company-category">カテゴリ</label>
            <select
              id="chaos-map-company-category"
              value={draft.categoryId}
              onChange={e => setDraft(prev => ({ ...prev, categoryId: e.target.value }))}
            >
              {categories.map(cat => (
                <option key={cat.id} value={cat.id}>{cat.label}</option>
              ))}
            </select>
          </div>
          <div className="form-group">
            <label htmlFor="chaos-map-company-difficulty">難易度</label>
            <select
              id="chaos-map-company-difficulty"
              value={draft.difficulty}
              onChange={e => setDraft(prev => ({ ...prev, difficulty: e.target.value as ChaosMapDifficulty }))}
            >
              {CHAOS_MAP_DIFFICULTIES.map(d => (
                <option key={d} value={d}>{d}</option>
              ))}
            </select>
          </div>
          <div className="form-group">
            <label>
              <input
                type="checkbox"
                checked={draft.isNg}
                onChange={e => setDraft(prev => ({ ...prev, isNg: e.target.checked }))}
              />
              {' '}紹介不可（NG企業）
            </label>
          </div>
          {badgeCatalog.length > 0 && (
            <div className="form-group">
              <label>バッジ</label>
              <div className="chaos-map-badge-checkbox-group">
                {badgeCatalog.map(badge => (
                  <label key={badge.id} className="chaos-map-badge-checkbox">
                    <input
                      type="checkbox"
                      checked={draft.badgeIds.includes(badge.id)}
                      onChange={() => toggleBadge(badge.id)}
                    />
                    {' '}{badge.emoji} {badge.label}
                  </label>
                ))}
              </div>
            </div>
          )}
          <div className="form-group">
            <label htmlFor="chaos-map-company-memo">メモ</label>
            <textarea
              id="chaos-map-company-memo"
              value={draft.memo}
              onChange={e => setDraft(prev => ({ ...prev, memo: e.target.value }))}
              rows={3}
            />
          </div>
          <div className="chaos-map-modal-actions">
            {isEditing && onDelete && (
              <button
                type="button"
                className="delete-user-button"
                onClick={() => { if (window.confirm(`「${draft.name}」をカオスマップから削除します。よろしいですか？`)) onDelete(); }}
              >
                削除
              </button>
            )}
            <div style={{ flex: 1 }} />
            <button type="button" className="secondary-action-button" onClick={onClose}>キャンセル</button>
            <button type="submit" className="submit-button" disabled={!draft.name.trim()}>保存</button>
          </div>
        </form>
      </div>
    </div>
  );
};

interface BadgeCatalogModalProps {
  badgeCatalog: ChaosMapBadge[];
  onSave: (badges: ChaosMapBadge[]) => void;
  onClose: () => void;
}

const BadgeCatalogModal: React.FC<BadgeCatalogModalProps> = ({ badgeCatalog, onSave, onClose }) => {
  const [badges, setBadges] = useState<ChaosMapBadge[]>(badgeCatalog);
  const [newEmoji, setNewEmoji] = useState('');
  const [newLabel, setNewLabel] = useState('');

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newEmoji.trim() || !newLabel.trim()) return;
    const id = `badge_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
    setBadges(prev => [...prev, { id, emoji: newEmoji.trim(), label: newLabel.trim() }]);
    setNewEmoji('');
    setNewLabel('');
  };

  const handleRemove = (id: string) => setBadges(prev => prev.filter(b => b.id !== id));

  return (
    <div className="modal-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-labelledby="chaos-map-badge-modal-title">
      <div className="modal-content" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h3 id="chaos-map-badge-modal-title">バッジ管理</h3>
          <button onClick={onClose} className="close-button" aria-label="閉じる">&times;</button>
        </div>
        <div className="modal-body">
          <form onSubmit={handleAdd} className="form-group" style={{ display: 'flex', gap: '0.5rem', alignItems: 'flex-end', marginBottom: '1rem' }}>
            <div style={{ flex: '0 0 4rem' }}>
              <label htmlFor="chaos-map-new-badge-emoji">絵文字</label>
              <input id="chaos-map-new-badge-emoji" type="text" value={newEmoji} onChange={e => setNewEmoji(e.target.value)} placeholder="🌿" />
            </div>
            <div style={{ flex: 1 }}>
              <label htmlFor="chaos-map-new-badge-label">ラベル</label>
              <input id="chaos-map-new-badge-label" type="text" value={newLabel} onChange={e => setNewLabel(e.target.value)} placeholder="例: 定年が65歳" />
            </div>
            <button type="submit" className="submit-button" disabled={!newEmoji.trim() || !newLabel.trim()}>追加</button>
          </form>
          <ul className="user-management-list">
            {badges.map(badge => (
              <li key={badge.id} className="user-management-item">
                <span className="user-management-name">{badge.emoji} {badge.label}</span>
                <div className="user-management-actions">
                  <button onClick={() => handleRemove(badge.id)} className="delete-user-button">削除</button>
                </div>
              </li>
            ))}
            {badges.length === 0 && <p className="no-data-message">バッジがありません。</p>}
          </ul>
          <div className="chaos-map-modal-actions">
            <div style={{ flex: 1 }} />
            <button type="button" className="secondary-action-button" onClick={onClose}>キャンセル</button>
            <button type="button" className="submit-button" onClick={() => onSave(badges)}>保存</button>
          </div>
        </div>
      </div>
    </div>
  );
};

export interface ChaosMapViewProps {
  config: ChaosMapConfig | null;
  isLoading: boolean;
  onAddCompany: (company: Omit<ChaosMapCompany, 'id' | 'createdBy' | 'createdAt'>) => Promise<void>;
  onUpdateCompany: (id: string, patch: Partial<ChaosMapCompany>) => Promise<void>;
  onDeleteCompany: (id: string) => Promise<void>;
  onSaveBadgeCatalog: (badges: ChaosMapBadge[]) => Promise<void>;
}

export const ChaosMapView: React.FC<ChaosMapViewProps> = ({ config, isLoading, onAddCompany, onUpdateCompany, onDeleteCompany, onSaveBadgeCatalog }) => {
  const [activeMapKind, setActiveMapKind] = useState<ChaosMapKind>('firm');
  const [searchTerm, setSearchTerm] = useState('');
  const [showNg, setShowNg] = useState(false);
  const [editingCompany, setEditingCompany] = useState<ChaosMapCompany | null>(null);
  const [newCompanyPreset, setNewCompanyPreset] = useState<{ categoryId: string; difficulty: ChaosMapDifficulty } | null>(null);
  const [isBadgeModalOpen, setIsBadgeModalOpen] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  const categories = useMemo(
    () => (config?.categories || []).filter(c => c.mapKind === activeMapKind).sort((a, b) => a.order - b.order),
    [config, activeMapKind]
  );
  const badgeCatalog = config?.badgeCatalog || [];
  const badgeById = useMemo(() => new Map(badgeCatalog.map(b => [b.id, b])), [badgeCatalog]);

  const filteredCompanies = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();
    return (config?.companies || []).filter(c => {
      if (c.mapKind !== activeMapKind) return false;
      if (!showNg && c.isNg) return false;
      if (term && !c.name.toLowerCase().includes(term)) return false;
      return true;
    });
  }, [config, activeMapKind, showNg, searchTerm]);

  const companiesByCell = useMemo(() => {
    const map = new Map<string, ChaosMapCompany[]>();
    filteredCompanies.forEach(c => {
      const key = `${c.categoryId}::${c.difficulty}`;
      const arr = map.get(key) || [];
      arr.push(c);
      map.set(key, arr);
    });
    return map;
  }, [filteredCompanies]);

  const handleSaveNewCompany = async (draft: CompanyDraft) => {
    setIsSaving(true);
    try {
      await onAddCompany({ mapKind: activeMapKind, ...draft });
      setNewCompanyPreset(null);
    } finally {
      setIsSaving(false);
    }
  };

  const handleSaveEditedCompany = async (draft: CompanyDraft) => {
    if (!editingCompany) return;
    setIsSaving(true);
    try {
      await onUpdateCompany(editingCompany.id, draft);
      setEditingCompany(null);
    } finally {
      setIsSaving(false);
    }
  };

  const handleDeleteCompany = async () => {
    if (!editingCompany) return;
    setIsSaving(true);
    try {
      await onDeleteCompany(editingCompany.id);
      setEditingCompany(null);
    } finally {
      setIsSaving(false);
    }
  };

  const handleSaveBadges = async (badges: ChaosMapBadge[]) => {
    setIsSaving(true);
    try {
      await onSaveBadgeCatalog(badges);
      setIsBadgeModalOpen(false);
    } finally {
      setIsSaving(false);
    }
  };

  if (isLoading) {
    return <p className="no-data-message">紹介先カオスマップを読み込み中...</p>;
  }

  return (
    <div className="chaos-map-view">
      <div className="chaos-map-toolbar">
        <div className="chaos-map-kind-switcher" role="group" aria-label="マップ種別切り替え">
          {(Object.keys(CHAOS_MAP_KIND_LABELS) as ChaosMapKind[]).map(kind => (
            <button key={kind} onClick={() => setActiveMapKind(kind)} disabled={activeMapKind === kind}>
              {CHAOS_MAP_KIND_LABELS[kind]}
            </button>
          ))}
        </div>
        <input
          type="text"
          className="chaos-map-search-input"
          placeholder="企業名で検索"
          value={searchTerm}
          onChange={e => setSearchTerm(e.target.value)}
        />
        <label className="chaos-map-ng-toggle">
          <input type="checkbox" checked={showNg} onChange={e => setShowNg(e.target.checked)} />
          {' '}NG企業を表示
        </label>
        <button type="button" className="secondary-action-button" onClick={() => setIsBadgeModalOpen(true)}>バッジ管理</button>
        <button
          type="button"
          className="submit-button"
          onClick={() => setNewCompanyPreset({ categoryId: categories[0]?.id || '', difficulty: 'B' })}
          disabled={categories.length === 0}
        >
          ＋企業を追加
        </button>
      </div>

      {categories.length === 0 ? (
        <p className="no-data-message">このマップにはまだカテゴリがありません。</p>
      ) : (
        <div className="chaos-map-grid" style={{ gridTemplateColumns: `auto repeat(${categories.length}, 1fr)` }}>
          <div className="chaos-map-corner-cell" />
          {categories.map(cat => (
            <div key={cat.id} className="chaos-map-category-header">{cat.label}</div>
          ))}
          {CHAOS_MAP_DIFFICULTIES.map(difficulty => (
            <React.Fragment key={difficulty}>
              <div className={`chaos-map-difficulty-header chaos-map-difficulty-${difficulty}`}>{difficulty}</div>
              {categories.map(cat => {
                const key = `${cat.id}::${difficulty}`;
                const companies = companiesByCell.get(key) || [];
                return (
                  <div key={key} className="chaos-map-cell">
                    {companies.map(company => (
                      <button
                        key={company.id}
                        type="button"
                        className={`chaos-map-chip${company.isNg ? ' chaos-map-chip-ng' : ''}`}
                        onClick={() => setEditingCompany(company)}
                        title={company.memo || undefined}
                      >
                        {company.name}
                        {(company.badgeIds || []).map(badgeId => {
                          const badge = badgeById.get(badgeId);
                          return badge ? <span key={badgeId} className="chaos-map-chip-badge">{badge.emoji}</span> : null;
                        })}
                        {company.isNg && <span className="chaos-map-chip-ng-label">NG</span>}
                      </button>
                    ))}
                    <button
                      type="button"
                      className="chaos-map-cell-add-button"
                      onClick={() => setNewCompanyPreset({ categoryId: cat.id, difficulty })}
                      aria-label={`${cat.label}・難易度${difficulty}に企業を追加`}
                    >
                      ＋
                    </button>
                  </div>
                );
              })}
            </React.Fragment>
          ))}
        </div>
      )}

      {newCompanyPreset && (
        <CompanyModal
          mapKind={activeMapKind}
          categories={categories}
          badgeCatalog={badgeCatalog}
          initialDraft={emptyDraft(newCompanyPreset.categoryId, newCompanyPreset.difficulty)}
          isEditing={false}
          onSave={handleSaveNewCompany}
          onClose={() => setNewCompanyPreset(null)}
        />
      )}
      {editingCompany && (
        <CompanyModal
          mapKind={editingCompany.mapKind}
          categories={(config?.categories || []).filter(c => c.mapKind === editingCompany.mapKind).sort((a, b) => a.order - b.order)}
          badgeCatalog={badgeCatalog}
          initialDraft={draftFromCompany(editingCompany)}
          isEditing
          onSave={handleSaveEditedCompany}
          onDelete={handleDeleteCompany}
          onClose={() => setEditingCompany(null)}
        />
      )}
      {isBadgeModalOpen && (
        <BadgeCatalogModal
          badgeCatalog={badgeCatalog}
          onSave={handleSaveBadges}
          onClose={() => setIsBadgeModalOpen(false)}
        />
      )}
      {isSaving && <p className="chaos-map-saving-indicator">保存中...</p>}
    </div>
  );
};
