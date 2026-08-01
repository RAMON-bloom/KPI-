import { findOwnDataFile, readFileContent, createOwnDataFile, updateFileContent, listTeammateDataFiles, findTeamsConfigFile, createTeamsConfigFile, findMediaConfigFile, createMediaConfigFile, ensureDomainPermission, listPermissions, grantIndividualPermission, revokePermission } from './googleDrive';

const LOCAL_CACHE_PREFIX = 'kpiUserDataCache:';
const DRIVE_FILE_ID_CACHE_PREFIX = 'kpiDriveFileId:';
const PENDING_SYNC_PREFIX = 'kpiPendingSync:';
const LAST_SYNCED_AT_PREFIX = 'kpiLastSyncedAt:';
const LEGACY_APPDATA_KEY = 'kpiAppData';
const MEDIA_CONFIG_CACHE_KEY = 'kpiMediaConfigCache';
const SCHEMA_VERSION = 1;

function pendingSyncKey(email: string): string {
  return `${PENDING_SYNC_PREFIX}${email}`;
}

function lastSyncedAtKey(email: string): string {
  return `${LAST_SYNCED_AT_PREFIX}${email}`;
}

/** Epoch ms of the last successful Drive write for this user, or null if there's never been one. */
export function getLastSyncedAt(email: string): number | null {
  try {
    const raw = localStorage.getItem(lastSyncedAtKey(email));
    return raw ? Number(raw) : null;
  } catch {
    return null;
  }
}

function setLastSyncedAt(email: string, timestamp: number): void {
  try {
    localStorage.setItem(lastSyncedAtKey(email), String(timestamp));
  } catch {
    // ignore
  }
}

/**
 * Whether the most recent Drive write for this user failed (e.g. the Google session expired
 * or was revoked right as the debounced save fired) and hasn't been successfully retried yet.
 * The data itself isn't lost — it's still in this browser's local cache — but it never reached
 * Drive, so other devices/sessions won't see it until this resolves.
 */
export function hasPendingSync(email: string): boolean {
  try {
    return localStorage.getItem(pendingSyncKey(email)) === '1';
  } catch {
    return false;
  }
}

function markPendingSync(email: string): void {
  try {
    localStorage.setItem(pendingSyncKey(email), '1');
  } catch {
    // ignore
  }
}

function clearPendingSync(email: string): void {
  try {
    localStorage.removeItem(pendingSyncKey(email));
  } catch {
    // ignore
  }
}

type SyncStatusListener = (email: string, hasPending: boolean) => void;
const syncStatusListeners = new Set<SyncStatusListener>();

/**
 * Notified immediately after every save attempt (success or failure) for the given email —
 * lets the UI show a warning right when a KPI/pipeline entry actually fails to sync, instead
 * of only finding out on the next periodic check up to a minute later. Returns an unsubscribe
 * function.
 */
export function onSyncStatusChange(listener: SyncStatusListener): () => void {
  syncStatusListeners.add(listener);
  return () => { syncStatusListeners.delete(listener); };
}

function notifySyncStatus(email: string, hasPending: boolean): void {
  syncStatusListeners.forEach(listener => listener(email, hasPending));
}

/** The shared media-config cache is not user-specific — everyone reads the same list. */
export function readMediaConfigCache<T = any>(): T | null {
  try {
    const raw = localStorage.getItem(MEDIA_CONFIG_CACHE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function writeMediaConfigCache(data: unknown): void {
  try {
    localStorage.setItem(MEDIA_CONFIG_CACHE_KEY, JSON.stringify(data));
  } catch {
    // ignore
  }
}

function cacheKey(email: string): string {
  return `${LOCAL_CACHE_PREFIX}${email}`;
}

function driveFileIdCacheKey(email: string): string {
  return `${DRIVE_FILE_ID_CACHE_PREFIX}${email}`;
}

function getCachedDriveFileId(email: string): string | null {
  try {
    return localStorage.getItem(driveFileIdCacheKey(email));
  } catch {
    return null;
  }
}

function setCachedDriveFileId(email: string, fileId: string): void {
  try {
    localStorage.setItem(driveFileIdCacheKey(email), fileId);
  } catch {
    // ignore
  }
}

export function readLocalCache<T = any>(email: string): T | null {
  try {
    const raw = localStorage.getItem(cacheKey(email));
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

export function writeLocalCache(email: string, data: unknown): void {
  try {
    localStorage.setItem(cacheKey(email), JSON.stringify(data));
  } catch (err) {
    console.error('Failed to write local cache', err);
  }
}

/** The pre-Google-login localStorage blob, kept around only to power the one-time migration prompt. */
export function readLegacyAppData(): { users: string[]; userData: Record<string, any> } | null {
  try {
    const raw = localStorage.getItem(LEGACY_APPDATA_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

export interface LoadResult<T> {
  data: T | null;
  driveFileId: string | null;
  source: 'drive' | 'cache' | 'new';
}

/**
 * Loads the signed-in user's own data: Drive is the source of truth, local cache is a fallback.
 * Skips the `files.list` search (one Drive round-trip) whenever we already know the file id from
 * a previous session, falling back to a full search only if that cached id turns out to be stale.
 */
export async function loadOwnData<T = any>(email: string): Promise<LoadResult<T>> {
  try {
    const cachedId = getCachedDriveFileId(email);
    if (cachedId) {
      try {
        const content = await readFileContent<T>(cachedId);
        writeLocalCache(email, content);
        // Best-effort self-heal: re-grant domain sharing if a prior save silently failed to
        // set it (this is why a teammate's data could be invisible in the all-users view).
        ensureDomainPermission(cachedId, 'reader').catch(() => {});
        return { data: content, driveFileId: cachedId, source: 'drive' };
      } catch (err) {
        console.warn('Cached Drive file id is stale, falling back to a full search', err);
      }
    }
    const existing = await findOwnDataFile();
    if (existing) {
      const content = await readFileContent<T>(existing.id);
      writeLocalCache(email, content);
      setCachedDriveFileId(email, existing.id);
      ensureDomainPermission(existing.id, 'reader').catch(() => {});
      return { data: content, driveFileId: existing.id, source: 'drive' };
    }
  } catch (err) {
    console.error('Failed to load from Drive, falling back to local cache', err);
    const cached = readLocalCache<T>(email);
    if (cached) return { data: cached, driveFileId: null, source: 'cache' };
  }
  return { data: null, driveFileId: null, source: 'new' };
}

export async function createInitialDriveFile(email: string, data: unknown): Promise<string> {
  const payload = { ...(data as object), schemaVersion: SCHEMA_VERSION };
  const fileId = await createOwnDataFile(payload, email);
  writeLocalCache(email, payload);
  setCachedDriveFileId(email, fileId);
  return fileId;
}

let debounceTimer: ReturnType<typeof setTimeout> | null = null;
let pendingSave: { email: string; driveFileId: string | null; data: unknown; onFileCreated: (id: string) => void } | null = null;
// A promise chain used as a mutex: every write is appended with .then() so it can only start
// once the previous one has fully finished. Without this, two saves fired close together (e.g.
// two debounce cycles back to back) could both be in flight at once, and if the network
// reorders their responses, the OLDER request's response can land after the newer one and
// silently overwrite it with stale content — even though the newer save "succeeded" from the
// caller's point of view. This is why data entered well outside the 2s debounce window could
// still vanish: it wasn't a debounce-timing race, it was an in-flight write race.
let writeQueue: Promise<void> = Promise.resolve();

async function performSave(
  email: string,
  driveFileId: string | null,
  data: unknown,
  onFileCreated: (id: string) => void
): Promise<void> {
  try {
    const payload = { ...(data as object), schemaVersion: SCHEMA_VERSION };
    if (driveFileId) {
      await updateFileContent(driveFileId, payload);
    } else {
      const newId = await createOwnDataFile(payload, email);
      setCachedDriveFileId(email, newId);
      onFileCreated(newId);
    }
    clearPendingSync(email);
    setLastSyncedAt(email, Date.now());
    notifySyncStatus(email, false);
  } catch (err) {
    // The data is still safe in this browser's local cache — just flag that Drive hasn't seen
    // it yet (e.g. the Google session expired/was revoked right as this fired) so it can be
    // retried once a valid session is available again (see retryPendingSyncIfNeeded).
    console.error('Failed to sync data to Drive', err);
    markPendingSync(email);
    notifySyncStatus(email, true);
  }
}

/** Appends the currently-pending save (if any) to the write queue and returns it. */
function enqueuePendingSave(): Promise<void> {
  writeQueue = writeQueue.then(async () => {
    const save = pendingSave;
    pendingSave = null;
    if (save) await performSave(save.email, save.driveFileId, save.data, save.onFileCreated);
  });
  return writeQueue;
}

/**
 * Writes to the local cache immediately (fast UI), then debounces the Drive sync
 * (~2s idle) so rapid KPI-entry keystrokes don't hammer the Drive API.
 */
export function saveOwnDataDebounced(
  email: string,
  driveFileId: string | null,
  data: unknown,
  onFileCreated: (id: string) => void
): void {
  writeLocalCache(email, data);
  pendingSave = { email, driveFileId, data, onFileCreated };
  if (debounceTimer) clearTimeout(debounceTimer);
  debounceTimer = setTimeout(() => {
    debounceTimer = null;
    enqueuePendingSave();
  }, 2000);
}

/**
 * Immediately performs any pending debounced save instead of waiting out the idle timer, and
 * waits for it (and anything already ahead of it in the write queue) to actually finish. Call
 * this before signing out and on visibility/pagehide changes — without it, a save queued right
 * before the user signs out could be lost: the timer either never fires (page/context gone) or
 * fires after the session's already cleared and the write fails silently.
 */
export async function flushPendingSave(): Promise<void> {
  if (debounceTimer) {
    clearTimeout(debounceTimer);
    debounceTimer = null;
  }
  await enqueuePendingSave();
}

/**
 * Pushes the given data to Drive right now, regardless of whether anything is actually
 * pending — for a manual "同期" button, where the user wants to confirm/force a write rather
 * than wait for the debounce or a change to trigger one.
 */
export async function forceSyncNow(
  email: string,
  driveFileId: string | null,
  data: unknown,
  onFileCreated: (id: string) => void
): Promise<void> {
  if (debounceTimer) {
    clearTimeout(debounceTimer);
    debounceTimer = null;
  }
  writeLocalCache(email, data);
  pendingSave = { email, driveFileId, data, onFileCreated };
  await enqueuePendingSave();
}

/**
 * If an earlier save for this user failed (see markPendingSync above — most commonly an
 * expired/revoked Google session right as the debounced write fired), re-attempts it using
 * whatever is currently in local cache, the most complete copy of the user's data we have.
 * Call this whenever the app confirms it has a valid session (initial load, and periodically
 * while the tab stays open) so a save that failed once doesn't stay lost until the user
 * happens to change something again.
 */
export async function retryPendingSyncIfNeeded(
  email: string,
  driveFileId: string | null,
  onFileCreated: (id: string) => void
): Promise<void> {
  if (!hasPendingSync(email)) return;
  const cached = readLocalCache(email);
  if (!cached) return;
  pendingSave = { email, driveFileId, data: cached, onFileCreated };
  await enqueuePendingSave();
}

export interface TeammateData<T> {
  email: string;
  data: T;
  driveFileId: string;
}

/** Fetches every teammate's domain-shared data file (used by the cross-user / team views). */
export async function loadAllTeammatesData<T = any>(): Promise<TeammateData<T>[]> {
  const files = await listTeammateDataFiles();
  const results = await Promise.all(
    files.map(async (file) => {
      try {
        const content = await readFileContent<T>(file.id);
        return { email: file.ownerEmail || file.name, data: content, driveFileId: file.id };
      } catch (err) {
        console.error(`Failed to read teammate file ${file.id}`, err);
        return null;
      }
    })
  );
  return results.filter((r) => r !== null) as TeammateData<T>[];
}

/**
 * Reconciles who (beyond the file's owner) has direct WRITE access to a personal data file, down
 * to an exact desired set of individual accounts — used to grant ミドル proxy-entry write access
 * to exactly the teammates who are currently eligible (share a team with this file's owner and
 * hold the ミドル role), and to revoke it again the moment that's no longer true. Only ever
 * touches `type: 'user'` permissions — the file's pre-existing domain-wide reader grant (`type:
 * 'domain'`, used for the 全ユーザー/チーム別 read-only aggregation) is untouched.
 */
export async function syncIndividualWriterPermissions(fileId: string, desiredEmails: string[]): Promise<void> {
  const permissions = await listPermissions(fileId);
  const currentWriterEmails = permissions.filter(p => p.type === 'user' && p.role === 'writer' && p.emailAddress);
  const toRevoke = currentWriterEmails.filter(p => !desiredEmails.includes(p.emailAddress!));
  const toGrant = desiredEmails.filter(email => !currentWriterEmails.some(p => p.emailAddress === email));
  await Promise.all([
    ...toRevoke.map(p => revokePermission(fileId, p.id).catch(err => console.error(`Failed to revoke Drive permission for ${p.emailAddress}`, err))),
    ...toGrant.map(email => grantIndividualPermission(fileId, email, 'writer').catch(err => console.error(`Failed to grant Drive permission to ${email}`, err))),
  ]);
}

/**
 * ミドルによる代理入力: fetches the target teammate's LATEST data (not whatever stale copy may
 * be cached in memory — someone else may have saved in the meantime), replaces just that one
 * date's entry with the given values (same full-replace semantics as the teammate's own
 * DateEntryModal save), and writes the merged result back. Requires this fileId to already have
 * been granted direct 'writer' access to the signed-in account (see
 * syncIndividualWriterPermissions) — otherwise the underlying Drive write fails with 403.
 */
export async function overwriteTeammateEntry<T extends { entries: any[] } = any>(
  driveFileId: string,
  date: string,
  values: unknown
): Promise<T> {
  const latest = await readFileContent<T>(driveFileId);
  const otherEntries = (latest.entries || []).filter((e: any) => e.date !== date);
  const newEntry = { id: Date.now(), date, values };
  const updatedEntries = [...otherEntries, newEntry].sort((a: any, b: any) => a.date.localeCompare(b.date));
  const updated = { ...latest, entries: updatedEntries };
  await updateFileContent(driveFileId, updated);
  return updated;
}

/**
 * スプレッドシート取込み（チームメンバー分も含めて取込む場合）専用: overwriteTeammateEntry と
 * 同じ「対象フィールドだけ上書き、他は変更しない」マージだが、日付1件ずつではなく
 * `countsByDate`（複数日分）をまとめて1回のfetch+writeで反映する。取込みは数十〜数百日分に
 * 及ぶことがあり、日付ごとに逐次overwriteTeammateEntryを呼ぶとDrive API呼び出しが同数発生して
 * 遅く・失敗しやすくなるため、対象メンバー1人につき1回のDrive書き込みで済ませる。
 */
export async function overwriteTeammateEntries<T extends { entries: any[] } = any>(
  driveFileId: string,
  countsByDate: Record<string, Record<string, number>>
): Promise<T> {
  const latest = await readFileContent<T>(driveFileId);
  const entriesByDateMap = new Map<string, any>((latest.entries || []).map((e: any) => [e.date, e]));
  Object.entries(countsByDate).forEach(([dateStr, counts]) => {
    const existing = entriesByDateMap.get(dateStr);
    const values: Record<string, number> = existing ? { ...existing.values } : {};
    Object.entries(counts).forEach(([key, count]) => { values[key] = count; });
    entriesByDateMap.set(dateStr, { id: existing?.id ?? Date.now(), date: dateStr, values });
  });
  const updatedEntries = Array.from(entriesByDateMap.values()).sort((a: any, b: any) => a.date.localeCompare(b.date));
  const updated = { ...latest, entries: updatedEntries };
  await updateFileContent(driveFileId, updated);
  return updated;
}

/**
 * チーム作成・編集権限保持者による代理での候補者「非表示」切り替え: fetches the target
 * teammate's LATEST data (same re-fetch-before-write pattern as overwriteTeammateEntry, to avoid
 * clobbering a concurrent save), flips isHidden on just the matching candidate, and writes the
 * merged result back. Requires this fileId to already have been granted direct 'writer' access to
 * the signed-in account (see syncIndividualWriterPermissions) — otherwise the underlying Drive
 * write fails with 403.
 */
export async function overwriteTeammateCandidateVisibility<T extends { candidates: any[] } = any>(
  driveFileId: string,
  candidateId: string,
  nextIsHidden: boolean
): Promise<T> {
  const latest = await readFileContent<T>(driveFileId);
  const updatedCandidates = (latest.candidates || []).map((c: any) =>
    c.id === candidateId ? { ...c, isHidden: nextIsHidden } : c
  );
  const updated = { ...latest, candidates: updatedCandidates };
  await updateFileContent(driveFileId, updated);
  return updated;
}

/**
 * ミドルによる代理での候補者パイプライン全項目編集: fetches the target teammate's LATEST data
 * (same re-fetch-before-write pattern as the other overwriteTeammate* functions above), re-derives
 * the patched candidate AND any KPI actuals it triggers (recommendationRecorded/exitRecorded/
 * stage-advance deltas) from that FRESH copy — not from whatever the ミドル's local aggregate cache
 * last saw — merges both back in, and writes the result. `patch` is only the fields the ミドル
 * actually changed (never the whole candidate object), so a stale local cache can only ever
 * overwrite the handful of fields genuinely edited, not clobber the rest of the candidate.
 * `computeUpdate` is the caller's computeStageAdvanceUpdate, injected so this data-layer module
 * doesn't need to know about KPI business rules. Requires this fileId to already have been granted
 * direct 'writer' access to the signed-in account (see syncIndividualWriterPermissions) —
 * otherwise the underlying Drive write fails with 403.
 */
export async function overwriteTeammateCandidatePatch<T extends { candidates: any[]; entries: any[] } = any>(
  driveFileId: string,
  candidateId: string,
  patch: Record<string, any>,
  computeUpdate: (prevCandidate: any, nextCandidate: any, todayStr: string) => { candidate: any; kpiDeltas: Record<string, number> },
  todayStr: string
): Promise<T> {
  const latest = await readFileContent<T>(driveFileId);
  const candidates = latest.candidates || [];
  const prevCandidate = candidates.find((c: any) => c.id === candidateId);
  if (!prevCandidate) throw new Error('対象の候補者が見つかりませんでした（既に削除されている可能性があります）。');
  const patchedCandidate = { ...prevCandidate, ...patch };
  const { candidate: finalCandidate, kpiDeltas } = computeUpdate(prevCandidate, patchedCandidate, todayStr);
  const updatedCandidates = candidates.map((c: any) => (c.id === candidateId ? finalCandidate : c));
  let updatedEntries = latest.entries || [];
  if (Object.keys(kpiDeltas).length > 0) {
    const entriesByDateMap = new Map<string, any>(updatedEntries.map((e: any) => [e.date, e]));
    const existingEntry = entriesByDateMap.get(todayStr);
    const values: Record<string, number> = existingEntry ? { ...existingEntry.values } : {};
    Object.entries(kpiDeltas).forEach(([key, delta]) => { values[key] = (values[key] || 0) + delta; });
    entriesByDateMap.set(todayStr, { id: existingEntry?.id ?? Date.now(), date: todayStr, values });
    updatedEntries = Array.from(entriesByDateMap.values()).sort((a: any, b: any) => a.date.localeCompare(b.date));
  }
  const updated = { ...latest, candidates: updatedCandidates, entries: updatedEntries };
  await updateFileContent(driveFileId, updated);
  return updated;
}

/**
 * ミドルによる代理での候補者新規登録: fetches the target teammate's LATEST data (same
 * re-fetch-before-write pattern as overwriteTeammateCandidatePatch above), computes the new
 * candidate AND any KPI actuals its initial state triggers (via the caller's
 * computeStageAdvanceUpdate, called with `undefined` as the previous candidate — same "brand
 * new" path the候補者本人's own registration flow uses) against that FRESH copy, appends it to
 * `candidates`, merges the KPI deltas into `entries`, and writes the result back. Requires this
 * fileId to already have been granted direct 'writer' access to the signed-in account (see
 * syncIndividualWriterPermissions) — otherwise the underlying Drive write fails with 403.
 */
export async function addTeammateCandidate<T extends { candidates: any[]; entries: any[] } = any>(
  driveFileId: string,
  candidateData: Record<string, any>,
  computeUpdate: (prevCandidate: any, nextCandidate: any, todayStr: string) => { candidate: any; kpiDeltas: Record<string, number> },
  todayStr: string
): Promise<T> {
  const latest = await readFileContent<T>(driveFileId);
  const { candidate: finalCandidate, kpiDeltas } = computeUpdate(undefined, candidateData, todayStr);
  const updatedCandidates = [...(latest.candidates || []), finalCandidate];
  let updatedEntries = latest.entries || [];
  if (Object.keys(kpiDeltas).length > 0) {
    const entriesByDateMap = new Map<string, any>(updatedEntries.map((e: any) => [e.date, e]));
    const existingEntry = entriesByDateMap.get(todayStr);
    const values: Record<string, number> = existingEntry ? { ...existingEntry.values } : {};
    Object.entries(kpiDeltas).forEach(([key, delta]) => { values[key] = (values[key] || 0) + delta; });
    entriesByDateMap.set(todayStr, { id: existingEntry?.id ?? Date.now(), date: todayStr, values });
    updatedEntries = Array.from(entriesByDateMap.values()).sort((a: any, b: any) => a.date.localeCompare(b.date));
  }
  const updated = { ...latest, candidates: updatedCandidates, entries: updatedEntries };
  await updateFileContent(driveFileId, updated);
  return updated;
}

/**
 * 候補者パイプラインのスプレッドシート取込み（チームメンバー分も含めて取込む場合）専用:
 * addTeammateCandidate と同じ「新規登録」の扱いだが、1件ずつ都度fetch+writeするのではなく
 * `candidatesData`（取込みで生成された複数件の新規候補者）をまとめて1回のfetch+writeで反映する。
 * 1件ずつ呼ぶと後続の書き込みが直前の書き込みを再取得する前に走り、先に追加した候補者が消えて
 * しまう競合が起きうるため、対象メンバー1人につき1回のDrive書き込みで済ませる。
 */
export async function addTeammateCandidatesBulk<T extends { candidates: any[]; entries: any[] } = any>(
  driveFileId: string,
  candidatesData: Record<string, any>[],
  computeUpdate: (prevCandidate: any, nextCandidate: any, todayStr: string) => { candidate: any; kpiDeltas: Record<string, number> },
  todayStr: string
): Promise<T> {
  const latest = await readFileContent<T>(driveFileId);
  const updatedCandidates = [...(latest.candidates || [])];
  let updatedEntries = latest.entries || [];
  const entriesByDateMap = new Map<string, any>(updatedEntries.map((e: any) => [e.date, e]));
  candidatesData.forEach(candidateData => {
    const { candidate: finalCandidate, kpiDeltas } = computeUpdate(undefined, candidateData, todayStr);
    updatedCandidates.push(finalCandidate);
    if (Object.keys(kpiDeltas).length > 0) {
      const existingEntry = entriesByDateMap.get(todayStr);
      const values: Record<string, number> = existingEntry ? { ...existingEntry.values } : {};
      Object.entries(kpiDeltas).forEach(([key, delta]) => { values[key] = (values[key] || 0) + delta; });
      entriesByDateMap.set(todayStr, { id: existingEntry?.id ?? Date.now(), date: todayStr, values });
    }
  });
  updatedEntries = Array.from(entriesByDateMap.values()).sort((a: any, b: any) => a.date.localeCompare(b.date));
  const updated = { ...latest, candidates: updatedCandidates, entries: updatedEntries };
  await updateFileContent(driveFileId, updated);
  return updated;
}

/**
 * 開発者（TEAMS_ADMIN_EMAIL）による、他ユーザーの「お問い合わせ」投稿への返信・ステータス
 * 変更・削除: fetches the target user's LATEST data (same re-fetch-before-write pattern as
 * overwriteTeammateEntry/overwriteTeammateCandidateVisibility, to avoid clobbering a concurrent
 * save), then either merges `patch` into the matching post or, if `patch` is null, removes it.
 * Requires this fileId to already have been granted direct 'writer' access to the signed-in
 * account (see syncIndividualWriterPermissions) — otherwise the underlying Drive write fails
 * with 403.
 */
export async function overwriteTeammateFeedbackPost<T extends { feedbackPosts?: any[] } = any>(
  driveFileId: string,
  postId: string,
  patch: Record<string, unknown> | null
): Promise<T> {
  const latest = await readFileContent<T>(driveFileId);
  const posts = latest.feedbackPosts || [];
  const updatedPosts = patch === null
    ? posts.filter((p: any) => p.id !== postId)
    : posts.map((p: any) => (p.id === postId ? { ...p, ...patch } : p));
  const updated = { ...latest, feedbackPosts: updatedPosts };
  await updateFileContent(driveFileId, updated);
  return updated;
}

export interface TeamsConfigResult<T> {
  data: T | null;
  driveFileId: string | null;
  ownerEmail: string | null;
}

/** Loads the single shared teams-config file, if one has been created yet. */
export async function loadTeamsConfig<T = any>(): Promise<TeamsConfigResult<T>> {
  const existing = await findTeamsConfigFile();
  if (!existing) return { data: null, driveFileId: null, ownerEmail: null };
  const content = await readFileContent<T>(existing.id);
  // Only succeeds when the loading user is the file's owner (drive.file scope); harmlessly
  // fails otherwise. Self-heals sharing if the owner's own client failed to set it up before.
  ensureDomainPermission(existing.id, 'writer').catch(() => {});
  return { data: content, driveFileId: existing.id, ownerEmail: existing.ownerEmail ?? null };
}

/**
 * Creates the shared teams-config file the first time anyone sets up a team, or updates it
 * if it already exists. Only the original creator's browser can successfully update it
 * (drive.file scope only grants write access to files this app instance created) — callers
 * should surface the resulting error as "only the creator can edit teams".
 */
export async function saveTeamsConfig(
  driveFileId: string | null,
  data: unknown,
  creatorEmail: string
): Promise<string> {
  if (driveFileId) {
    await updateFileContent(driveFileId, data);
    return driveFileId;
  }
  return createTeamsConfigFile(data, creatorEmail);
}

export interface MediaConfigResult<T> {
  data: T | null;
  driveFileId: string | null;
  ownerEmail: string | null;
}

/** Loads the single shared media-config file (the scouting media list), if one exists yet. */
export async function loadMediaConfig<T = any>(adminEmail: string): Promise<MediaConfigResult<T>> {
  const existing = await findMediaConfigFile(adminEmail);
  if (!existing) return { data: null, driveFileId: null, ownerEmail: null };
  const content = await readFileContent<T>(existing.id);
  writeMediaConfigCache(content);
  ensureDomainPermission(existing.id, 'writer').catch(() => {});
  return { data: content, driveFileId: existing.id, ownerEmail: existing.ownerEmail ?? null };
}

/**
 * Creates the shared media-config file the first time the app runs after this feature
 * shipped, or updates it if it already exists. Same drive.file-scope constraint as teams:
 * only the original creator's browser can successfully update it afterwards.
 */
export async function saveMediaConfig(
  driveFileId: string | null,
  data: unknown,
  creatorEmail: string
): Promise<string> {
  writeMediaConfigCache(data);
  if (driveFileId) {
    await updateFileContent(driveFileId, data);
    return driveFileId;
  }
  return createMediaConfigFile(data, creatorEmail);
}
