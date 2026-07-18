import { findOwnDataFile, readFileContent, createOwnDataFile, updateFileContent, listTeammateDataFiles, findTeamsConfigFile, createTeamsConfigFile, findMediaConfigFile, createMediaConfigFile, ensureDomainPermission } from './googleDrive';

const LOCAL_CACHE_PREFIX = 'kpiUserDataCache:';
const DRIVE_FILE_ID_CACHE_PREFIX = 'kpiDriveFileId:';
const PENDING_SYNC_PREFIX = 'kpiPendingSync:';
const LEGACY_APPDATA_KEY = 'kpiAppData';
const MEDIA_CONFIG_CACHE_KEY = 'kpiMediaConfigCache';
const SCHEMA_VERSION = 1;

function pendingSyncKey(email: string): string {
  return `${PENDING_SYNC_PREFIX}${email}`;
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
}

/** Fetches every teammate's domain-shared data file (used by the cross-user / team views). */
export async function loadAllTeammatesData<T = any>(): Promise<TeammateData<T>[]> {
  const files = await listTeammateDataFiles();
  const results = await Promise.all(
    files.map(async (file) => {
      try {
        const content = await readFileContent<T>(file.id);
        return { email: file.ownerEmail || file.name, data: content };
      } catch (err) {
        console.error(`Failed to read teammate file ${file.id}`, err);
        return null;
      }
    })
  );
  return results.filter((r) => r !== null) as TeammateData<T>[];
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
