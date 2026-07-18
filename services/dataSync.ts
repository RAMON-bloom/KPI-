import { findOwnDataFile, readFileContent, createOwnDataFile, updateFileContent, listTeammateDataFiles, findTeamsConfigFile, createTeamsConfigFile, findMediaConfigFile, createMediaConfigFile, ensureDomainPermission } from './googleDrive';

const LOCAL_CACHE_PREFIX = 'kpiUserDataCache:';
const DRIVE_FILE_ID_CACHE_PREFIX = 'kpiDriveFileId:';
const LEGACY_APPDATA_KEY = 'kpiAppData';
const MEDIA_CONFIG_CACHE_KEY = 'kpiMediaConfigCache';
const SCHEMA_VERSION = 1;

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
  if (debounceTimer) clearTimeout(debounceTimer);
  debounceTimer = setTimeout(async () => {
    try {
      const payload = { ...(data as object), schemaVersion: SCHEMA_VERSION };
      if (driveFileId) {
        await updateFileContent(driveFileId, payload);
      } else {
        const newId = await createOwnDataFile(payload, email);
        setCachedDriveFileId(email, newId);
        onFileCreated(newId);
      }
    } catch (err) {
      console.error('Failed to sync data to Drive', err);
    }
  }, 2000);
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
