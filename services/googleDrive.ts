import { getCurrentSession, refreshTokenSilently } from './googleAuth';

const DATA_FILE_NAME = 'kpi-manager-data.json';
const TEAMS_FILE_NAME = 'kpi-manager-teams.json';
const MEDIA_FILE_NAME = 'kpi-manager-media.json';
const APP_TAG = 'kpi-manager-v1';
const ALLOWED_DOMAIN = 'bloom-firm.com';

export interface DriveFileRef {
  id: string;
  name: string;
  modifiedTime: string;
  ownerEmail?: string;
}

async function authorizedFetch(url: string, init: RequestInit = {}, retried = false): Promise<Response> {
  const session = getCurrentSession();
  if (!session) throw new Error('ログインが必要です。');
  const res = await fetch(url, {
    ...init,
    headers: {
      ...(init.headers || {}),
      Authorization: `Bearer ${session.accessToken}`,
    },
  });
  if (res.status === 401 && !retried) {
    await refreshTokenSilently();
    return authorizedFetch(url, init, true);
  }
  return res;
}

async function grantDomainPermission(fileId: string, role: 'reader' | 'writer'): Promise<boolean> {
  try {
    const res = await authorizedFetch(`https://www.googleapis.com/drive/v3/files/${fileId}/permissions`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ role, type: 'domain', domain: ALLOWED_DOMAIN }),
    });
    if (!res.ok) {
      console.error(`Failed to grant domain permission on Drive file ${fileId}: ${res.status} ${await res.text()}`);
      return false;
    }
    return true;
  } catch (err) {
    // Non-fatal: the file still exists and is usable by its owner even if sharing fails.
    console.error('Failed to grant domain permission on Drive file', err);
    return false;
  }
}

/**
 * Checks whether a file already has a domain-wide permission, and grants one if it doesn't.
 * Self-heals files whose sharing failed silently at creation time (e.g. a transient API
 * error) — call this opportunistically whenever a file is loaded, not just when it's created.
 */
export async function ensureDomainPermission(fileId: string, role: 'reader' | 'writer'): Promise<void> {
  try {
    const res = await authorizedFetch(`https://www.googleapis.com/drive/v3/files/${fileId}/permissions?fields=permissions(type,role)`);
    if (res.ok) {
      const data = await res.json();
      const hasDomainPermission = (data.permissions || []).some((p: any) => p.type === 'domain');
      if (hasDomainPermission) return;
      console.warn(`Drive file ${fileId} was missing its domain-wide permission — re-granting now.`);
    }
  } catch (err) {
    console.error(`Failed to check permissions on Drive file ${fileId}`, err);
  }
  await grantDomainPermission(fileId, role);
}

/** Finds the signed-in user's own kpi-manager-data.json in their My Drive, if it exists. */
export async function findOwnDataFile(): Promise<DriveFileRef | null> {
  const q = `name='${DATA_FILE_NAME}' and appProperties has { key='app' and value='${APP_TAG}' } and 'me' in owners and trashed=false`;
  const res = await authorizedFetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&fields=files(id,name,modifiedTime)&spaces=drive`
  );
  if (!res.ok) throw new Error('Driveファイルの検索に失敗しました。');
  const data = await res.json();
  return data.files?.[0] ?? null;
}

/**
 * Finds every teammate's kpi-manager-data.json shared domain-wide (requires drive.readonly).
 * Deliberately matches on filename alone, NOT `appProperties` (see history below), and uses
 * `corpora=domain` — the default `files.list` corpus only covers files the caller owns or
 * that were shared with them individually; files shared via a `type: domain` permission (our
 * sharing model) only show up when `corpora=domain` is explicitly requested. This — not the
 * appProperties issue — was the actual root cause of the "全ユーザー" cross-user visibility bug.
 */
export async function listTeammateDataFiles(): Promise<DriveFileRef[]> {
  const q = `name='${DATA_FILE_NAME}' and trashed=false`;
  const url = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&fields=files(id,name,modifiedTime,properties,owners(emailAddress))&spaces=drive&corpora=domain`;
  const res = await authorizedFetch(url);
  console.log('[listTeammateDataFiles] query:', q, 'status:', res.status);
  if (!res.ok) throw new Error('チームメンバーのデータ検索に失敗しました。');
  const data = await res.json();
  const files = (data.files ?? []).map((f: any) => ({
    id: f.id,
    name: f.name,
    modifiedTime: f.modifiedTime,
    ownerEmail: f.properties?.ownerEmail || f.owners?.[0]?.emailAddress,
  }));
  console.log('[listTeammateDataFiles] found', files.length, 'file(s):', files.map((f: DriveFileRef) => f.ownerEmail || f.id));
  return files;
}

export async function readFileContent<T = any>(fileId: string): Promise<T> {
  const res = await authorizedFetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`);
  if (!res.ok) throw new Error('Driveファイルの読み込みに失敗しました。');
  return res.json();
}

async function createJsonFile(name: string, appProperties: Record<string, string>, properties: Record<string, string>, content: unknown): Promise<string> {
  const metadata = { name, mimeType: 'application/json', appProperties, properties };
  const boundary = `kpi_boundary_${Math.random().toString(16).slice(2)}`;
  const body =
    `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n` +
    `--${boundary}\r\nContent-Type: application/json\r\n\r\n${JSON.stringify(content)}\r\n` +
    `--${boundary}--`;
  const res = await authorizedFetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', {
    method: 'POST',
    headers: { 'Content-Type': `multipart/related; boundary=${boundary}` },
    body,
  });
  if (!res.ok) throw new Error('Driveファイルの作成に失敗しました。');
  const data = await res.json();
  return data.id;
}

/** Creates the signed-in user's own data file and immediately shares it domain-wide (read-only). */
export async function createOwnDataFile(content: unknown, ownerEmail: string): Promise<string> {
  const fileId = await createJsonFile(DATA_FILE_NAME, { app: APP_TAG }, { ownerEmail }, content);
  await grantDomainPermission(fileId, 'reader');
  return fileId;
}

export async function updateFileContent(fileId: string, content: unknown): Promise<void> {
  const res = await authorizedFetch(`https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(content),
  });
  if (!res.ok) throw new Error('Driveファイルの保存に失敗しました。');
}

/**
 * Finds the single shared teams-config file (created by whoever first set up teams), if it exists.
 * Matches on filename alone (not `appProperties`) and uses `corpora=domain` — see the comment
 * on listTeammateDataFiles for why both of those matter for a domain-shared file.
 */
export async function findTeamsConfigFile(): Promise<DriveFileRef | null> {
  const q = `name='${TEAMS_FILE_NAME}' and trashed=false`;
  const res = await authorizedFetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&fields=files(id,name,modifiedTime,owners(emailAddress))&spaces=drive&corpora=domain`
  );
  if (!res.ok) throw new Error('チーム設定ファイルの検索に失敗しました。');
  const data = await res.json();
  const file = data.files?.[0];
  if (!file) return null;
  return { id: file.id, name: file.name, modifiedTime: file.modifiedTime, ownerEmail: file.owners?.[0]?.emailAddress };
}

/** Creates the shared teams-config file; the creator becomes the only one who can edit it (drive.file scope). */
export async function createTeamsConfigFile(content: unknown, creatorEmail: string): Promise<string> {
  const fileId = await createJsonFile(TEAMS_FILE_NAME, { app: APP_TAG, kind: 'teams-config' }, { ownerEmail: creatorEmail }, content);
  await grantDomainPermission(fileId, 'writer');
  return fileId;
}

/**
 * Finds the single shared media-config file (the list of scouting media sources), if it exists.
 * Matches on filename alone (not `appProperties`) and uses `corpora=domain` — see the comment
 * on listTeammateDataFiles for why both of those matter for a domain-shared file.
 */
export async function findMediaConfigFile(): Promise<DriveFileRef | null> {
  const q = `name='${MEDIA_FILE_NAME}' and trashed=false`;
  const res = await authorizedFetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&fields=files(id,name,modifiedTime,owners(emailAddress))&spaces=drive&corpora=domain`
  );
  if (!res.ok) throw new Error('媒体設定ファイルの検索に失敗しました。');
  const data = await res.json();
  const file = data.files?.[0];
  if (!file) return null;
  return { id: file.id, name: file.name, modifiedTime: file.modifiedTime, ownerEmail: file.owners?.[0]?.emailAddress };
}

/** Creates the shared media-config file; the creator becomes the only one who can edit it (drive.file scope). */
export async function createMediaConfigFile(content: unknown, creatorEmail: string): Promise<string> {
  const fileId = await createJsonFile(MEDIA_FILE_NAME, { app: APP_TAG, kind: 'media-config' }, { ownerEmail: creatorEmail }, content);
  await grantDomainPermission(fileId, 'writer');
  return fileId;
}

function escapeDriveQueryValue(value: string): string {
  return value.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

export interface InterviewLogFile {
  id: string;
  name: string;
  modifiedTime: string;
  webViewLink?: string;
}

/**
 * Searches the signed-in user's own Drive (drive.readonly — this app was never granted
 * access to files it didn't create, so this only ever searches the current user's Drive,
 * not other users') for Google Docs whose name or content mentions the candidate. Google
 * Meet saves auto-generated transcripts/notes as Google Docs, so this is restricted to that
 * mimeType.
 */
export async function searchInterviewLogsByName(candidateName: string): Promise<InterviewLogFile[]> {
  const escaped = escapeDriveQueryValue(candidateName);
  const q = `mimeType='application/vnd.google-apps.document' and trashed=false and (name contains '${escaped}' or fullText contains '${escaped}')`;
  const res = await authorizedFetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&fields=files(id,name,modifiedTime,webViewLink)&orderBy=modifiedTime desc&pageSize=5&spaces=drive`
  );
  if (!res.ok) throw new Error('面談ログの検索に失敗しました。');
  const data = await res.json();
  return data.files ?? [];
}

/** Exports a native Google Doc's content as plain text (Meet transcripts/notes are Google Docs). */
export async function exportGoogleDocAsText(fileId: string): Promise<string> {
  const res = await authorizedFetch(`https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=text/plain`);
  if (!res.ok) throw new Error('面談ログの読み込みに失敗しました。');
  return res.text();
}
