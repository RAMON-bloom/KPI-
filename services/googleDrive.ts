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

async function grantDomainPermission(fileId: string, role: 'reader' | 'writer'): Promise<void> {
  try {
    await authorizedFetch(`https://www.googleapis.com/drive/v3/files/${fileId}/permissions`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ role, type: 'domain', domain: ALLOWED_DOMAIN }),
    });
  } catch (err) {
    // Non-fatal: the file still exists and is usable by its owner even if sharing fails.
    console.error('Failed to grant domain permission on Drive file', err);
  }
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

/** Finds every teammate's kpi-manager-data.json shared domain-wide (requires drive.readonly). */
export async function listTeammateDataFiles(): Promise<DriveFileRef[]> {
  const q = `name='${DATA_FILE_NAME}' and appProperties has { key='app' and value='${APP_TAG}' } and trashed=false`;
  const res = await authorizedFetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&fields=files(id,name,modifiedTime,properties)&spaces=drive&corpora=allDrives&includeItemsFromAllDrives=true&supportsAllDrives=true`
  );
  if (!res.ok) throw new Error('チームメンバーのデータ検索に失敗しました。');
  const data = await res.json();
  return (data.files ?? []).map((f: any) => ({
    id: f.id,
    name: f.name,
    modifiedTime: f.modifiedTime,
    ownerEmail: f.properties?.ownerEmail,
  }));
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

/** Finds the single shared teams-config file (created by whoever first set up teams), if it exists. */
export async function findTeamsConfigFile(): Promise<DriveFileRef | null> {
  const q = `name='${TEAMS_FILE_NAME}' and appProperties has { key='kind' and value='teams-config' } and trashed=false`;
  const res = await authorizedFetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&fields=files(id,name,modifiedTime,owners(emailAddress))&spaces=drive`
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

/** Finds the single shared media-config file (the list of scouting media sources), if it exists. */
export async function findMediaConfigFile(): Promise<DriveFileRef | null> {
  const q = `name='${MEDIA_FILE_NAME}' and appProperties has { key='kind' and value='media-config' } and trashed=false`;
  const res = await authorizedFetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&fields=files(id,name,modifiedTime,owners(emailAddress))&spaces=drive`
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
