// Pushes パイプラインカレンダー entries (an application's scheduledDate) to the signed-in
// user's own Google Tasks default list ("My Tasks"), so they show up as a lightweight,
// due-date-only reminder in Google Calendar's task lane. Only the signed-in account's own
// Tasks can ever be written here — there is no way to write into a teammate's Tasks with this
// user's token — which is why the sync in index.tsx only ever runs over the current user's own
// candidates, never teammates' aggregated data.

export class GoogleTasksPermissionError extends Error {}

const TASKS_BASE = 'https://tasks.googleapis.com/tasks/v1/lists/@default/tasks';

async function tasksFetch(accessToken: string, path: string, init?: RequestInit): Promise<any> {
  const res = await fetch(`${TASKS_BASE}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(init?.body ? { 'Content-Type': 'application/json' } : {}),
    },
  });
  if (res.status === 401 || res.status === 403) {
    throw new GoogleTasksPermissionError('Googleタスクの利用権限が許可されていません。');
  }
  // A 404 on update/delete means the task was already removed on the Google side (e.g. the
  // user deleted it themselves) — treat that the same as a successful delete/no-op rather
  // than surfacing an error for something that's already in the desired end state.
  if (res.status === 404) return null;
  if (!res.ok) {
    throw new Error(`Google Tasks APIエラー（${res.status}）が発生しました。`);
  }
  if (res.status === 204) return null;
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

export interface PipelineTaskContent {
  title: string;
  notes?: string;
  dueDateISO: string; // yyyy-mm-dd — Google Tasks ignores the time-of-day/timezone portion
}

function toDueTimestamp(dueDateISO: string): string {
  return `${dueDateISO}T00:00:00.000Z`;
}

/** Creates a new task in the user's default list and returns its Google-assigned task ID. */
export async function createPipelineTask(accessToken: string, content: PipelineTaskContent): Promise<string> {
  const created = await tasksFetch(accessToken, '', {
    method: 'POST',
    body: JSON.stringify({ title: content.title, notes: content.notes, due: toDueTimestamp(content.dueDateISO) }),
  });
  return created.id;
}

/**
 * Updates an existing task in place. If it's since been deleted on the Google side, creates a
 * fresh one instead of silently dropping the link — the returned ID may therefore differ from
 * `taskId`; callers must re-store whatever ID comes back.
 */
export async function updatePipelineTask(accessToken: string, taskId: string, content: PipelineTaskContent): Promise<string> {
  const updated = await tasksFetch(accessToken, `/${encodeURIComponent(taskId)}`, {
    method: 'PATCH',
    body: JSON.stringify({ title: content.title, notes: content.notes, due: toDueTimestamp(content.dueDateISO) }),
  });
  return updated ? updated.id : createPipelineTask(accessToken, content);
}

export async function deletePipelineTask(accessToken: string, taskId: string): Promise<void> {
  await tasksFetch(accessToken, `/${encodeURIComponent(taskId)}`, { method: 'DELETE' });
}

export interface ExistingPipelineTask {
  id: string;
  title: string;
  dueDateISO?: string; // yyyy-mm-dd, derived from the task's `due` field
}

/**
 * Lists tasks currently in the user's default list — used only by the manual bulk resync
 * ("今すぐ同期") to recognize a task that already exists on Google's side even when the app's own
 * locally-cached id→task mapping has gone stale (e.g. a debounced Drive write from another
 * tab/device hadn't landed yet before this session's copy was loaded). Without this
 * cross-check, a stale mapping looks identical to "never created," and a resync would create a
 * genuine duplicate on top of the still-existing task. Paginates up to a generous cap; a
 * recruiter's own task list realistically never approaches it.
 */
export async function listPipelineTasks(accessToken: string): Promise<ExistingPipelineTask[]> {
  const results: ExistingPipelineTask[] = [];
  let pageToken: string | undefined;
  let pagesFetched = 0;
  do {
    const params = new URLSearchParams({ showCompleted: 'false', showHidden: 'true', maxResults: '100' });
    if (pageToken) params.set('pageToken', pageToken);
    const data = await tasksFetch(accessToken, `?${params.toString()}`);
    (data?.items || []).forEach((item: any) => {
      results.push({ id: item.id, title: item.title || '', dueDateISO: item.due ? item.due.slice(0, 10) : undefined });
    });
    pageToken = data?.nextPageToken;
    pagesFetched++;
  } while (pageToken && pagesFetched < 10);
  return results;
}
