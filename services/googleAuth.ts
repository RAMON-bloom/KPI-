export interface GoogleIdentity {
  email: string;
  name: string;
  picture?: string;
}

interface StoredSession {
  accessToken: string;
  expiresAt: number; // epoch ms
  identity: GoogleIdentity;
}

const SESSION_KEY = 'kpiGoogleSession';
const LAST_EMAIL_KEY = 'kpiLastSignedInEmail';
const ALLOWED_DOMAIN = 'bloom-firm.com';
// Full `drive` (not just drive.file/drive.readonly) is required so that a designated teams
// editor — someone other than a shared file's original creator — can actually write to it:
// drive.file only grants per-file write access to files this specific app instance created,
// regardless of any Drive-level "writer" sharing permission on the file. Broadening to full
// Drive access removes that restriction (the existing domain "writer" permission on
// kpi-manager-teams.json etc. becomes actionable for anyone holding it), at the cost of
// granting this app read/write access to each signed-in user's entire Drive, not just files
// it created — a real scope increase, accepted deliberately for the multi-editor teams feature.
// tasks: lets a user who opts in push their own パイプラインカレンダー entries to their own
// Google Tasks list (see services/googleTasks.ts) — another real scope increase, same
// re-consent story as the Gmail addition above (existing sessions need reauthorizeWithConsent).
const SCOPES = 'openid email profile https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/gmail.readonly https://www.googleapis.com/auth/tasks';

declare global {
  interface Window {
    google?: any;
  }
}

let gisReadyPromise: Promise<void> | null = null;

function waitForGis(): Promise<void> {
  if (gisReadyPromise) return gisReadyPromise;
  gisReadyPromise = new Promise((resolve, reject) => {
    const start = Date.now();
    const check = () => {
      if (window.google?.accounts?.oauth2) {
        resolve();
      } else if (Date.now() - start > 10000) {
        reject(new Error('Google Identity Servicesの読み込みに失敗しました。ネットワーク接続を確認してください。'));
      } else {
        setTimeout(check, 100);
      }
    };
    check();
  });
  return gisReadyPromise;
}

function getStoredSession(): StoredSession | null {
  try {
    const raw = sessionStorage.getItem(SESSION_KEY);
    if (!raw) return null;
    const session: StoredSession = JSON.parse(raw);
    if (!session.expiresAt || session.expiresAt < Date.now()) return null;
    return session;
  } catch {
    return null;
  }
}

function storeSession(session: StoredSession) {
  sessionStorage.setItem(SESSION_KEY, JSON.stringify(session));
}

function clearSession() {
  sessionStorage.removeItem(SESSION_KEY);
}

export function getLastKnownEmail(): string | null {
  try {
    return localStorage.getItem(LAST_EMAIL_KEY);
  } catch {
    return null;
  }
}

function setLastKnownEmail(email: string) {
  try {
    localStorage.setItem(LAST_EMAIL_KEY, email);
  } catch {
    // ignore storage failures (e.g. private browsing) — just means silent restore won't work later
  }
}

function clearLastKnownEmail() {
  try {
    localStorage.removeItem(LAST_EMAIL_KEY);
  } catch {
    // ignore
  }
}

async function fetchIdentity(accessToken: string): Promise<GoogleIdentity> {
  const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error('ユーザー情報の取得に失敗しました。');
  const data = await res.json();
  return { email: data.email, name: data.name, picture: data.picture };
}

function assertAllowedDomain(identity: GoogleIdentity) {
  if (!identity.email.toLowerCase().endsWith(`@${ALLOWED_DOMAIN}`)) {
    throw new Error(`${ALLOWED_DOMAIN} のGoogleアカウントでログインしてください。`);
  }
}

function requestToken(prompt: '' | 'consent', hint?: string): Promise<GoogleIdentity> {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  if (!clientId) {
    return Promise.reject(new Error('GOOGLE_CLIENT_IDが設定されていません（.env.local を確認してください）。'));
  }
  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (fn: () => void) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeoutId);
      fn();
    };
    // initTokenClient's callback is never invoked if the popup fails to open (e.g. blocked
    // by the browser), so without this the UI would be stuck showing "signing in" forever.
    // A silent (prompt: '') restore attempt should fail fast rather than wait the full 30s.
    const timeoutMs = prompt === '' ? 8000 : 30000;
    const timeoutId = setTimeout(() => {
      finish(() => reject(new Error('ログイン用のポップアップを開けませんでした。ポップアップブロッカーを解除してもう一度お試しください。')));
    }, timeoutMs);

    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: SCOPES,
      hd: ALLOWED_DOMAIN,
      hint,
      callback: async (response: any) => {
        if (response.error) {
          finish(() => reject(new Error('Googleログインが完了しませんでした。もう一度お試しください。')));
          return;
        }
        try {
          const identity = await fetchIdentity(response.access_token);
          assertAllowedDomain(identity);
          const expiresAt = Date.now() + ((response.expires_in ?? 3600) * 1000);
          storeSession({ accessToken: response.access_token, expiresAt, identity });
          setLastKnownEmail(identity.email);
          finish(() => resolve(identity));
        } catch (err) {
          clearSession();
          finish(() => reject(err instanceof Error ? err : new Error(String(err))));
        }
      },
      error_callback: (error: any) => {
        const message = error?.type === 'popup_failed_to_open'
          ? 'ログイン用のポップアップを開けませんでした。ポップアップブロッカーを解除してもう一度お試しください。'
          : error?.type === 'popup_closed'
          ? 'ログインがキャンセルされました。'
          : 'Googleログインでエラーが発生しました。';
        finish(() => reject(new Error(message)));
      },
    } as any);
    client.requestAccessToken({ prompt, hint } as any);
  });
}

/** Returns the currently signed-in identity/token from this browser session, if any and not expired. */
export function getCurrentSession(): { accessToken: string; identity: GoogleIdentity } | null {
  const session = getStoredSession();
  if (!session) return null;
  return { accessToken: session.accessToken, identity: session.identity };
}

/**
 * Signs the user in. Called from a click handler, so — unlike an automatic page-load
 * attempt — the browser allows the popup this opens. Tries silently first (skips the
 * account-chooser/consent screen if this browser already has an active Google session and
 * previously granted consent, using the last-known email as a hint), and only falls back to
 * the full interactive consent prompt if that doesn't work (e.g. first-ever sign-in, or the
 * browser session/consent has since expired).
 */
export async function signIn(): Promise<GoogleIdentity> {
  await waitForGis();
  const lastEmail = getLastKnownEmail() ?? undefined;
  if (lastEmail) {
    try {
      return await requestToken('', lastEmail);
    } catch {
      // Fall through to the interactive flow — e.g. the browser session/consent expired.
    }
  }
  return requestToken('consent', lastEmail);
}

/** Silently re-requests an access token using the existing browser session (no prompt shown). */
export async function refreshTokenSilently(): Promise<GoogleIdentity> {
  await waitForGis();
  const lastEmail = getLastKnownEmail() ?? undefined;
  return requestToken('', lastEmail);
}

/**
 * Forces the interactive consent screen even if a session already exists — used when a
 * newly-added scope (e.g. Gmail) needs to be granted by users who signed in before it existed;
 * a silent request would keep re-using the older, narrower grant instead of prompting for it.
 */
export async function reauthorizeWithConsent(): Promise<GoogleIdentity> {
  await waitForGis();
  const lastEmail = getLastKnownEmail() ?? undefined;
  return requestToken('consent', lastEmail);
}

export function signOut() {
  const session = getStoredSession();
  clearSession();
  clearLastKnownEmail();
  if (session?.accessToken && window.google?.accounts?.oauth2?.revoke) {
    window.google.accounts.oauth2.revoke(session.accessToken, () => {});
  }
}
