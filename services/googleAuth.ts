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
const SCOPES = 'openid email profile https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.readonly';

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

function getLastKnownEmail(): string | null {
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

/** Opens the Google account chooser / consent prompt. */
export async function signIn(): Promise<GoogleIdentity> {
  await waitForGis();
  return requestToken('consent');
}

/** Silently re-requests an access token using the existing browser session (no prompt shown). */
export async function refreshTokenSilently(): Promise<GoogleIdentity> {
  await waitForGis();
  const lastEmail = getLastKnownEmail() ?? undefined;
  return requestToken('', lastEmail);
}

/**
 * Called once on app load to keep users signed in across browser restarts. The access token
 * itself is short-lived and only lives in sessionStorage, but as long as the browser still has
 * an active Google session (and the user hasn't revoked access), a silent token request using
 * the last-known email as a hint re-authenticates without showing any prompt. Returns null if
 * there's nothing to restore or the silent attempt fails, in which case the caller should show
 * the normal "Googleでログイン" screen.
 */
export async function tryRestoreSession(): Promise<GoogleIdentity | null> {
  const existing = getCurrentSession();
  if (existing) return existing.identity;

  const lastEmail = getLastKnownEmail();
  if (!lastEmail) return null;

  try {
    await waitForGis();
    return await requestToken('', lastEmail);
  } catch {
    return null;
  }
}

export function signOut() {
  const session = getStoredSession();
  clearSession();
  clearLastKnownEmail();
  if (session?.accessToken && window.google?.accounts?.oauth2?.revoke) {
    window.google.accounts.oauth2.revoke(session.accessToken, () => {});
  }
}
