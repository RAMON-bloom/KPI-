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

function requestToken(prompt: '' | 'consent'): Promise<GoogleIdentity> {
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
    const timeoutId = setTimeout(() => {
      finish(() => reject(new Error('ログイン用のポップアップを開けませんでした。ポップアップブロッカーを解除してもう一度お試しください。')));
    }, 30000);

    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: SCOPES,
      hd: ALLOWED_DOMAIN,
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
    client.requestAccessToken({ prompt });
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
  return requestToken('');
}

export function signOut() {
  const session = getStoredSession();
  clearSession();
  if (session?.accessToken && window.google?.accounts?.oauth2?.revoke) {
    window.google.accounts.oauth2.revoke(session.accessToken, () => {});
  }
}
