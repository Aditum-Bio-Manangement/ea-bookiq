import {
  PublicClientApplication,
  createNestablePublicClientApplication,
  type IPublicClientApplication,
  type Configuration,
  type AccountInfo,
  type AuthenticationResult,
  type SilentRequest,
  type PopupRequest,
} from "@azure/msal-browser";
import { MSAL_CONFIG, GRAPH_SCOPES } from "../config/offices";
import {
  formatSignInError,
  getBrokerRedirectUri,
  getMailboxLoginHint,
  isInteractionRequiredError,
  isNestedAppAuthSupported,
  isUnrecoverableInvalidGrant,
} from "./msalHelpers";

export {
  formatSignInError,
  getBrokerRedirectUri,
  getMailboxLoginHint,
  isInteractionRequiredError,
  isNestedAppAuthSupported,
  isUnrecoverableInvalidGrant,
  routeAuthFailure,
} from "./msalHelpers";

let msalInstance: IPublicClientApplication | null = null;
let isNaaMode = false;

export type MsalRuntime = {
  createPublicClientApplication: (config: Configuration) => IPublicClientApplication;
  createNestablePublicClientApplication: (
    config: Configuration
  ) => Promise<IPublicClientApplication>;
};

const defaultRuntime: MsalRuntime = {
  createPublicClientApplication: (config) => new PublicClientApplication(config),
  createNestablePublicClientApplication,
};

let runtime: MsalRuntime = defaultRuntime;

function getOfficeGlobal(): typeof Office | undefined {
  return typeof Office !== "undefined" ? Office : undefined;
}

function currentLoginHint(): string | undefined {
  return getMailboxLoginHint(getOfficeGlobal(), getAccount()?.username);
}

// Polyfill history.replaceState if not available (e.g., in iframes)
function ensureHistoryApi(): void {
  if (typeof window !== "undefined" && typeof window.history !== "undefined") {
    if (typeof window.history.replaceState !== "function") {
      window.history.replaceState = function () {
        // No-op polyfill for restricted contexts
      };
    }
    if (typeof window.history.pushState !== "function") {
      window.history.pushState = function () {
        // No-op polyfill for restricted contexts
      };
    }
  }
}

function requireRedirectUri(): string {
  const redirectUri = process.env.NEXT_PUBLIC_REDIRECT_URI;
  if (!redirectUri) {
    throw new Error(
      "Missing NEXT_PUBLIC_REDIRECT_URI. Set it to the hosted task pane URL before build or start."
    );
  }
  return redirectUri;
}

function currentHost(): string {
  if (typeof window !== "undefined" && window.location?.host) {
    return window.location.host;
  }
  return "";
}

function setActiveAccountSafe(account: AccountInfo | null | undefined): void {
  if (!msalInstance || !account) return;
  try {
    msalInstance.setActiveAccount(account);
  } catch (err) {
    // NAA does not support setActiveAccount.
    console.warn("[AB Book IQ] setActiveAccount skipped:", err);
  }
}

function tokenRequest(account?: AccountInfo | null): SilentRequest {
  const loginHint = currentLoginHint();
  return {
    scopes: GRAPH_SCOPES,
    ...(account ? { account } : {}),
    ...(loginHint ? { loginHint } : {}),
  };
}

function popupRequest(account?: AccountInfo | null): PopupRequest {
  return {
    ...tokenRequest(account),
    redirectUri: requireRedirectUri(),
  };
}

/**
 * Clear MSAL tokens and auth state from memory and browser storage.
 * Only used on explicit sign-out or an unrecoverable invalid_grant.
 */
export async function clearMsalCache(): Promise<void> {
  if (msalInstance) {
    try {
      await msalInstance.clearCache();
    } catch (err) {
      console.warn("[AB Book IQ] MSAL clearCache warning:", err);
    }
  }

  if (typeof window !== "undefined") {
    const storages = [window.localStorage, window.sessionStorage];
    for (const storage of storages) {
      if (!storage) continue;
      const keysToRemove: string[] = [];
      for (let i = 0; i < storage.length; i++) {
        const key = storage.key(i);
        if (!key) continue;
        const normalized = key.toLowerCase();
        if (
          normalized.includes("msal") ||
          normalized.startsWith("server-telemetry") ||
          normalized.includes("login.windows") ||
          normalized.includes("login.microsoftonline")
        ) {
          keysToRemove.push(key);
        }
      }
      for (const key of keysToRemove) {
        storage.removeItem(key);
      }
    }
  }

  msalInstance = null;
  isNaaMode = false;
}

async function clearCacheIfUnrecoverableGrant(error: unknown): Promise<void> {
  if (!isUnrecoverableInvalidGrant(error)) return;
  console.warn("[AB Book IQ] Clearing MSAL cache after unrecoverable invalid_grant");
  await clearMsalCache();
}

function getMsalConfig(): Configuration {
  return {
    auth: {
      clientId: MSAL_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${MSAL_CONFIG.tenantId}`,
      redirectUri: requireRedirectUri(),
      navigateToLoginRequestUrl: false,
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false,
    },
    system: {
      allowRedirectInIframe: true,
    },
  };
}

function getNestedAppAuthConfig(): Configuration {
  return {
    auth: {
      clientId: MSAL_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${MSAL_CONFIG.tenantId}`,
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false,
    },
  };
}

async function initializeStandardMsal(): Promise<IPublicClientApplication> {
  ensureHistoryApi();
  const instance = runtime.createPublicClientApplication(getMsalConfig());
  try {
    await instance.initialize();
  } catch (initError) {
    console.warn("[AB Book IQ] MSAL initialize warning (may be in iframe):", initError);
  }
  isNaaMode = false;
  console.log("[AB Book IQ] Standard MSAL initialized with popup fallback");
  return instance;
}

/**
 * Initialize MSAL. Uses Nested App Auth when the host supports NestedAppAuth 1.1,
 * otherwise the existing popup PublicClientApplication.
 */
export async function initializeMsal(): Promise<IPublicClientApplication> {
  if (msalInstance) {
    return msalInstance;
  }

  ensureHistoryApi();

  if (isNestedAppAuthSupported(getOfficeGlobal())) {
    try {
      const host = currentHost();
      if (host) {
        console.log("[AB Book IQ] NAA broker redirect URI:", getBrokerRedirectUri(host));
      }
      msalInstance = await runtime.createNestablePublicClientApplication(getNestedAppAuthConfig());
      try {
        await msalInstance.initialize();
      } catch (initError) {
        console.warn("[AB Book IQ] NAA initialize warning:", initError);
      }
      isNaaMode = true;
      console.log("[AB Book IQ] Nested App Auth MSAL initialized");
      return msalInstance;
    } catch (naaError) {
      console.warn("[AB Book IQ] NAA init failed, using popup fallback:", naaError);
    }
  }

  msalInstance = await initializeStandardMsal();
  return msalInstance;
}

export function isUsingNestedAppAuth(): boolean {
  return isNaaMode;
}

/**
 * Get the current account
 */
export function getAccount(): AccountInfo | null {
  if (!msalInstance) return null;
  try {
    const active = msalInstance.getActiveAccount();
    if (active) return active;
  } catch {
    // getActiveAccount is supported by NAA, but be defensive.
  }
  const accounts = msalInstance.getAllAccounts();
  return accounts.length > 0 ? accounts[0] : null;
}

async function acquireSilentResult(): Promise<AuthenticationResult> {
  if (!msalInstance) {
    await initializeMsal();
  }

  const cached = getAccount();
  const request = tokenRequest(cached);
  try {
    const response = await msalInstance!.acquireTokenSilent(request);
    setActiveAccountSafe(response.account);
    return response;
  } catch (silentError) {
    await clearCacheIfUnrecoverableGrant(silentError);
    if (msalInstance && isInteractionRequiredError(silentError)) {
      try {
        const ssoResponse = await msalInstance.ssoSilent(tokenRequest(getAccount()));
        setActiveAccountSafe(ssoResponse.account);
        return ssoResponse;
      } catch (ssoError) {
        await clearCacheIfUnrecoverableGrant(ssoError);
        throw ssoError;
      }
    }
    throw silentError;
  }
}

/**
 * Try to establish a session without opening a popup.
 * Returns the account on success, or null when interaction is required.
 */
export async function trySilentSignIn(): Promise<AccountInfo | null> {
  if (!msalInstance) {
    await initializeMsal();
  }

  try {
    const response = await acquireSilentResult();
    return response.account ?? getAccount();
  } catch (error) {
    if (isInteractionRequiredError(error) || isUnrecoverableInvalidGrant(error)) {
      return null;
    }
    throw error;
  }
}

/**
 * Acquire a token for Microsoft Graph. Silent only.
 * Callers must route InteractionRequiredAuthError to the Sign In screen
 * instead of opening a popup outside a user click.
 */
export async function acquireGraphToken(): Promise<string> {
  if (!msalInstance) {
    await initializeMsal();
  }

  const response = await acquireSilentResult();
  return response.accessToken;
}

/**
 * Interactive sign-in. Must be called from a user click so the popup is not blocked.
 */
export async function signIn(): Promise<AccountInfo | null> {
  if (!msalInstance) {
    await initializeMsal();
  }

  try {
    const response = await msalInstance!.loginPopup(popupRequest(getAccount()));
    setActiveAccountSafe(response.account);
    return response.account;
  } catch (error) {
    console.error("[AB Book IQ] Sign in failed:", error);
    // Closing, blocking, or failing the popup is not a reason to wipe cached tokens.
    throw formatSignInError(error);
  }
}

/**
 * Sign out the user and clear cached tokens.
 */
export async function signOut(): Promise<void> {
  try {
    if (msalInstance) {
      const account = getAccount();
      if (account && !isNaaMode) {
        await msalInstance.logoutPopup({ account });
      }
    }
  } catch (err) {
    console.warn("[AB Book IQ] Sign out warning:", err);
  } finally {
    await clearMsalCache();
  }
}

/**
 * Check if user is signed in
 */
export function isSignedIn(): boolean {
  return getAccount() !== null;
}

/** Test-only: replace MSAL constructors and drop in-memory state. */
export function __resetMsalForTests(overrides?: Partial<MsalRuntime>): void {
  msalInstance = null;
  isNaaMode = false;
  runtime = {
    ...defaultRuntime,
    ...overrides,
  };
}
