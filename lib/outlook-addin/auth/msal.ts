import {
  PublicClientApplication,
  type IPublicClientApplication,
  type Configuration,
  type AccountInfo,
  type AuthenticationResult,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";
import { MSAL_CONFIG, GRAPH_SCOPES } from "../config/offices";

let msalInstance: IPublicClientApplication | null = null;
let failedSignInCount = 0;

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

function getAuthErrorCode(error: unknown): string | undefined {
  if (error && typeof error === "object" && "errorCode" in error) {
    const code = (error as { errorCode?: unknown }).errorCode;
    if (typeof code === "string" && code.length > 0) {
      return code;
    }
  }
  return undefined;
}

function getAuthErrorMessage(error: unknown): string {
  if (error && typeof error === "object") {
    const authError = error as { errorMessage?: string; message?: string };
    if (authError.errorMessage) {
      return authError.errorMessage;
    }
    if (authError.message) {
      return authError.message;
    }
  }
  if (error instanceof Error) {
    return error.message;
  }
  return "Sign in failed";
}

function formatSignInError(error: unknown, cacheCleared: boolean): Error {
  const errorCode = getAuthErrorCode(error);
  const message = getAuthErrorMessage(error);
  const details = errorCode
    ? `${message} (error code: ${errorCode})`
    : message;
  const suffix = cacheCleared
    ? " Sign-in cache was cleared after 2 failed attempts."
    : "";
  const formatted = new Error(`${details}${suffix}`);
  if (errorCode) {
    (formatted as Error & { errorCode?: string }).errorCode = errorCode;
  }
  return formatted;
}

/**
 * Clear MSAL tokens and auth state from memory and browser storage.
 * Used after repeated sign-in failures so the next attempt starts clean.
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
}

// Build MSAL config at runtime to ensure redirectUri is correct
function getMsalConfig(): Configuration {
  return {
    auth: {
      clientId: MSAL_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${MSAL_CONFIG.tenantId}`,
      redirectUri: requireRedirectUri(),
      // Disable navigation client to prevent history.replaceState errors in iframes
      navigateToLoginRequestUrl: false,
    },
    cache: {
      // Use localStorage for persistent sessions across browser sessions
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false,
    },
    system: {
      // Prevent MSAL from using history API which fails in iframes
      allowRedirectInIframe: true,
    },
  };
}

/**
 * Initialize MSAL - uses standard popup authentication
 * NAA (Nested App Authentication) is disabled due to broker URI issues
 */
export async function initializeMsal(): Promise<IPublicClientApplication> {
  if (msalInstance) {
    return msalInstance;
  }

  // Ensure history API is available (polyfill for iframe contexts)
  ensureHistoryApi();

  // Use standard MSAL with popup authentication
  // NAA mode disabled - broker redirect URIs have compatibility issues
  msalInstance = new PublicClientApplication(getMsalConfig());

  // Initialize MSAL - wrap in try-catch to handle iframe/restricted contexts
  try {
    await msalInstance.initialize();
  } catch (initError) {
    // In some iframe contexts, history.replaceState may not be available
    // MSAL can still function for popup-based auth
    console.warn("[AB Book IQ] MSAL initialize warning (may be in iframe):", initError);
  }

  console.log("[AB Book IQ] Standard MSAL initialized with popup auth");
  return msalInstance;
}

/**
 * Get the current account
 */
export function getAccount(): AccountInfo | null {
  if (!msalInstance) return null;
  const accounts = msalInstance.getAllAccounts();
  return accounts.length > 0 ? accounts[0] : null;
}

/**
 * Acquire a token for Microsoft Graph
 */
export async function acquireGraphToken(): Promise<string> {
  if (!msalInstance) {
    await initializeMsal();
  }

  const account = getAccount();
  const request = {
    scopes: GRAPH_SCOPES,
    account: account || undefined,
  };

  try {
    // Try silent token acquisition first
    const response: AuthenticationResult =
      await msalInstance!.acquireTokenSilent(request);
    return response.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      // Need interactive login - use popup with redirectUri
      const popupRequest = {
        ...request,
        redirectUri: requireRedirectUri(),
      };
      const response = await msalInstance!.acquireTokenPopup(popupRequest);
      return response.accessToken;
    }
    throw error;
  }
}

/**
 * Sign in the user
 */
export async function signIn(): Promise<AccountInfo | null> {
  if (!msalInstance) {
    await initializeMsal();
  }

  try {
    const response = await msalInstance!.loginPopup({
      scopes: GRAPH_SCOPES,
      redirectUri: requireRedirectUri(),
    });
    failedSignInCount = 0;
    return response.account;
  } catch (error) {
    failedSignInCount += 1;
    console.error("[AB Book IQ] Sign in failed:", error);

    let cacheCleared = false;
    if (failedSignInCount >= 2) {
      await clearMsalCache();
      failedSignInCount = 0;
      cacheCleared = true;
    }

    throw formatSignInError(error, cacheCleared);
  }
}

/**
 * Sign out the user
 */
export async function signOut(): Promise<void> {
  if (!msalInstance) return;
  const account = getAccount();
  if (account) {
    await msalInstance.logoutPopup({ account });
  }
}

/**
 * Check if user is signed in
 */
export function isSignedIn(): boolean {
  return getAccount() !== null;
}
