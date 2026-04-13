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

// Get redirect URI at runtime (must be called client-side)
function getRedirectUri(): string {
  if (process.env.NEXT_PUBLIC_REDIRECT_URI) {
    return process.env.NEXT_PUBLIC_REDIRECT_URI;
  }
  if (typeof window !== "undefined") {
    return `${window.location.origin}/taskpane`;
  }
  return "http://localhost:3000/taskpane";
}

// Build MSAL config at runtime to ensure redirectUri is correct
function getMsalConfig(): Configuration {
  return {
    auth: {
      clientId: MSAL_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${MSAL_CONFIG.tenantId}`,
      redirectUri: getRedirectUri(),
      // Disable navigation client to prevent history.replaceState errors in iframes
      navigateToLoginRequestUrl: false,
    },
    cache: {
      // Use localStorage for persistent sessions across browser sessions
      cacheLocation: "localStorage",
      storeAuthStateInCookie: true,
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
    console.warn("[EA BookIQ] MSAL initialize warning (may be in iframe):", initError);
  }
  
  console.log("[EA BookIQ] Standard MSAL initialized with popup auth");
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
        redirectUri: getRedirectUri(),
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
      redirectUri: getRedirectUri(),
    });
    return response.account;
  } catch (error) {
    console.error("[EA BookIQ] Sign in failed:", error);
    throw error;
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
