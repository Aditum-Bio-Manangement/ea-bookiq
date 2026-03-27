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
let isNaaMode = false;

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
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
    },
  };
}

/**
 * Initialize MSAL - attempts NAA first, falls back to standard MSAL
 */
export async function initializeMsal(): Promise<IPublicClientApplication> {
  if (msalInstance) {
    return msalInstance;
  }

  // Try Nested App Authentication (NAA) first for Outlook add-in context
  if (typeof Office !== "undefined" && Office.context) {
    try {
      // Check if NAA is available (Office.js 1.3+)
      const { createNestablePublicClientApplication } = await import(
        "@azure/msal-browser"
      );
      if (createNestablePublicClientApplication) {
        msalInstance = await createNestablePublicClientApplication({
          auth: {
            clientId: MSAL_CONFIG.clientId,
            authority: `https://login.microsoftonline.com/${MSAL_CONFIG.tenantId}`,
          },
        });
        isNaaMode = true;
        console.log("[EA BookIQ] NAA mode initialized");
        return msalInstance;
      }
    } catch {
      console.log("[EA BookIQ] NAA not available, using standard MSAL");
    }
  }

  // Fall back to standard MSAL
  msalInstance = new PublicClientApplication(getMsalConfig());
  await msalInstance.initialize();
  console.log("[EA BookIQ] Standard MSAL initialized");
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
