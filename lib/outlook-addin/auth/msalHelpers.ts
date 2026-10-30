/**
 * Pure auth helpers. Kept free of MSAL instance state so they can be unit tested
 * without spinning up a browser PublicClientApplication.
 */

const INTERACTION_REQUIRED_CODES = new Set([
  "interaction_required",
  "login_required",
  "consent_required",
  "interaction_required_auth",
  "no_tokens_found",
  "no_account_error",
  "no_account_in_silent_request",
  "monitor_window_timeout",
  "token_renewal_error",
]);

const UNRECOVERABLE_GRANT_CODES = new Set([
  "invalid_grant",
]);

const UNRECOVERABLE_GRANT_MARKERS = [
  "invalid_grant",
  "aadsts70000",
  "aadsts700082",
  "aadsts700084",
  "aadsts50173",
];

export type OfficeRequirementsHost = {
  context?: {
    requirements?: {
      isSetSupported?: (name: string, version: string) => boolean;
    };
    mailbox?: {
      userProfile?: {
        emailAddress?: string;
      };
    };
  };
};

export type AuthFailureRoute = "sign-in" | "error";

function readErrorCode(error: unknown): string | undefined {
  if (error && typeof error === "object" && "errorCode" in error) {
    const code = (error as { errorCode?: unknown }).errorCode;
    if (typeof code === "string" && code.length > 0) {
      return code;
    }
  }
  return undefined;
}

function readErrorText(error: unknown): string {
  if (!error) return "";
  if (typeof error === "string") return error;
  if (error && typeof error === "object") {
    const authError = error as {
      name?: string;
      errorCode?: string;
      errorMessage?: string;
      message?: string;
    };
    return [authError.name, authError.errorCode, authError.errorMessage, authError.message]
      .filter(Boolean)
      .join(" ");
  }
  if (error instanceof Error) {
    return error.message;
  }
  return String(error);
}

export function isNestedAppAuthSupported(office: OfficeRequirementsHost | undefined): boolean {
  try {
    return office?.context?.requirements?.isSetSupported?.("NestedAppAuth", "1.1") === true;
  } catch {
    return false;
  }
}

/**
 * NAA broker redirect URI. Host only, no path.
 * Example: roomiq.aditumbio.com -> brk-multihub://roomiq.aditumbio.com
 */
export function getBrokerRedirectUri(host: string): string {
  const trimmed = host.trim();
  if (!trimmed) {
    throw new Error("Cannot derive NAA broker redirect URI without window.location.host");
  }
  return `brk-multihub://${trimmed}`;
}

export function getMailboxLoginHint(
  office: OfficeRequirementsHost | undefined,
  fallback?: string | null
): string | undefined {
  try {
    const email = office?.context?.mailbox?.userProfile?.emailAddress;
    if (typeof email === "string" && email.includes("@")) {
      return email;
    }
  } catch {
    // Office profile can be missing outside Outlook or in preview mode.
  }
  if (typeof fallback === "string" && fallback.includes("@")) {
    return fallback;
  }
  return undefined;
}

export function isInteractionRequiredError(error: unknown): boolean {
  if (!error) return false;
  if (error && typeof error === "object") {
    const name = (error as { name?: string }).name;
    if (name === "InteractionRequiredAuthError") {
      return true;
    }
  }
  const code = readErrorCode(error);
  if (code && INTERACTION_REQUIRED_CODES.has(code)) {
    return true;
  }
  const text = readErrorText(error).toLowerCase();
  for (const marker of INTERACTION_REQUIRED_CODES) {
    if (text.includes(marker)) {
      return true;
    }
  }
  return false;
}

export function isUnrecoverableInvalidGrant(error: unknown): boolean {
  const code = readErrorCode(error);
  if (code && UNRECOVERABLE_GRANT_CODES.has(code)) {
    return true;
  }
  const text = readErrorText(error).toLowerCase();
  return UNRECOVERABLE_GRANT_MARKERS.some((marker) => text.includes(marker));
}

/**
 * Room loads and meeting changes should never auto-open a popup.
 * Interaction-required and unrecoverable grants go back to Sign In Required.
 */
export function routeAuthFailure(error: unknown): AuthFailureRoute {
  if (isInteractionRequiredError(error) || isUnrecoverableInvalidGrant(error)) {
    return "sign-in";
  }
  return "error";
}

export function getAuthErrorCode(error: unknown): string | undefined {
  return readErrorCode(error);
}

export function getAuthErrorMessage(error: unknown): string {
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

export function formatSignInError(error: unknown): Error {
  const errorCode = getAuthErrorCode(error);
  const message = getAuthErrorMessage(error);
  const details = errorCode ? `${message} (error code: ${errorCode})` : message;
  const formatted = new Error(details);
  if (errorCode) {
    (formatted as Error & { errorCode?: string }).errorCode = errorCode;
  }
  return formatted;
}
