import assert from "node:assert/strict";
import { describe, it } from "node:test";
import {
  formatSignInError,
  getBrokerRedirectUri,
  getMailboxLoginHint,
  isInteractionRequiredError,
  isNestedAppAuthSupported,
  isUnrecoverableInvalidGrant,
  routeAuthFailure,
  type OfficeRequirementsHost,
} from "./msalHelpers";

function officeHost(options: {
  supported?: boolean;
  email?: string;
  throwOnCheck?: boolean;
}): OfficeRequirementsHost {
  return {
    context: {
      requirements: {
        isSetSupported: (name: string, version: string) => {
          if (options.throwOnCheck) throw new Error("requirements unavailable");
          return name === "NestedAppAuth" && version === "1.1" && options.supported === true;
        },
      },
      mailbox: {
        userProfile: {
          emailAddress: options.email,
        },
      },
    },
  };
}

describe("isNestedAppAuthSupported", () => {
  it("returns true only for NestedAppAuth 1.1", () => {
    assert.equal(isNestedAppAuthSupported(officeHost({ supported: true })), true);
    assert.equal(isNestedAppAuthSupported(officeHost({ supported: false })), false);
    assert.equal(isNestedAppAuthSupported(undefined), false);
    assert.equal(isNestedAppAuthSupported({}), false);
    assert.equal(isNestedAppAuthSupported(officeHost({ throwOnCheck: true })), false);
  });
});

describe("getBrokerRedirectUri", () => {
  it("derives brk-multihub from the page host and never hardcodes the domain", () => {
    assert.equal(getBrokerRedirectUri("roomiq.aditumbio.com"), "brk-multihub://roomiq.aditumbio.com");
    assert.equal(getBrokerRedirectUri("localhost:3000"), "brk-multihub://localhost:3000");
    assert.throws(() => getBrokerRedirectUri("  "), /window.location.host/);
  });
});

describe("getMailboxLoginHint", () => {
  it("prefers the Outlook mailbox profile and falls back to a username", () => {
    assert.equal(
      getMailboxLoginHint(officeHost({ email: "user@aditumbio.com" })),
      "user@aditumbio.com"
    );
    assert.equal(getMailboxLoginHint(undefined, "cached@aditumbio.com"), "cached@aditumbio.com");
    assert.equal(getMailboxLoginHint(officeHost({}), "not-an-email"), undefined);
  });
});

describe("auth error classification", () => {
  it("treats InteractionRequiredAuthError and related codes as interaction required", () => {
    assert.equal(
      isInteractionRequiredError({ name: "InteractionRequiredAuthError", errorCode: "interaction_required" }),
      true
    );
    assert.equal(isInteractionRequiredError({ errorCode: "login_required" }), true);
    assert.equal(isInteractionRequiredError({ errorCode: "consent_required" }), true);
    assert.equal(isInteractionRequiredError({ errorCode: "no_account_in_silent_request" }), true);
    assert.equal(isInteractionRequiredError({ errorCode: "popup_window_error" }), false);
    assert.equal(isInteractionRequiredError({ errorCode: "user_cancelled" }), false);
  });

  it("detects unrecoverable invalid_grant and AADSTS refresh failures", () => {
    assert.equal(isUnrecoverableInvalidGrant({ errorCode: "invalid_grant" }), true);
    assert.equal(isUnrecoverableInvalidGrant({ message: "AADSTS700082: The refresh token has expired" }), true);
    assert.equal(isUnrecoverableInvalidGrant({ errorCode: "user_cancelled" }), false);
  });

  it("routes interaction-required and invalid_grant to the Sign In screen", () => {
    assert.equal(routeAuthFailure({ name: "InteractionRequiredAuthError" }), "sign-in");
    assert.equal(routeAuthFailure({ errorCode: "invalid_grant" }), "sign-in");
    assert.equal(routeAuthFailure({ errorCode: "user_cancelled" }), "error");
    assert.equal(routeAuthFailure(new Error("Graph 500")), "error");
  });

  it("formats sign-in errors without claiming the cache was cleared", () => {
    const formatted = formatSignInError({
      errorCode: "user_cancelled",
      errorMessage: "User cancelled the flow",
    });
    assert.match(formatted.message, /User cancelled the flow/);
    assert.doesNotMatch(formatted.message, /cache was cleared/i);
  });
});
