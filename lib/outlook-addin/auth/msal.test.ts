import assert from "node:assert/strict";
import { afterEach, beforeEach, describe, it } from "node:test";
import type { AccountInfo, AuthenticationResult, IPublicClientApplication } from "@azure/msal-browser";
import * as msal from "./msal";

process.env.NEXT_PUBLIC_REDIRECT_URI ??= "https://roomiq.aditumbio.com/taskpane";
process.env.NEXT_PUBLIC_AZURE_CLIENT_ID ??= "test-client-id";
process.env.NEXT_PUBLIC_AZURE_TENANT_ID ??= "test-tenant-id";

type StorageMock = {
  length: number;
  getItem: (key: string) => string | null;
  setItem: (key: string, value: string) => void;
  removeItem: (key: string) => void;
  key: (index: number) => string | null;
  clear: () => void;
};

function createStorage(): StorageMock {
  const map = new Map<string, string>();
  return {
    get length() {
      return map.size;
    },
    getItem: (key) => (map.has(key) ? map.get(key)! : null),
    setItem: (key, value) => {
      map.set(key, value);
    },
    removeItem: (key) => {
      map.delete(key);
    },
    key: (index) => Array.from(map.keys())[index] ?? null,
    clear: () => {
      map.clear();
    },
  };
}

function installBrowserGlobals(host = "roomiq.aditumbio.com") {
  const localStorage = createStorage();
  const sessionStorage = createStorage();
  const history = {
    replaceState() {},
    pushState() {},
  };
  (globalThis as { window?: unknown }).window = {
    location: { host, origin: `https://${host}` },
    localStorage,
    sessionStorage,
    history,
  };
  (globalThis as { localStorage?: StorageMock }).localStorage = localStorage;
  (globalThis as { sessionStorage?: StorageMock }).sessionStorage = sessionStorage;
  return { localStorage, sessionStorage };
}

function setOffice(supported: boolean, email = "user@aditumbio.com") {
  (globalThis as { Office?: unknown }).Office = {
    context: {
      requirements: {
        isSetSupported: (name: string, version: string) =>
          name === "NestedAppAuth" && version === "1.1" && supported,
      },
      mailbox: {
        userProfile: { emailAddress: email },
      },
    },
  };
}

const mockAccount = {
  homeAccountId: "home",
  localAccountId: "local",
  environment: "login.windows.net",
  tenantId: "tenant",
  username: "user@aditumbio.com",
} as AccountInfo;

function authResult(token = "graph-token"): AuthenticationResult {
  return {
    authority: "https://login.microsoftonline.com/tenant",
    uniqueId: "uid",
    tenantId: "tenant",
    scopes: ["User.Read"],
    account: mockAccount,
    idToken: "id",
    idTokenClaims: {},
    accessToken: token,
    fromCache: true,
    expiresOn: new Date(Date.now() + 3600_000),
    tokenType: "Bearer",
    correlationId: "corr",
  } as AuthenticationResult;
}

function interactionError(code = "interaction_required") {
  const error = new Error("interaction required");
  (error as Error & { name: string; errorCode: string }).name = "InteractionRequiredAuthError";
  (error as Error & { errorCode: string }).errorCode = code;
  return error;
}

function grantError() {
  const error = new Error("AADSTS700082: The refresh token has expired due to inactivity");
  (error as Error & { errorCode: string }).errorCode = "invalid_grant";
  return error;
}

function createMockPca(overrides: Partial<IPublicClientApplication> = {}) {
  const calls = {
    initialize: 0,
    acquireTokenSilent: 0,
    ssoSilent: 0,
    loginPopup: 0,
    acquireTokenPopup: 0,
    logoutPopup: 0,
    clearCache: 0,
    setActiveAccount: 0,
    lastSilentRequest: undefined as unknown,
    lastSsoRequest: undefined as unknown,
    lastPopupRequest: undefined as unknown,
  };

  let accounts: AccountInfo[] = [];
  let active: AccountInfo | null = null;

  const pca = {
    initialize: async () => {
      calls.initialize += 1;
    },
    getAllAccounts: () => accounts,
    getActiveAccount: () => active,
    setActiveAccount: (account) => {
      calls.setActiveAccount += 1;
      active = account;
      if (account && !accounts.includes(account)) {
        accounts = [account];
      }
    },
    acquireTokenSilent: async (request) => {
      calls.acquireTokenSilent += 1;
      calls.lastSilentRequest = request;
      throw interactionError();
    },
    ssoSilent: async (request) => {
      calls.ssoSilent += 1;
      calls.lastSsoRequest = request;
      throw interactionError();
    },
    loginPopup: async (request) => {
      calls.loginPopup += 1;
      calls.lastPopupRequest = request;
      accounts = [mockAccount];
      return authResult();
    },
    acquireTokenPopup: async () => {
      calls.acquireTokenPopup += 1;
      throw new Error("acquireTokenPopup should not be used for Graph calls");
    },
    logoutPopup: async () => {
      calls.logoutPopup += 1;
    },
    clearCache: async () => {
      calls.clearCache += 1;
      accounts = [];
      active = null;
    },
    ...overrides,
  } as IPublicClientApplication;

  return { pca, calls, setAccounts: (next: AccountInfo[]) => { accounts = next; } };
}

describe("initializeMsal", () => {
  beforeEach(() => {
    installBrowserGlobals();
    setOffice(false);
    msal.__resetMsalForTests();
  });

  afterEach(() => {
    msal.__resetMsalForTests();
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("uses Nested App Auth when NestedAppAuth 1.1 is supported", async () => {
    setOffice(true);
    const nestable = createMockPca();
    const standard = createMockPca();
    msal.__resetMsalForTests({
      createNestablePublicClientApplication: async () => nestable.pca,
      createPublicClientApplication: () => standard.pca,
    });

    const instance = await msal.initializeMsal();
    assert.equal(instance, nestable.pca);
    assert.equal(msal.isUsingNestedAppAuth(), true);
    assert.equal(standard.calls.initialize, 0);
  });

  it("falls back to the popup PublicClientApplication on older Outlook hosts", async () => {
    setOffice(false);
    const nestable = createMockPca();
    const standard = createMockPca();
    msal.__resetMsalForTests({
      createNestablePublicClientApplication: async () => nestable.pca,
      createPublicClientApplication: () => standard.pca,
    });

    const instance = await msal.initializeMsal();
    assert.equal(instance, standard.pca);
    assert.equal(msal.isUsingNestedAppAuth(), false);
  });

  it("falls back when nestable construction fails", async () => {
    setOffice(true);
    const standard = createMockPca();
    msal.__resetMsalForTests({
      createNestablePublicClientApplication: async () => {
        throw new Error("broker redirect URI issue");
      },
      createPublicClientApplication: () => standard.pca,
    });

    const instance = await msal.initializeMsal();
    assert.equal(instance, standard.pca);
    assert.equal(msal.isUsingNestedAppAuth(), false);
  });
});

describe("silent and interactive auth", () => {
  beforeEach(() => {
    installBrowserGlobals();
    setOffice(false, "user@aditumbio.com");
  });

  afterEach(() => {
    msal.__resetMsalForTests();
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("trySilentSignIn uses acquireTokenSilent then ssoSilent with a mailbox login hint", async () => {
    const silentRequests: unknown[] = [];
    const ssoRequests: unknown[] = [];
    const tracked = createMockPca({
      acquireTokenSilent: async (request) => {
        silentRequests.push(request);
        throw interactionError();
      },
      ssoSilent: async (request) => {
        ssoRequests.push(request);
        return authResult();
      },
    });
    msal.__resetMsalForTests({
      createPublicClientApplication: () => tracked.pca,
      createNestablePublicClientApplication: async () => tracked.pca,
    });

    const account = await msal.trySilentSignIn();
    assert.equal(account?.username, "user@aditumbio.com");
    assert.equal(silentRequests.length, 1);
    assert.equal(ssoRequests.length, 1);
    assert.equal((silentRequests[0] as { loginHint?: string }).loginHint, "user@aditumbio.com");
    assert.equal((ssoRequests[0] as { loginHint?: string }).loginHint, "user@aditumbio.com");
    assert.equal(tracked.calls.acquireTokenPopup, 0);
  });

  it("trySilentSignIn returns null on interaction-required without clearing the cache", async () => {
    const tracked = createMockPca();
    msal.__resetMsalForTests({
      createPublicClientApplication: () => tracked.pca,
    });

    const account = await msal.trySilentSignIn();
    assert.equal(account, null);
    assert.equal(tracked.calls.clearCache, 0);
    assert.equal(tracked.calls.loginPopup, 0);
  });

  it("clears cache only for unrecoverable invalid_grant", async () => {
    const tracked = createMockPca({
      acquireTokenSilent: async () => {
        throw grantError();
      },
    });
    msal.__resetMsalForTests({
      createPublicClientApplication: () => tracked.pca,
    });

    const account = await msal.trySilentSignIn();
    assert.equal(account, null);
    assert.equal(tracked.calls.clearCache, 1);
  });

  it("acquireGraphToken stays silent and never opens a popup", async () => {
    const tracked = createMockPca({
      acquireTokenSilent: async () => {
        throw interactionError();
      },
    });
    msal.__resetMsalForTests({
      createPublicClientApplication: () => tracked.pca,
    });

    await assert.rejects(() => msal.acquireGraphToken(), /interaction required/);
    assert.equal(tracked.calls.acquireTokenPopup, 0);
    assert.equal(tracked.calls.loginPopup, 0);
  });

  it("signIn passes loginHint, sets the active account, and does not wipe cache on popup cancel", async () => {
    const popupRequests: unknown[] = [];
    const tracked = createMockPca({
      loginPopup: async (request) => {
        popupRequests.push(request);
        throw Object.assign(new Error("user cancelled the popup"), {
          errorCode: "user_cancelled",
        });
      },
    });
    msal.__resetMsalForTests({
      createPublicClientApplication: () => tracked.pca,
    });

    await assert.rejects(() => msal.signIn(), /user cancelled/i);
    assert.equal(tracked.calls.clearCache, 0);
    assert.equal((popupRequests[0] as { loginHint?: string }).loginHint, "user@aditumbio.com");

    const success = createMockPca();
    msal.__resetMsalForTests({
      createPublicClientApplication: () => success.pca,
    });
    const account = await msal.signIn();
    assert.equal(account?.username, "user@aditumbio.com");
    assert.equal((success.calls.lastPopupRequest as { loginHint?: string }).loginHint, "user@aditumbio.com");
    assert.equal(success.calls.setActiveAccount, 1);
    assert.equal(success.calls.clearCache, 0);
  });

  it("signOut clears the MSAL cache", async () => {
    const tracked = createMockPca();
    msal.__resetMsalForTests({
      createPublicClientApplication: () => tracked.pca,
    });
    await msal.initializeMsal();
    await msal.signIn();
    await msal.signOut();
    assert.equal(tracked.calls.clearCache, 1);
  });

  it("derives the broker URI from window.location.host during NAA init", async () => {
    installBrowserGlobals("roomiq.aditumbio.com");
    setOffice(true);
    const logs: string[] = [];
    const originalLog = console.log;
    console.log = (...args: unknown[]) => {
      logs.push(args.map(String).join(" "));
    };
    const tracked = createMockPca();
    msal.__resetMsalForTests({
      createNestablePublicClientApplication: async () => tracked.pca,
    });
    try {
      await msal.initializeMsal();
      assert.ok(logs.some((line) => line.includes("brk-multihub://roomiq.aditumbio.com")));
    } finally {
      console.log = originalLog;
    }
  });
});
