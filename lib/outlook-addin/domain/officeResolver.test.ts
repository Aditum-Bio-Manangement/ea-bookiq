import assert from "node:assert/strict";
import { afterEach, describe, it } from "node:test";
import { OFFICE_CONFIGS, STORAGE_KEYS, type OfficeConfig } from "../config/offices";
import {
  clearCachedOfficePreference,
  resolveOfficeFromList,
  setCachedOfficePreference,
} from "./officeResolver";

const CAMBRIDGE = OFFICE_CONFIGS.cambridge;
const OAKLAND = OFFICE_CONFIGS.oakland;
const BOTH: OfficeConfig[] = [CAMBRIDGE, OAKLAND];

function installLocalStorage() {
  const map = new Map<string, string>();
  const localStorage = {
    getItem: (key: string) => (map.has(key) ? map.get(key)! : null),
    setItem: (key: string, value: string) => {
      map.set(key, value);
    },
    removeItem: (key: string) => {
      map.delete(key);
    },
    clear: () => map.clear(),
  };
  (globalThis as { localStorage?: typeof localStorage }).localStorage = localStorage;
  return localStorage;
}

describe("resolveOfficeFromList", () => {
  afterEach(() => {
    delete (globalThis as { localStorage?: unknown }).localStorage;
  });

  it("returns none when the user is not authorized (empty office list)", () => {
    assert.deepEqual(resolveOfficeFromList([]), { type: "none" });
  });

  it("returns a single office for a direct EA member of one site", () => {
    assert.deepEqual(resolveOfficeFromList([CAMBRIDGE]), {
      type: "single",
      office: CAMBRIDGE,
    });
  });

  it("offers both sites to FTE-only users so they can choose Cambridge or Oakland", () => {
    installLocalStorage();
    assert.deepEqual(resolveOfficeFromList(BOTH), {
      type: "multiple",
      offices: BOTH,
      cached: undefined,
    });
  });

  it("reuses a cached site preference for FTE-only / multi-office users", () => {
    installLocalStorage();
    setCachedOfficePreference("oakland");
    assert.deepEqual(resolveOfficeFromList(BOTH), {
      type: "multiple",
      offices: BOTH,
      cached: OAKLAND,
    });
    assert.equal(
      globalThis.localStorage.getItem(STORAGE_KEYS.OFFICE_PREFERENCE),
      "oakland"
    );
    clearCachedOfficePreference();
    assert.equal(globalThis.localStorage.getItem(STORAGE_KEYS.OFFICE_PREFERENCE), null);
  });
});
