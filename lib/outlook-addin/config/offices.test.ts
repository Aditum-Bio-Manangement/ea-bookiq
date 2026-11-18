import assert from "node:assert/strict";
import { afterEach, describe, it } from "node:test";
import { DEFAULT_GROUP_IDS, getConfiguredGroupIds, getConfiguredOffices } from "./offices";

const ENV_KEYS = [
  "NEXT_PUBLIC_EA_CAMBRIDGE_GROUP_ID",
  "NEXT_PUBLIC_EA_OAKLAND_GROUP_ID",
  "NEXT_PUBLIC_AUTHORIZED_GROUP_IDS",
] as const;

const originalEnv = Object.fromEntries(
  ENV_KEYS.map((key) => [key, process.env[key]])
);

function restoreEnv() {
  for (const key of ENV_KEYS) {
    const value = originalEnv[key];
    if (value === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = value;
    }
  }
}

describe("getConfiguredGroupIds", () => {
  afterEach(restoreEnv);

  it("defaults to EA-Cambridge, EA-Oakland, and ABM-SG-FTE object ids", () => {
    restoreEnv();
    const ids = getConfiguredGroupIds();
    assert.equal(ids.cambridge, DEFAULT_GROUP_IDS.cambridge);
    assert.equal(ids.oakland, DEFAULT_GROUP_IDS.oakland);
    assert.deepEqual(ids.authorized, [DEFAULT_GROUP_IDS.fte]);
  });

  it("overrides group ids from NEXT_PUBLIC env vars", () => {
    process.env.NEXT_PUBLIC_EA_CAMBRIDGE_GROUP_ID = "AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA";
    process.env.NEXT_PUBLIC_EA_OAKLAND_GROUP_ID = "BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB";
    process.env.NEXT_PUBLIC_AUTHORIZED_GROUP_IDS =
      "CCCCCCCC-CCCC-CCCC-CCCC-CCCCCCCCCCCC, DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD";

    const ids = getConfiguredGroupIds();
    assert.equal(ids.cambridge, "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa");
    assert.equal(ids.oakland, "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb");
    assert.deepEqual(ids.authorized, [
      "cccccccc-cccc-cccc-cccc-cccccccccccc",
      "dddddddd-dddd-dddd-dddd-dddddddddddd",
    ]);

    const offices = getConfiguredOffices();
    assert.equal(offices[0].securityGroupId, ids.cambridge);
    assert.equal(offices[1].securityGroupId, ids.oakland);
  });
});
