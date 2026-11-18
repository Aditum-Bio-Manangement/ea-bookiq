import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { DEFAULT_GROUP_IDS, type OfficeConfig } from "../config/offices";
import {
  CHECK_MEMBER_GROUPS_BATCH_SIZE,
  evaluateDirectoryAccess,
  fetchMatchingGroupIds,
  type CheckMemberGroupsResponse,
  type GraphMemberGroupsClient,
} from "./groups";

const CAMBRIDGE: OfficeConfig = {
  id: "cambridge",
  name: "Cambridge",
  displayName: "Cambridge Office",
  securityGroupId: DEFAULT_GROUP_IDS.cambridge,
  securityGroupEmail: "ea-cambridge@aditumbio.com",
  building: "Cambridge",
};

const OAKLAND: OfficeConfig = {
  id: "oakland",
  name: "Oakland",
  displayName: "Oakland Office",
  securityGroupId: DEFAULT_GROUP_IDS.oakland,
  securityGroupEmail: "ea-oakland@aditumbio.com",
  building: "Oakland",
};

const OFFICES = [CAMBRIDGE, OAKLAND];
const AUTHORIZED = [DEFAULT_GROUP_IDS.fte];

function evaluate(memberGroupIds: string[]) {
  return evaluateDirectoryAccess(memberGroupIds, OFFICES, AUTHORIZED);
}

describe("evaluateDirectoryAccess", () => {
  it("authorizes a direct EA-Cambridge member and selects Cambridge only", () => {
    const result = evaluate([DEFAULT_GROUP_IDS.cambridge]);
    assert.equal(result.authorized, true);
    assert.deepEqual(
      result.offices.map((office) => office.id),
      ["cambridge"]
    );
  });

  it("authorizes a direct EA-Oakland member and selects Oakland only", () => {
    const result = evaluate([DEFAULT_GROUP_IDS.oakland]);
    assert.equal(result.authorized, true);
    assert.deepEqual(
      result.offices.map((office) => office.id),
      ["oakland"]
    );
  });

  it("lets a member of both EA groups pick either office", () => {
    const result = evaluate([DEFAULT_GROUP_IDS.cambridge, DEFAULT_GROUP_IDS.oakland]);
    assert.equal(result.authorized, true);
    assert.deepEqual(
      result.offices.map((office) => office.id),
      ["cambridge", "oakland"]
    );
  });

  it("authorizes an FTE-only member and offers both offices", () => {
    const result = evaluate([DEFAULT_GROUP_IDS.fte]);
    assert.equal(result.authorized, true);
    assert.deepEqual(
      result.offices.map((office) => office.id),
      ["cambridge", "oakland"]
    );
  });

  it("treats nested FTE membership the same as direct FTE (checkMemberGroups is transitive)", () => {
    // Graph POST /me/checkMemberGroups returns ABM-SG-FTE's object id for both
    // direct members and users nested under a child group. The app only sees the id.
    const result = evaluate([DEFAULT_GROUP_IDS.fte.toUpperCase()]);
    assert.equal(result.authorized, true);
    assert.equal(result.offices.length, 2);
  });

  it("keeps the EA office when the user is also in FTE", () => {
    const result = evaluate([DEFAULT_GROUP_IDS.fte, DEFAULT_GROUP_IDS.cambridge]);
    assert.equal(result.authorized, true);
    assert.deepEqual(
      result.offices.map((office) => office.id),
      ["cambridge"]
    );
  });

  it("denies a user in neither FTE nor an EA office group", () => {
    const result = evaluate(["11111111-1111-1111-1111-111111111111"]);
    assert.equal(result.authorized, false);
    assert.deepEqual(result.offices, []);
  });

  it("denies a user with no group memberships", () => {
    const result = evaluate([]);
    assert.equal(result.authorized, false);
    assert.deepEqual(result.offices, []);
  });
});

function createMemberGroupsClient(options: {
  onPost?: (path: string, body: { groupIds: string[] }) => CheckMemberGroupsResponse;
  onGet?: (path: string) => CheckMemberGroupsResponse;
}): GraphMemberGroupsClient {
  return {
    api: (path: string) => ({
      post: async (body: { groupIds: string[] }) => {
        if (!options.onPost) throw new Error(`unexpected POST ${path}`);
        return options.onPost(path, body);
      },
      get: async () => {
        if (!options.onGet) throw new Error(`unexpected GET ${path}`);
        return options.onGet(path);
      },
    }),
  };
}

describe("fetchMatchingGroupIds", () => {
  it("posts group ids to /me/checkMemberGroups and returns matches", async () => {
    const posted: string[][] = [];
    const client = createMemberGroupsClient({
      onPost: (path, body) => {
        assert.equal(path, "/me/checkMemberGroups");
        posted.push(body.groupIds);
        return { value: [DEFAULT_GROUP_IDS.fte, DEFAULT_GROUP_IDS.cambridge] };
      },
    });

    const matched = await fetchMatchingGroupIds(client, [
      DEFAULT_GROUP_IDS.cambridge,
      DEFAULT_GROUP_IDS.oakland,
      DEFAULT_GROUP_IDS.fte,
    ]);

    assert.equal(posted.length, 1);
    assert.equal(matched.has(DEFAULT_GROUP_IDS.fte), true);
    assert.equal(matched.has(DEFAULT_GROUP_IDS.cambridge), true);
    assert.equal(matched.has(DEFAULT_GROUP_IDS.oakland), false);
  });

  it("follows @odata.nextLink when Graph pages the checkMemberGroups result", async () => {
    const client = createMemberGroupsClient({
      onPost: () => ({
        value: [DEFAULT_GROUP_IDS.cambridge],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/checkMemberGroups?$skiptoken=page2",
      }),
      onGet: (path) => {
        assert.match(path, /skiptoken=page2/);
        return { value: [DEFAULT_GROUP_IDS.fte] };
      },
    });

    const matched = await fetchMatchingGroupIds(client, [
      DEFAULT_GROUP_IDS.cambridge,
      DEFAULT_GROUP_IDS.fte,
    ]);

    assert.equal(matched.has(DEFAULT_GROUP_IDS.cambridge), true);
    assert.equal(matched.has(DEFAULT_GROUP_IDS.fte), true);
  });

  it("batches requests when more than 20 group ids are checked", async () => {
    const batches: number[] = [];
    const ids = Array.from(
      { length: CHECK_MEMBER_GROUPS_BATCH_SIZE + 3 },
      (_, index) => `00000000-0000-0000-0000-${String(index).padStart(12, "0")}`
    );

    const client = createMemberGroupsClient({
      onPost: (_path, body) => {
        batches.push(body.groupIds.length);
        return { value: body.groupIds.slice(0, 1) };
      },
    });

    const matched = await fetchMatchingGroupIds(client, ids);
    assert.deepEqual(batches, [CHECK_MEMBER_GROUPS_BATCH_SIZE, 3]);
    assert.equal(matched.size, 2);
  });
});
