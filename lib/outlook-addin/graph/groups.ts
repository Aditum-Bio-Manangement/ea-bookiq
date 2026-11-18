import { getGraphClient } from "./graphClient";
import {
  getConfiguredGroupIds,
  getConfiguredOffices,
  type OfficeConfig,
} from "../config/offices";

/** Graph accepts at most 20 ids per /me/checkMemberGroups request. */
export const CHECK_MEMBER_GROUPS_BATCH_SIZE = 20;

export interface DirectoryAccessResult {
  authorized: boolean;
  offices: OfficeConfig[];
}

export interface CheckMemberGroupsResponse {
  value?: string[];
  "@odata.nextLink"?: string;
}

/**
 * Minimal Graph surface used for membership checks so unit tests can supply a fake client.
 */
export interface GraphMemberGroupsClient {
  api: (path: string) => {
    post: (body: { groupIds: string[] }) => Promise<CheckMemberGroupsResponse>;
    get: () => Promise<CheckMemberGroupsResponse>;
  };
}

/**
 * Decide authorization and office list from already-resolved group object ids.
 *
 * Office (Cambridge / Oakland) comes only from EA group membership. Membership in an
 * access group such as ABM-SG-FTE authorizes the user; if they have no EA group they
 * receive every office so the task pane can offer a site picker.
 */
export function evaluateDirectoryAccess(
  memberGroupIds: Iterable<string>,
  offices: OfficeConfig[] = getConfiguredOffices(),
  authorizedGroupIds: string[] = getConfiguredGroupIds().authorized
): DirectoryAccessResult {
  const membership = new Set(
    Array.from(memberGroupIds, (id) => id.trim().toLowerCase()).filter(Boolean)
  );

  const matchedOffices = offices.filter((office) =>
    membership.has(office.securityGroupId.toLowerCase())
  );

  const inAuthorizedGroup = authorizedGroupIds.some((id) =>
    membership.has(id.toLowerCase())
  );

  const authorized = inAuthorizedGroup || matchedOffices.length > 0;

  return {
    authorized,
    offices: authorized && matchedOffices.length === 0 ? offices : matchedOffices,
  };
}

/**
 * Transitive membership check for a known set of group object ids.
 *
 * Uses POST /me/checkMemberGroups (GroupMember.Read.All). Unlike GET /me/memberOf,
 * this is transitive (nested groups) and matches by object id, so groups with no
 * mail (ABM-SG-FTE) work. IDs are sent in batches of 20; @odata.nextLink is followed.
 */
export async function fetchMatchingGroupIds(
  client: GraphMemberGroupsClient,
  groupIds: string[]
): Promise<Set<string>> {
  const uniqueIds = [...new Set(groupIds.map((id) => id.trim()).filter(Boolean))];
  const matched = new Set<string>();

  for (let i = 0; i < uniqueIds.length; i += CHECK_MEMBER_GROUPS_BATCH_SIZE) {
    const batch = uniqueIds.slice(i, i + CHECK_MEMBER_GROUPS_BATCH_SIZE);
    let response = await client.api("/me/checkMemberGroups").post({
      groupIds: batch,
    });

    while (response) {
      for (const id of response.value ?? []) {
        if (id) matched.add(id.toLowerCase());
      }

      const nextLink = response["@odata.nextLink"];
      if (!nextLink) break;
      response = await client.api(nextLink).get();
    }
  }

  return matched;
}

export async function resolveUserAccess(): Promise<DirectoryAccessResult> {
  const offices = getConfiguredOffices();
  const { authorized: authorizedGroupIds } = getConfiguredGroupIds();
  const groupIdsToCheck = [
    ...offices.map((office) => office.securityGroupId),
    ...authorizedGroupIds,
  ];

  const client = await getGraphClient();
  const memberGroupIds = await fetchMatchingGroupIds(client, groupIdsToCheck);
  const access = evaluateDirectoryAccess(memberGroupIds, offices, authorizedGroupIds);
  console.log(
    "[AB Book IQ] checkMemberGroups matched",
    [...memberGroupIds],
    "authorized:",
    access.authorized,
    "offices:",
    access.offices.map((office) => office.name)
  );
  return access;
}

/**
 * Offices the signed-in user may book. Empty means not authorized.
 * FTE-only users receive both Cambridge and Oakland.
 */
export async function resolveUserOffices(): Promise<OfficeConfig[]> {
  const access = await resolveUserAccess();
  if (!access.authorized) return [];
  return access.offices;
}

/**
 * Check if user belongs to a specific office group
 */
export async function isUserInOffice(officeId: string): Promise<boolean> {
  const offices = await resolveUserOffices();
  return offices.some((office) => office.id === officeId);
}
