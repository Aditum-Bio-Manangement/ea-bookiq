import { getGraphClient } from "./graphClient";
import { OFFICE_CONFIGS, type OfficeConfig } from "../config/offices";

interface GraphGroup {
  id: string;
  displayName: string;
  mail: string | null;
  securityEnabled: boolean;
  mailEnabled: boolean;
}

// Target groups for office identification - only looking for these specific groups
const EA_CAMBRIDGE_EMAIL = "ea-cambridge@aditumbio.com";
const EA_OAKLAND_EMAIL = "ea-oakland@aditumbio.com";

/**
 * Check if the user is a member of a specific group by email
 * Uses checkMemberGroups API for direct membership check
 */
async function checkGroupMembership(groupEmail: string): Promise<boolean> {
  const client = await getGraphClient();

  try {
    // First, find the group ID by email
    const groupResponse = await client
      .api("/groups")
      .filter(`mail eq '${groupEmail}'`)
      .select("id,displayName,mail")
      .get();

    if (!groupResponse.value || groupResponse.value.length === 0) {
      console.log(`[AB Book IQ] Group not found: ${groupEmail}`);
      return false;
    }

    const groupId = groupResponse.value[0].id;
    const groupName = groupResponse.value[0].displayName;
    console.log(`[AB Book IQ] Found group ${groupName} (${groupEmail}) with ID: ${groupId}`);

    // Check if the current user is a member of this group
    const memberCheckResponse = await client
      .api("/me/checkMemberGroups")
      .post({
        groupIds: [groupId]
      });

    const isMember = memberCheckResponse.value && memberCheckResponse.value.includes(groupId);
    console.log(`[AB Book IQ] User membership in ${groupName}: ${isMember}`);

    return isMember;
  } catch (err) {
    console.error(`[AB Book IQ] Error checking membership for ${groupEmail}:`, err);
    return false;
  }
}

/**
 * Determine which office(s) the user belongs to based on EA-Cambridge or EA-Oakland group membership
 * Directly checks membership in ea-cambridge@aditumbio.com and ea-oakland@aditumbio.com
 */
export async function resolveUserOffices(): Promise<OfficeConfig[]> {
  const matchedOffices: OfficeConfig[] = [];

  // Check Cambridge membership
  console.log("[AB Book IQ] Checking membership in EA-Cambridge...");
  const isCambridgeMember = await checkGroupMembership(EA_CAMBRIDGE_EMAIL);
  if (isCambridgeMember) {
    console.log("[AB Book IQ] User IS a member of EA-Cambridge");
    matchedOffices.push(OFFICE_CONFIGS.cambridge);
  }

  // Check Oakland membership
  console.log("[AB Book IQ] Checking membership in EA-Oakland...");
  const isOaklandMember = await checkGroupMembership(EA_OAKLAND_EMAIL);
  if (isOaklandMember) {
    console.log("[AB Book IQ] User IS a member of EA-Oakland");
    matchedOffices.push(OFFICE_CONFIGS.oakland);
  }

  console.log("[AB Book IQ] Resolved offices:", matchedOffices.map(o => o.name));
  return matchedOffices;
}

/**
 * Check if user belongs to a specific office group
 */
export async function isUserInOffice(officeId: string): Promise<boolean> {
  const offices = await resolveUserOffices();
  return offices.some((office) => office.id === officeId);
}
