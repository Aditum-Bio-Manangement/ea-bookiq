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
 * Uses memberOf API which works for all group types including distribution lists
 */
async function checkGroupMembership(groupEmail: string): Promise<boolean> {
  const client = await getGraphClient();

  try {
    // Get all groups the user is a member of (including transitive membership)
    // This works for security groups, Microsoft 365 groups, and distribution lists
    const memberOfResponse = await client
      .api("/me/memberOf")
      .select("id,displayName,mail,mailEnabled,securityEnabled")
      .top(999)
      .get();

    if (!memberOfResponse.value || memberOfResponse.value.length === 0) {
      console.log(`[AB Book IQ] User has no group memberships`);
      return false;
    }

    // Check if any of the user's groups match the target email
    const normalizedTargetEmail = groupEmail.toLowerCase();
    const matchingGroup = memberOfResponse.value.find((group: any) => {
      const groupMail = group.mail?.toLowerCase();
      return groupMail === normalizedTargetEmail;
    });

    if (matchingGroup) {
      console.log(`[AB Book IQ] User IS a member of ${matchingGroup.displayName} (${groupEmail})`);
      return true;
    }

    console.log(`[AB Book IQ] User is NOT a member of ${groupEmail}`);
    return false;
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
