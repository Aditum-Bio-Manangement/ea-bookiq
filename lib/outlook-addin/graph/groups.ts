import { getGraphClient } from "./graphClient";
import { OFFICE_CONFIGS, type OfficeConfig } from "../config/offices";

interface GraphGroup {
  id: string;
  displayName: string;
  mail: string | null;
  securityEnabled: boolean;
  mailEnabled: boolean;
}

interface GroupMembershipResponse {
  value: GraphGroup[];
}

/**
 * Get the user's transitive group memberships
 */
export async function getTransitiveGroupMemberships(): Promise<GraphGroup[]> {
  const client = await getGraphClient();

  const response: GroupMembershipResponse = await client
    .api("/me/transitiveMemberOf/microsoft.graph.group")
    .select("id,displayName,mail,securityEnabled,mailEnabled")
    .header("ConsistencyLevel", "eventual")
    .get();

  // Debug: Log all groups returned
  console.log("[v0] All group memberships:", response.value.map(g => ({
    displayName: g.displayName,
    mail: g.mail,
    id: g.id
  })));

  return response.value;
}

/**
 * Determine which office(s) the user belongs to based on group membership
 * Matches by email address OR display name containing the office identifier
 */
export async function resolveUserOffices(): Promise<OfficeConfig[]> {
  const groups = await getTransitiveGroupMemberships();
  const matchedOffices: OfficeConfig[] = [];

  for (const config of Object.values(OFFICE_CONFIGS)) {
    // Extract the group name part (e.g., "ea-cambridge" from "ea-cambridge@aditumbio.com")
    const groupNamePattern = config.securityGroupEmail.split("@")[0].toLowerCase();

    const isMatch = groups.some((group) => {
      // Match by mail address
      const mailMatch = group.mail?.toLowerCase() === config.securityGroupEmail.toLowerCase();
      // Match by displayName containing the pattern (e.g., "EA-Cambridge" or "EA Cambridge")
      const nameMatch = group.displayName?.toLowerCase().includes(groupNamePattern) ||
        group.displayName?.toLowerCase().includes(groupNamePattern.replace("-", " "));

      if (mailMatch || nameMatch) {
        console.log("[v0] Matched group for", config.name, ":", {
          displayName: group.displayName,
          mail: group.mail,
          matchedBy: mailMatch ? "mail" : "displayName"
        });
      }

      return mailMatch || nameMatch;
    });

    if (isMatch) {
      matchedOffices.push(config);
    }
  }

  console.log("[v0] Resolved offices:", matchedOffices.map(o => o.name));
  return matchedOffices;
}

/**
 * Check if user belongs to a specific office group
 */
export async function isUserInOffice(officeId: string): Promise<boolean> {
  const offices = await resolveUserOffices();
  return offices.some((office) => office.id === officeId);
}
