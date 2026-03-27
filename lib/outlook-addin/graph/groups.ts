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
 * Get the user's group memberships (both direct and transitive)
 * Tries multiple approaches since mail-enabled security groups may not appear in all queries
 */
export async function getUserGroupMemberships(): Promise<GraphGroup[]> {
  const client = await getGraphClient();
  const allGroups: Map<string, GraphGroup> = new Map();

  // Approach 1: Try memberOf (direct memberships)
  try {
    const directResponse: GroupMembershipResponse = await client
      .api("/me/memberOf/microsoft.graph.group")
      .select("id,displayName,mail,securityEnabled,mailEnabled")
      .get();

    console.log("[v0] Direct memberOf groups:", directResponse.value.length);
    directResponse.value.forEach(g => allGroups.set(g.id, g));
  } catch (err) {
    console.log("[v0] memberOf query failed:", err);
  }

  // Approach 2: Try transitiveMemberOf with ConsistencyLevel header
  try {
    const transitiveResponse: GroupMembershipResponse = await client
      .api("/me/transitiveMemberOf/microsoft.graph.group")
      .select("id,displayName,mail,securityEnabled,mailEnabled")
      .header("ConsistencyLevel", "eventual")
      .count(true)
      .get();

    console.log("[v0] Transitive memberOf groups:", transitiveResponse.value.length);
    transitiveResponse.value.forEach(g => allGroups.set(g.id, g));
  } catch (err) {
    console.log("[v0] transitiveMemberOf query failed:", err);
  }

  const groups = Array.from(allGroups.values());

  // Debug: Log all groups returned
  console.log("[v0] All group memberships:", groups.map(g => ({
    displayName: g.displayName,
    mail: g.mail,
    id: g.id,
    mailEnabled: g.mailEnabled,
    securityEnabled: g.securityEnabled
  })));

  return groups;
}

/**
 * Determine which office(s) the user belongs to based on group membership
 * Matches by email address OR display name
 */
export async function resolveUserOffices(): Promise<OfficeConfig[]> {
  const groups = await getUserGroupMemberships();
  const matchedOffices: OfficeConfig[] = [];

  for (const config of Object.values(OFFICE_CONFIGS)) {
    // Extract the group name part (e.g., "ea-cambridge" from "ea-cambridge@aditumbio.com")
    const groupNamePattern = config.securityGroupEmail.split("@")[0].toLowerCase();

    const isMatch = groups.some((group) => {
      // Match by mail address (exact match)
      const mailMatch = group.mail?.toLowerCase() === config.securityGroupEmail.toLowerCase();

      // Match by displayName (e.g., "EA-Cambridge", "EA Cambridge", "EA_Cambridge")
      const displayNameLower = group.displayName?.toLowerCase() || "";
      const nameMatch =
        displayNameLower === groupNamePattern ||
        displayNameLower === groupNamePattern.replace("-", "") ||
        displayNameLower === groupNamePattern.replace("-", " ") ||
        displayNameLower === groupNamePattern.replace("-", "_") ||
        displayNameLower.includes(groupNamePattern);

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
