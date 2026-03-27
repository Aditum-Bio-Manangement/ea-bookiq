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

  return response.value;
}

/**
 * Determine which office(s) the user belongs to based on group membership
 */
export async function resolveUserOffices(): Promise<OfficeConfig[]> {
  const groups = await getTransitiveGroupMemberships();
  const matchedOffices: OfficeConfig[] = [];

  for (const config of Object.values(OFFICE_CONFIGS)) {
    const isMatch = groups.some(
      (group) =>
        group.mail?.toLowerCase() === config.securityGroupEmail.toLowerCase()
    );
    if (isMatch) {
      matchedOffices.push(config);
    }
  }

  return matchedOffices;
}

/**
 * Check if user belongs to a specific office group
 */
export async function isUserInOffice(officeId: string): Promise<boolean> {
  const offices = await resolveUserOffices();
  return offices.some((office) => office.id === officeId);
}
