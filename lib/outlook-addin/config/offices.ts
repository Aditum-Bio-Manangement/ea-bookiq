// Office configuration for Cambridge and Oakland locations
// Maps Entra security groups (by object id) to room lists

export interface OfficeConfig {
  id: string;
  name: string;
  displayName: string;
  securityGroupId: string;
  securityGroupEmail: string;
  roomListId?: string; // Will be populated from Graph Places API
  building: string;
}

/** Default Entra object ids. Override with NEXT_PUBLIC_* env vars at build time. */
export const DEFAULT_GROUP_IDS = {
  cambridge: "76ea470c-2fcb-4924-8e87-25f0b85fb2ad",
  oakland: "fd203838-00b5-43a5-b3d9-93350d2a8db9",
  fte: "232beced-9e97-4c41-ab80-83d95632bdfe",
} as const;

export const OFFICE_CONFIGS: Record<string, OfficeConfig> = {
  cambridge: {
    id: "cambridge",
    name: "Cambridge",
    displayName: "Cambridge Office",
    securityGroupId: DEFAULT_GROUP_IDS.cambridge,
    securityGroupEmail: "ea-cambridge@aditumbio.com",
    building: "Cambridge",
  },
  oakland: {
    id: "oakland",
    name: "Oakland",
    displayName: "Oakland Office",
    securityGroupId: DEFAULT_GROUP_IDS.oakland,
    securityGroupEmail: "ea-oakland@aditumbio.com",
    building: "Oakland",
  },
};

function normalizeGroupId(value: string | undefined | null): string | undefined {
  const trimmed = value?.trim();
  return trimmed ? trimmed.toLowerCase() : undefined;
}

function parseGroupIdList(value: string | undefined, fallback: string[]): string[] {
  if (!value?.trim()) {
    return fallback.map((id) => id.toLowerCase());
  }

  const parsed = value
    .split(",")
    .map((id) => id.trim().toLowerCase())
    .filter(Boolean);

  return parsed.length > 0 ? parsed : fallback.map((id) => id.toLowerCase());
}

/**
 * Group object ids used for authorization and office detection.
 * NEXT_PUBLIC_AUTHORIZED_GROUP_IDS is a comma-separated list (ABM-SG-FTE by default).
 */
export function getConfiguredGroupIds(): {
  cambridge: string;
  oakland: string;
  authorized: string[];
} {
  return {
    cambridge:
      normalizeGroupId(process.env.NEXT_PUBLIC_EA_CAMBRIDGE_GROUP_ID) ??
      DEFAULT_GROUP_IDS.cambridge,
    oakland:
      normalizeGroupId(process.env.NEXT_PUBLIC_EA_OAKLAND_GROUP_ID) ??
      DEFAULT_GROUP_IDS.oakland,
    authorized: parseGroupIdList(
      process.env.NEXT_PUBLIC_AUTHORIZED_GROUP_IDS,
      [DEFAULT_GROUP_IDS.fte]
    ),
  };
}

export function getConfiguredOffices(): OfficeConfig[] {
  const ids = getConfiguredGroupIds();
  return [
    { ...OFFICE_CONFIGS.cambridge, securityGroupId: ids.cambridge },
    { ...OFFICE_CONFIGS.oakland, securityGroupId: ids.oakland },
  ];
}

const getRedirectUri = () => {
  const redirectUri = process.env.NEXT_PUBLIC_REDIRECT_URI;
  if (!redirectUri) {
    throw new Error(
      "Missing NEXT_PUBLIC_REDIRECT_URI. Set it to the hosted task pane URL before build or start."
    );
  }
  return redirectUri;
};

export const MSAL_CONFIG = {
  clientId: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID || "demo-client-id",
  tenantId: process.env.NEXT_PUBLIC_AZURE_TENANT_ID || "common",
  get redirectUri() {
    return getRedirectUri();
  },
};

// Check if MSAL is properly configured
export const isMsalConfigured = () => {
  return Boolean(
    process.env.NEXT_PUBLIC_AZURE_CLIENT_ID &&
      process.env.NEXT_PUBLIC_AZURE_TENANT_ID &&
      process.env.NEXT_PUBLIC_REDIRECT_URI
  );
};

export const GRAPH_SCOPES = [
  "openid",
  "profile",
  "email",
  "offline_access",
  "User.Read",
  "Calendars.Read.Shared",
  "Place.Read.All",
  "GroupMember.Read.All",
];

// Local storage keys
export const STORAGE_KEYS = {
  OFFICE_PREFERENCE: "room-assist-office-preference",
  LAST_SELECTED_ROOM: "room-assist-last-room",
};
