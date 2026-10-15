// Office configuration for Cambridge and Oakland locations
// Maps security groups to room lists

export interface OfficeConfig {
  id: string;
  name: string;
  displayName: string;
  securityGroupEmail: string;
  roomListId?: string; // Will be populated from Graph Places API
  building: string;
}

export const OFFICE_CONFIGS: Record<string, OfficeConfig> = {
  cambridge: {
    id: "cambridge",
    name: "Cambridge",
    displayName: "Cambridge Office",
    securityGroupEmail: "ea-cambridge@aditumbio.com",
    building: "Cambridge",
  },
  oakland: {
    id: "oakland",
    name: "Oakland",
    displayName: "Oakland Office",
    securityGroupEmail: "ea-oakland@aditumbio.com",
    building: "Oakland",
  },
};

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
