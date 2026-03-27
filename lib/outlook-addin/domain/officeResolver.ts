import { resolveUserOffices } from "../graph/groups";
import { OFFICE_CONFIGS, STORAGE_KEYS, type OfficeConfig } from "../config/offices";

export type OfficeResolutionResult =
  | { type: "single"; office: OfficeConfig }
  | { type: "multiple"; offices: OfficeConfig[]; cached?: OfficeConfig }
  | { type: "none" };

/**
 * Resolve which office the user should default to
 */
export async function resolveOffice(): Promise<OfficeResolutionResult> {
  const offices = await resolveUserOffices();

  if (offices.length === 0) {
    return { type: "none" };
  }

  if (offices.length === 1) {
    return { type: "single", office: offices[0] };
  }

  // Multiple offices - check for cached preference
  const cachedOfficeId = getCachedOfficePreference();
  const cachedOffice = cachedOfficeId
    ? offices.find((o) => o.id === cachedOfficeId)
    : undefined;

  return {
    type: "multiple",
    offices,
    cached: cachedOffice,
  };
}

/**
 * Get cached office preference from local storage
 */
export function getCachedOfficePreference(): string | null {
  if (typeof localStorage === "undefined") return null;
  return localStorage.getItem(STORAGE_KEYS.OFFICE_PREFERENCE);
}

/**
 * Save office preference to local storage
 */
export function setCachedOfficePreference(officeId: string): void {
  if (typeof localStorage === "undefined") return;
  localStorage.setItem(STORAGE_KEYS.OFFICE_PREFERENCE, officeId);
}

/**
 * Clear cached office preference
 */
export function clearCachedOfficePreference(): void {
  if (typeof localStorage === "undefined") return;
  localStorage.removeItem(STORAGE_KEYS.OFFICE_PREFERENCE);
}

/**
 * Get office config by ID
 */
export function getOfficeById(officeId: string): OfficeConfig | undefined {
  return OFFICE_CONFIGS[officeId];
}

/**
 * Get all available offices
 */
export function getAllOffices(): OfficeConfig[] {
  return Object.values(OFFICE_CONFIGS);
}
