import { resolveUserOffices } from "../graph/groups";
import { getConfiguredOffices, STORAGE_KEYS, type OfficeConfig } from "../config/offices";

export type OfficeResolutionResult =
  | { type: "single"; office: OfficeConfig }
  | { type: "multiple"; offices: OfficeConfig[]; cached?: OfficeConfig }
  | { type: "none" };

/**
 * Map an authorized office list to the task-pane flow.
 * Zero offices → Access Denied. One office → that site. Several (EA member of
 * both, or FTE-only) → picker, honoring a cached preference when present.
 */
export function resolveOfficeFromList(offices: OfficeConfig[]): OfficeResolutionResult {
  if (offices.length === 0) {
    return { type: "none" };
  }

  if (offices.length === 1) {
    return { type: "single", office: offices[0] };
  }

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
 * Resolve which office the user should default to
 */
export async function resolveOffice(): Promise<OfficeResolutionResult> {
  return resolveOfficeFromList(await resolveUserOffices());
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
  return getConfiguredOffices().find((office) => office.id === officeId);
}

/**
 * Get all available offices
 */
export function getAllOffices(): OfficeConfig[] {
  return getConfiguredOffices();
}
