import { getGraphClient } from "./graphClient";
import type { OfficeConfig } from "../config/offices";

export interface Room {
  id: string;
  displayName: string;
  emailAddress: string;
  capacity: number;
  building: string | null;
  floorNumber: number | null;
  floorLabel: string | null;
  label: string | null;
  tags: string[];
  audioDeviceName: string | null;
  videoDeviceName: string | null;
  displayDeviceName: string | null;
  isWheelChairAccessible: boolean;
}

export interface RoomList {
  id: string;
  displayName: string;
  emailAddress: string;
}

interface PlacesResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

/**
 * Get all room lists in the tenant
 */
export async function getRoomLists(): Promise<RoomList[]> {
  const client = await getGraphClient();

  const response: PlacesResponse<RoomList> = await client
    .api("/places/microsoft.graph.roomList")
    .get();

  return response.value;
}

/**
 * Get rooms for a specific room list
 */
export async function getRoomsFromRoomList(
  roomListId: string
): Promise<Room[]> {
  const client = await getGraphClient();

  const response: PlacesResponse<Room> = await client
    .api(`/places/${roomListId}/microsoft.graph.roomList/rooms`)
    .get();

  return response.value.map(normalizeRoom);
}

/**
 * Get rooms by building name
 */
export async function getRoomsByBuilding(building: string): Promise<Room[]> {
  const client = await getGraphClient();

  try {
    const response: PlacesResponse<Room> = await client
      .api("/places/microsoft.graph.room")
      .filter(`building eq '${building}'`)
      .get();

    return response.value.map(normalizeRoom);
  } catch {
    // Filter might not be supported, fall back to getting all rooms
    console.warn(
      "[EA BookIQ] Building filter not supported, fetching all rooms"
    );
    return getAllRoomsForOffice(building);
  }
}

/**
 * Get all rooms from the tenant
 */
export async function getAllRooms(): Promise<Room[]> {
  const client = await getGraphClient();

  try {
    const response: PlacesResponse<Room> = await client
      .api("/places/microsoft.graph.room")
      .top(100)
      .get();

    console.log("[v0] All rooms from Graph API:", response.value.map(r => ({
      displayName: r.displayName,
      emailAddress: r.emailAddress,
      building: r.building
    })));

    return response.value.map(normalizeRoom);
  } catch (err) {
    console.error("[v0] Failed to get rooms:", err);
    return [];
  }
}

/**
 * Get rooms for an office configuration
 * Filters rooms by display name containing the office location (e.g., "- Cambridge", "- Oakland")
 */
export async function getRoomsForOffice(office: OfficeConfig): Promise<Room[]> {
  // Get all rooms and filter by location in display name
  const allRooms = await getAllRooms();

  const locationPattern = office.name.toLowerCase();

  const filteredRooms = allRooms.filter((room) => {
    const displayNameLower = room.displayName.toLowerCase();
    const buildingLower = (room.building || "").toLowerCase();

    // Match by display name containing location (e.g., "Board Room - Cambridge")
    const nameMatch =
      displayNameLower.includes(`- ${locationPattern}`) ||
      displayNameLower.includes(`-${locationPattern}`) ||
      displayNameLower.endsWith(locationPattern);

    // Also match by building field if populated
    const buildingMatch = buildingLower === locationPattern;

    return nameMatch || buildingMatch;
  });

  console.log("[v0] Filtered rooms for", office.name, ":", filteredRooms.map(r => r.displayName));

  return filteredRooms;
}

/**
 * Get all rooms and filter by building (fallback method)
 */
async function getAllRoomsForOffice(building: string): Promise<Room[]> {
  const client = await getGraphClient();

  const response: PlacesResponse<Room> = await client
    .api("/places/microsoft.graph.room")
    .get();

  return response.value
    .filter(
      (room) => room.building?.toLowerCase() === building.toLowerCase()
    )
    .map(normalizeRoom);
}

/**
 * Normalize room data to ensure consistent structure
 */
function normalizeRoom(room: Room): Room {
  return {
    id: room.id,
    displayName: room.displayName || "Unknown Room",
    emailAddress: room.emailAddress,
    capacity: room.capacity || 0,
    building: room.building || null,
    floorNumber: room.floorNumber || null,
    floorLabel: room.floorLabel || null,
    label: room.label || null,
    tags: room.tags || [],
    audioDeviceName: room.audioDeviceName || null,
    videoDeviceName: room.videoDeviceName || null,
    displayDeviceName: room.displayDeviceName || null,
    isWheelChairAccessible: room.isWheelChairAccessible || false,
  };
}
