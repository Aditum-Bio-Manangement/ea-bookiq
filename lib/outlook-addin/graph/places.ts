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
      "[Room Assist] Building filter not supported, fetching all rooms"
    );
    return getAllRoomsForOffice(building);
  }
}

/**
 * Get rooms for an office configuration
 * Tries room list first, then falls back to building filter
 */
export async function getRoomsForOffice(office: OfficeConfig): Promise<Room[]> {
  // If we have a room list ID, use it
  if (office.roomListId) {
    return getRoomsFromRoomList(office.roomListId);
  }

  // Try to find a room list matching the office
  const roomLists = await getRoomLists();
  const matchingList = roomLists.find(
    (list) =>
      list.displayName.toLowerCase().includes(office.name.toLowerCase()) ||
      list.emailAddress.toLowerCase().includes(office.name.toLowerCase())
  );

  if (matchingList) {
    return getRoomsFromRoomList(matchingList.id);
  }

  // Fall back to building filter
  return getRoomsByBuilding(office.building);
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
