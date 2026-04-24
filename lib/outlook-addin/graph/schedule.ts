import { getGraphClient } from "./graphClient";
import type { Room } from "./places";

export interface TimeSlot {
  dateTime: string;
  timeZone: string;
}

export interface ScheduleItem {
  status: "free" | "busy" | "tentative" | "workingElsewhere" | "oof" | "unknown";
  start: TimeSlot;
  end: TimeSlot;
  subject?: string;
  location?: string;
  isPrivate?: boolean;
}

export interface ScheduleInfo {
  scheduleId: string;
  availabilityView: string;
  scheduleItems: ScheduleItem[];
  workingHours?: {
    daysOfWeek: string[];
    startTime: string;
    endTime: string;
    timeZone: { name: string };
  };
}

export interface GetScheduleResponse {
  value: ScheduleInfo[];
}

export interface RoomAvailability {
  room: Room;
  isAvailable: boolean;
  scheduleItems: ScheduleItem[];
  availabilityView: string;
}

/**
 * Get free/busy schedule for multiple rooms
 */
export async function getSchedule(
  roomEmails: string[],
  startTime: Date,
  endTime: Date,
  timeZone: string = Intl.DateTimeFormat().resolvedOptions().timeZone
): Promise<GetScheduleResponse> {
  const client = await getGraphClient();

  const requestBody = {
    schedules: roomEmails,
    startTime: {
      dateTime: startTime.toISOString().replace("Z", ""),
      timeZone: timeZone,
    },
    endTime: {
      dateTime: endTime.toISOString().replace("Z", ""),
      timeZone: timeZone,
    },
    availabilityViewInterval: 30, // 30-minute intervals
  };

  const response: GetScheduleResponse = await client
    .api("/me/calendar/getSchedule")
    .post(requestBody);

  return response;
}

/**
 * Check if two time ranges overlap
 */
function timeRangesOverlap(
  start1: Date, end1: Date,
  start2: Date, end2: Date
): boolean {
  // Two ranges overlap if one starts before the other ends AND ends after the other starts
  return start1 < end2 && end1 > start2;
}

/**
 * Check availability for a list of rooms during a specific time window
 */
export async function checkRoomAvailability(
  rooms: Room[],
  startTime: Date,
  endTime: Date,
  timeZone?: string
): Promise<RoomAvailability[]> {
  if (rooms.length === 0) {
    return [];
  }

  const roomEmails = rooms.map((r) => r.emailAddress);
  const scheduleResponse = await getSchedule(
    roomEmails,
    startTime,
    endTime,
    timeZone
  );

  const availability: RoomAvailability[] = [];

  for (const scheduleInfo of scheduleResponse.value) {
    const room = rooms.find(
      (r) => r.emailAddress.toLowerCase() === scheduleInfo.scheduleId.toLowerCase()
    );

    if (!room) continue;

    // Check if any busy/tentative items actually overlap with the requested time window
    const hasConflict = scheduleInfo.scheduleItems.some((item) => {
      if (item.status !== "busy" && item.status !== "tentative") {
        return false;
      }

      // Parse the schedule item times
      const itemStart = new Date(item.start.dateTime + (item.start.dateTime.endsWith("Z") ? "" : "Z"));
      const itemEnd = new Date(item.end.dateTime + (item.end.dateTime.endsWith("Z") ? "" : "Z"));

      // Check if this busy time overlaps with our requested window
      return timeRangesOverlap(startTime, endTime, itemStart, itemEnd);
    });

    availability.push({
      room,
      isAvailable: !hasConflict,
      scheduleItems: scheduleInfo.scheduleItems,
      availabilityView: scheduleInfo.availabilityView,
    });
  }

  return availability;
}

/**
 * Filter rooms to only those that are available
 */
export function filterAvailableRooms(
  availability: RoomAvailability[]
): RoomAvailability[] {
  return availability.filter((a) => a.isAvailable);
}

/**
 * Sort rooms by preference (available first, then by capacity)
 */
export function sortRoomsByPreference(
  availability: RoomAvailability[],
  preferredCapacity?: number
): RoomAvailability[] {
  return [...availability].sort((a, b) => {
    // Available rooms first
    if (a.isAvailable !== b.isAvailable) {
      return a.isAvailable ? -1 : 1;
    }

    // If preferred capacity is set, prefer rooms closest to that capacity
    if (preferredCapacity) {
      const aDiff = Math.abs(a.room.capacity - preferredCapacity);
      const bDiff = Math.abs(b.room.capacity - preferredCapacity);
      if (aDiff !== bDiff) {
        return aDiff - bDiff;
      }
    }

    // Sort by capacity ascending (smaller rooms first for efficiency)
    return a.room.capacity - b.room.capacity;
  });
}
