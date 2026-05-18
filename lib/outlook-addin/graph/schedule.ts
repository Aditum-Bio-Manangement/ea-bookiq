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
 * Parse a schedule item datetime into a Date object
 * The Graph API returns datetime in the format specified by the timeZone field
 */
function parseScheduleDateTime(timeSlot: TimeSlot): Date {
  // If the datetime already has timezone info (Z or +/-offset), parse directly
  if (timeSlot.dateTime.endsWith("Z") || /[+-]\d{2}:\d{2}$/.test(timeSlot.dateTime)) {
    return new Date(timeSlot.dateTime);
  }

  // For datetimes without timezone suffix, treat as the specified timezone
  // Since we request in user's local timezone, we can parse as local
  // The datetime format is typically "2026-05-22T13:30:00.0000000"
  return new Date(timeSlot.dateTime);
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

  console.log("[AB Book IQ] Schedule API request:", {
    requestedStart: startTime.toISOString(),
    requestedEnd: endTime.toISOString(),
    timeZone: timeZone,
    roomCount: rooms.length,
  });

  const availability: RoomAvailability[] = [];

  for (const scheduleInfo of scheduleResponse.value) {
    const room = rooms.find(
      (r) => r.emailAddress.toLowerCase() === scheduleInfo.scheduleId.toLowerCase()
    );

    if (!room) continue;

    console.log(`[AB Book IQ] Schedule for ${room.displayName}:`, {
      availabilityView: scheduleInfo.availabilityView,
      scheduleItems: scheduleInfo.scheduleItems.map(item => ({
        status: item.status,
        start: item.start,
        end: item.end,
        subject: item.subject || '(no subject)'
      }))
    });

    // Check if any busy/tentative items actually overlap with the requested time window
    let hasConflict = false;

    for (const item of scheduleInfo.scheduleItems) {
      if (item.status !== "busy" && item.status !== "tentative") {
        continue;
      }

      // Parse the schedule item times using proper timezone handling
      const itemStart = parseScheduleDateTime(item.start);
      const itemEnd = parseScheduleDateTime(item.end);

      // Check if this busy time overlaps with our requested window
      const overlaps = timeRangesOverlap(startTime, endTime, itemStart, itemEnd);

      console.log(`[AB Book IQ] Checking overlap for ${room.displayName}:`, {
        itemStart: itemStart.toISOString(),
        itemEnd: itemEnd.toISOString(),
        requestedStart: startTime.toISOString(),
        requestedEnd: endTime.toISOString(),
        overlaps: overlaps,
        status: item.status,
      });

      if (overlaps) {
        hasConflict = true;
        break;
      }
    }

    // Also check the availabilityView string which shows busy status
    // '0' = free, '1' = tentative, '2' = busy, '3' = OOF, '4' = working elsewhere
    if (!hasConflict && scheduleInfo.availabilityView) {
      const busyIndicators = ['1', '2', '3']; // tentative, busy, OOF
      const hasBusyView = scheduleInfo.availabilityView.split('').some(char => busyIndicators.includes(char));
      if (hasBusyView) {
        console.log(`[AB Book IQ] ${room.displayName} marked unavailable via availabilityView: ${scheduleInfo.availabilityView}`);
        hasConflict = true;
      }
    }

    console.log(`[AB Book IQ] ${room.displayName} final availability: ${hasConflict ? 'UNAVAILABLE' : 'AVAILABLE'}`);

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
