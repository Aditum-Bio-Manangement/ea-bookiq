import type { Room } from "../graph/places";
import { addRoomAttendee, addRoomLocation, removeRoomLocation, isInOutlookContext, removeRoomAttendee } from "../office/appointment";
import { showNotification } from "../office/eventHandlers";

export type BookingMode = "both" | "attendee" | "location";

export interface BookingResult {
  success: boolean;
  message: string;
  room?: Room;
  isPreviewMode?: boolean;
  replacedRoomEmail?: string;
  mode?: BookingMode;
}

/**
 * Book a room by adding it to the appointment
 * Removes any existing rooms first, then adds the new room using the best available APIs
 * @param mode - "both" (default): add as attendee + set location, "attendee": only add as attendee, "location": only set location
 */
export async function bookRoom(
  room: Room,
  allRoomEmails?: string[],
  mode: BookingMode = "both"
): Promise<BookingResult> {
  try {
    // Check if we're in Outlook context
    if (!isInOutlookContext()) {
      console.log("[AB Book IQ] Preview mode - simulating room booking for:", room.displayName);
      return {
        success: true,
        message: `${room.displayName} selected. Open this add-in from Outlook to add the room to your meeting.`,
        room,
        isPreviewMode: true,
        mode,
      };
    }

    // We allow multiple rooms to be added (as attendees and/or locations), so
    // we never remove existing rooms here. Each mode adds exactly what it says.
    if (mode === "both") {
      // Book = add as attendee AND as the room-resource location.
      await addRoomAttendee(room.displayName, room.emailAddress);
      await addRoomLocation(room.displayName, room.emailAddress);
      console.log("[AB Book IQ] Booked room as attendee + location:", room.displayName);
    } else if (mode === "attendee") {
      // Attendee only. We intentionally do NOT remove the location afterwards.
      // On classic Outlook desktop the room is a resolved resource whose
      // attendee entry and location/enhancedLocation entry are linked, so
      // removing the location also strips the attendee (the room then ends up
      // added nowhere). New Outlook / OWA already keep the fields independent,
      // so adding only the attendee is correct everywhere.
      await addRoomAttendee(room.displayName, room.emailAddress);
      console.log("[AB Book IQ] Added room as attendee only:", room.displayName);
    } else if (mode === "location") {
      // Location only (room resource). addRoomLocation never adds an attendee.
      await addRoomLocation(room.displayName, room.emailAddress);
      console.log("[AB Book IQ] Added room as location only:", room.displayName);
    }

    // Build notification message
    let message: string;
    if (mode === "both") {
      message = `${room.displayName} added to meeting.`;
    } else if (mode === "attendee") {
      message = `${room.displayName} added as attendee.`;
    } else {
      message = `Location set to ${room.displayName}.`;
    }

    showNotification(message);

    return {
      success: true,
      message,
      room,
      mode,
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to book room";
    console.error("[AB Book IQ] Booking failed:", error);

    return {
      success: false,
      message,
      room,
      mode,
    };
  }
}

/**
 * Remove a room from the appointment
 * Removes from both required attendees and resources, and clears location
 */
export async function unbookRoom(room: Room, allRoomEmails?: string[]): Promise<BookingResult> {
  try {
    // Check if we're in Outlook context
    if (!isInOutlookContext()) {
      return {
        success: true,
        message: `${room.displayName} removed. Open this add-in from Outlook to update the meeting.`,
        room,
        isPreviewMode: true,
      };
    }

    // Remove ONLY this room (since multiple rooms may be booked): take it out
    // of the attendees and out of the location.
    await removeRoomAttendee(room.emailAddress);
    await removeRoomLocation(room.displayName, room.emailAddress);

    console.log("[AB Book IQ] Room unbooked:", room.displayName);
    showNotification(`${room.displayName} removed from meeting.`);

    return {
      success: true,
      message: `${room.displayName} removed from meeting.`,
      room,
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to remove room";
    console.error("[AB Book IQ] Unbook failed:", error);

    return {
      success: false,
      message,
      room,
    };
  }
}
