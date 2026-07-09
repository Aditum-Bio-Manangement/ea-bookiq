import type { Room } from "../graph/places";
import { addRoomAttendee, addRoomLocation, removeRoomLocation, clearLocation, isInOutlookContext, removeRoomAttendee, isClassicOutlookDesktop, markRoomBooked } from "../office/appointment";
import { showNotification } from "../office/eventHandlers";

/**
 * Fire-and-forget persistence of the Booked marker. This is a best-effort
 * safety net (mainly for classic Outlook's unreliable attendee reads) and must
 * NEVER block or fail the booking flow — so we don't await it and swallow any
 * error. The real attendee/location APIs determine booking success.
 */
function persistBookedMarker(emailAddress: string, booked: boolean): void {
  void markRoomBooked(emailAddress, booked).catch((err) => {
    console.log("[AB Book IQ] Persist booked marker failed (non-critical):", err);
  });
}

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
      // Persist the Booked state so it survives refresh/reopen (mainly for
      // classic Outlook where attendee reads are unreliable). Fire-and-forget.
      persistBookedMarker(room.emailAddress, true);
      console.log("[AB Book IQ] Booked room as attendee + location:", room.displayName);
    } else if (mode === "attendee") {
      // Attendee only.
      await addRoomAttendee(room.displayName, room.emailAddress);
      // On classic Outlook desktop, adding a required attendee auto-fills the
      // location text with the room name. Clear ONLY that plain-text location
      // so the result is attendee-only (like new Outlook / OWA). We must NOT
      // use enhancedLocation.removeAsync here — on classic that is linked to
      // the resource and would strip the attendee too.
      if (isClassicOutlookDesktop()) {
        await clearLocation(room.displayName);
      }
      // Attendee-only is not "Booked", so make sure it isn't persisted as such.
      persistBookedMarker(room.emailAddress, false);
      console.log("[AB Book IQ] Added room as attendee only:", room.displayName);
    } else if (mode === "location") {
      // Location only (room resource). addRoomLocation never adds an attendee.
      await addRoomLocation(room.displayName, room.emailAddress);
      // Location-only is not "Booked".
      persistBookedMarker(room.emailAddress, false);
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
    // Clear the persisted Booked marker for this room (fire-and-forget).
    persistBookedMarker(room.emailAddress, false);

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
