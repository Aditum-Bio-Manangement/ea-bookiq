import type { Room } from "../graph/places";
import { addRoomAttendee, setLocation, isRoomAlreadyAdded, isInOutlookContext, removeRoomAttendee, removeAllRooms } from "../office/appointment";
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
 * Removes any existing rooms first, then adds the new room to both required attendees AND resources
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

    let replacedRoom = false;

    // Handle attendee addition (for "both" or "attendee" modes)
    if (mode === "both" || mode === "attendee") {
      // First, remove ALL existing rooms from the meeting to ensure only one room is booked
      if (allRoomEmails && allRoomEmails.length > 0) {
        console.log("[AB Book IQ] Removing all existing rooms before booking new one");
        await removeAllRooms(allRoomEmails);
        replacedRoom = true;
      }

      // Now add the new room as attendee (to both required attendees and resources)
      console.log("[AB Book IQ] Adding room:", room.displayName);
      await addRoomAttendee(room.displayName, room.emailAddress);
    }

    // Handle location setting (for "both" or "location" modes)
    if (mode === "both" || mode === "location") {
      await setLocation(room.displayName);
      console.log("[AB Book IQ] Set location to:", room.displayName);
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

    // Remove the specific room (or all rooms if allRoomEmails provided)
    if (allRoomEmails && allRoomEmails.length > 0) {
      // Remove all rooms to ensure clean state
      await removeAllRooms(allRoomEmails);
    } else {
      // Remove just this specific room
      await removeRoomAttendee(room.emailAddress);
      await setLocation("");
    }

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
