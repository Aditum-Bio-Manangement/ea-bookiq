import type { Room } from "../graph/places";
import { addRoomAttendee, setLocation, isRoomAlreadyAdded, isInOutlookContext, removeRoomAttendee, getAddedRoomEmails, clearLocation } from "../office/appointment";
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
 * If another room is already booked as attendee, it will be replaced
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
      // Preview mode - show success message but note it's a simulation
      console.log("[AB Book IQ] Preview mode - simulating room booking for:", room.displayName);
      return {
        success: true,
        message: `${room.displayName} selected. Open this add-in from Outlook to add the room to your meeting.`,
        room,
        isPreviewMode: true,
        mode,
      };
    }

    let replacedRoomEmail: string | undefined;

    // Handle attendee addition (for "both" or "attendee" modes)
    if (mode === "both" || mode === "attendee") {
      // Check if this specific room is already added as attendee
      const alreadyAdded = await isRoomAlreadyAdded(room.emailAddress);

      if (alreadyAdded) {
        // Room is already added - prevent duplicate booking
        if (mode === "attendee") {
          return {
            success: false,
            message: `${room.displayName} is already added as an attendee.`,
            room,
            mode,
          };
        }
        // In "both" mode, if room is already added, just update location and return success
        // without adding it again
        if (mode === "both") {
          await setLocation(room.displayName);
          showNotification(`${room.displayName} location updated.`);
          return {
            success: true,
            message: `${room.displayName} is already booked. Location updated.`,
            room,
            mode,
          };
        }
      } else {
        // Room is not added yet - check for and replace any existing room first
        if (mode === "both" && allRoomEmails && allRoomEmails.length > 0) {
          const addedRooms = await getAddedRoomEmails(allRoomEmails);
          if (addedRooms.size > 0) {
            // Remove existing room(s) - typically there's only one
            for (const existingRoomEmail of addedRooms) {
              // Don't remove the room we're trying to add
              if (existingRoomEmail.toLowerCase() !== room.emailAddress.toLowerCase()) {
                console.log("[AB Book IQ] Removing existing room:", existingRoomEmail);
                await removeRoomAttendee(existingRoomEmail);
                replacedRoomEmail = existingRoomEmail;
              }
            }
          }
        }

        // Add room as attendee
        await addRoomAttendee(room.displayName, room.emailAddress);
      }
    }

    // Handle location setting (for "both" or "location" modes)
    if (mode === "both" || mode === "location") {
      await setLocation(room.displayName);
    }

    // Build notification message
    let message: string;
    if (mode === "both") {
      message = replacedRoomEmail
        ? `${room.displayName} replaced previous room.`
        : `${room.displayName} added to meeting.`;
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
      replacedRoomEmail,
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
 */
export async function unbookRoom(room: Room): Promise<BookingResult> {
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

    // Remove the room attendee
    await removeRoomAttendee(room.emailAddress);

    // Clear the location if it matches this room
    await clearLocation(room.displayName);

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
