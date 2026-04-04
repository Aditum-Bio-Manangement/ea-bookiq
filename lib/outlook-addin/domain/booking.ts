import type { Room } from "../graph/places";
import { addRoomAttendee, setLocation, isRoomAlreadyAdded, isInOutlookContext } from "../office/appointment";
import { showNotification } from "../office/eventHandlers";

export interface BookingResult {
  success: boolean;
  message: string;
  room?: Room;
  isPreviewMode?: boolean;
}

/**
 * Book a room by adding it to the appointment
 */
export async function bookRoom(room: Room): Promise<BookingResult> {
  try {
    // Check if we're in Outlook context
    if (!isInOutlookContext()) {
      // Preview mode - show success message but note it's a simulation
      console.log("[EA BookIQ] Preview mode - simulating room booking for:", room.displayName);
      return {
        success: true,
        message: `${room.displayName} selected. Open this add-in from Outlook to add the room to your meeting.`,
        room,
        isPreviewMode: true,
      };
    }

    // Check if room is already added
    const alreadyAdded = await isRoomAlreadyAdded(room.emailAddress);
    if (alreadyAdded) {
      return {
        success: false,
        message: `${room.displayName} is already added to this meeting.`,
        room,
      };
    }

    // Add room as attendee
    await addRoomAttendee(room.displayName, room.emailAddress);

    // Set location
    await setLocation(room.displayName);

    showNotification(`${room.displayName} added to meeting.`);

    return {
      success: true,
      message: `${room.displayName} has been added to your meeting.`,
      room,
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to book room";
    console.error("[EA BookIQ] Booking failed:", error);

    return {
      success: false,
      message,
      room,
    };
  }
}

/**
 * Remove a room from the appointment (if supported)
 * Note: Office.js may not support removing specific attendees in all versions
 */
export async function removeRoom(room: Room): Promise<BookingResult> {
  // Note: Office.js doesn't have a direct removeAttendee method in older requirement sets
  // For now, we'll return a message indicating the user needs to remove manually
  showNotification(
    `To remove ${room.displayName}, please use the attendee list.`,
    "informational"
  );

  return {
    success: false,
    message: `Please remove ${room.displayName} manually from the attendee list.`,
    room,
  };
}
