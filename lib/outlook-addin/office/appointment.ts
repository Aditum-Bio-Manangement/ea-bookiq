/// <reference types="@types/office-js" />

export interface MeetingWindow {
  start: Date | null;
  end: Date | null;
  complete: boolean;
  timeZone: string;
}

export interface Attendee {
  displayName: string;
  emailAddress: string;
  recipientType: "required" | "optional" | "resource";
}

/**
 * Safely get the current mailbox item as AppointmentCompose, handling cases where Office.js is not loaded
 */
function getMailboxItem(): Office.AppointmentCompose | null {
  try {
    if (typeof Office === "undefined" || !Office.context?.mailbox?.item) {
      return null;
    }
    // Cast to AppointmentCompose since this add-in only runs in appointment compose mode
    return Office.context.mailbox.item as Office.AppointmentCompose;
  } catch {
    return null;
  }
}

/**
 * Check if we're running in a valid Outlook context
 */
export function isInOutlookContext(): boolean {
  return getMailboxItem() !== null;
}

/**
 * Get the current meeting time window from the appointment being composed
 */
export async function getMeetingWindow(): Promise<MeetingWindow> {
  return new Promise((resolve) => {
    const item = getMailboxItem();
    if (!item) {
      resolve({
        start: null,
        end: null,
        complete: false,
        timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
      });
      return;
    }

    let startTime: Date | null = null;
    let endTime: Date | null = null;
    let resolved = 0;

    const checkComplete = () => {
      resolved++;
      if (resolved === 2) {
        resolve({
          start: startTime,
          end: endTime,
          complete: startTime !== null && endTime !== null,
          timeZone: Office.context?.mailbox?.userProfile?.timeZone ||
            Intl.DateTimeFormat().resolvedOptions().timeZone,
        });
      }
    };

    item.start.getAsync((result: any) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        startTime = result.value;
      }
      checkComplete();
    });

    item.end.getAsync((result: any) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        endTime = result.value;
      }
      checkComplete();
    });
  });
}

/**
 * Get current attendees of the appointment
 */
export async function getCurrentAttendees(): Promise<Attendee[]> {
  return new Promise((resolve) => {
    const item = getMailboxItem();
    if (!item) {
      resolve([]);
      return;
    }

    const attendees: Attendee[] = [];
    let resolved = 0;

    const checkComplete = () => {
      resolved++;
      if (resolved === 2) {
        resolve(attendees);
      }
    };

    item.requiredAttendees.getAsync((result: any) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        for (const att of result.value) {
          attendees.push({
            displayName: att.displayName,
            emailAddress: att.emailAddress,
            recipientType: "required",
          });
        }
      }
      checkComplete();
    });

    item.optionalAttendees.getAsync((result: any) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        for (const att of result.value) {
          attendees.push({
            displayName: att.displayName,
            emailAddress: att.emailAddress,
            recipientType: "optional",
          });
        }
      }
      checkComplete();
    });
  });
}

/**
 * Add a room as a required attendee
 */
export async function addRoomAttendee(
  displayName: string,
  emailAddress: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    const item = getMailboxItem();
    if (!item) {
      reject(new Error("No appointment item available"));
      return;
    }

    item.requiredAttendees.addAsync(
      [{ displayName, emailAddress }],
      (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message || "Failed to add attendee"));
        }
      }
    );
  });
}

/**
 * Set the meeting location
 */
export async function setLocation(location: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const item = getMailboxItem();
    if (!item) {
      reject(new Error("No appointment item available"));
      return;
    }

    item.location.setAsync(location, (result: any) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error?.message || "Failed to set location"));
      }
    });
  });
}

/**
 * Get current location
 */
export async function getLocation(): Promise<string> {
  return new Promise((resolve) => {
    const item = getMailboxItem();
    if (!item) {
      resolve("");
      return;
    }

    item.location.getAsync((result: any) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        resolve("");
      }
    });
  });
}

/**
 * Check if a room is already added as an attendee
 */
export async function isRoomAlreadyAdded(emailAddress: string): Promise<boolean> {
  const attendees = await getCurrentAttendees();
  return attendees.some(
    (a) => a.emailAddress.toLowerCase() === emailAddress.toLowerCase()
  );
}

/**
 * Get organizer email
 */
export function getOrganizerEmail(): string {
  try {
    return Office.context?.mailbox?.userProfile?.emailAddress || "";
  } catch {
    return "";
  }
}

/**
 * Get organizer display name
 */
export function getOrganizerDisplayName(): string {
  try {
    return Office.context?.mailbox?.userProfile?.displayName || "";
  } catch {
    return "";
  }
}
