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
 * Get current attendees of the appointment (including resources in newer requirement sets)
 */
export async function getCurrentAttendees(): Promise<Attendee[]> {
  return new Promise((resolve) => {
    const item = getMailboxItem();
    if (!item) {
      resolve([]);
      return;
    }

    const attendees: Attendee[] = [];
    let pendingCount = 2; // required + optional attendees

    // Check if resources API is available (RequirementSet 1.7+)
    // In classic Outlook, resources might be accessible through a different API
    const hasResources = item && typeof (item as any).resources?.getAsync === 'function';
    if (hasResources) {
      pendingCount = 3; // also fetch resources
    }

    console.log("[AB Book IQ] hasResources API:", hasResources);

    const checkComplete = () => {
      pendingCount--;
      if (pendingCount === 0) {
        console.log("[AB Book IQ] All attendees fetched:", attendees.length, "total");
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
        console.log("[AB Book IQ] Required attendees:", result.value.length, result.value.map((a: any) => a.displayName));
      } else {
        console.log("[AB Book IQ] Failed to get required attendees:", result.error?.message);
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
        console.log("[AB Book IQ] Optional attendees:", result.value.length);
      } else {
        console.log("[AB Book IQ] Failed to get optional attendees:", result.error?.message);
      }
      checkComplete();
    });

    // Try to get resources (rooms/equipment) if available
    if (hasResources) {
      (item as any).resources.getAsync((result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          for (const res of result.value) {
            attendees.push({
              displayName: res.displayName,
              emailAddress: res.emailAddress,
              recipientType: "resource",
            });
          }
          console.log("[AB Book IQ] Resources (rooms):", result.value.length, result.value.map((r: any) => r.displayName));
        } else {
          console.log("[AB Book IQ] Failed to get resources:", result.error?.message);
        }
        checkComplete();
      });
    }
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
 * Get list of room email addresses that are already added to the meeting
 * Checks attendees/resources AND location field as fallback
 */
export async function getAddedRoomEmails(allRoomEmails: string[], allRoomNames?: string[]): Promise<Set<string>> {
  const attendees = await getCurrentAttendees();
  const attendeeEmails = new Set(attendees.map(a => a.emailAddress.toLowerCase()));
  const addedRooms = new Set<string>();

  console.log("[AB Book IQ] Checking for added rooms. Attendees:",
    attendees.map(a => ({ name: a.displayName, email: a.emailAddress, type: a.recipientType }))
  );

  // Check attendee emails
  for (const roomEmail of allRoomEmails) {
    if (attendeeEmails.has(roomEmail.toLowerCase())) {
      addedRooms.add(roomEmail.toLowerCase());
      console.log(`[AB Book IQ] Found booked room via attendee: ${roomEmail}`);
    }
  }

  // Also check location as a fallback (in case resources API doesn't work)
  if (allRoomNames && allRoomNames.length > 0) {
    const location = await getLocation();
    if (location) {
      console.log(`[AB Book IQ] Current location: "${location}"`);
      // Check if location matches any room name
      for (let i = 0; i < allRoomNames.length; i++) {
        const roomName = allRoomNames[i];
        const roomEmail = allRoomEmails[i];
        // Check if location contains the room name (case-insensitive)
        if (location.toLowerCase().includes(roomName.toLowerCase()) ||
          roomName.toLowerCase().includes(location.toLowerCase())) {
          if (!addedRooms.has(roomEmail.toLowerCase())) {
            addedRooms.add(roomEmail.toLowerCase());
            console.log(`[AB Book IQ] Found booked room via location match: ${roomName} -> ${roomEmail}`);
          }
        }
      }
    }
  }

  return addedRooms;
}

/**
 * Remove a room attendee by setting attendees list without that room
 */
export async function removeRoomAttendee(emailToRemove: string): Promise<void> {
  return new Promise(async (resolve, reject) => {
    const item = getMailboxItem();
    if (!item) {
      reject(new Error("No appointment item available"));
      return;
    }

    try {
      // Get current required attendees
      const attendees = await getCurrentAttendees();

      // Filter out the room to remove
      const remainingRequired = attendees
        .filter(a => a.recipientType === "required" && a.emailAddress.toLowerCase() !== emailToRemove.toLowerCase())
        .map(a => ({ displayName: a.displayName, emailAddress: a.emailAddress }));

      // Set the filtered list (this replaces all required attendees)
      item.requiredAttendees.setAsync(remainingRequired, (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message || "Failed to remove attendee"));
        }
      });
    } catch (err) {
      reject(err);
    }
  });
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
