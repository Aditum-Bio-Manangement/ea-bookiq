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
 * Add a room ONLY as a required attendee (no location changes).
 *
 * This uses requiredAttendees.addAsync, which behaves consistently on every
 * Outlook version: it adds the room to the attendee/Required line and does NOT
 * touch the location field.
 *
 * IMPORTANT: We intentionally do NOT use enhancedLocation.addAsync here. On
 * classic Outlook desktop that API also adds the room as a resource (which
 * shows up as a second attendee), causing duplicate entries. On new Outlook it
 * only sets the location. To get identical behavior everywhere we avoid it and
 * set the location separately as text (see setRoomAsLocation / bookRoom).
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
          console.log("[AB Book IQ] Added room to requiredAttendees:", displayName);
          resolve();
        } else {
          console.log("[AB Book IQ] Failed to add to requiredAttendees:", result.error?.message);
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
 * Set a room as the meeting location WITHOUT adding it as an attendee.
 *
 * Uses the plain text location field on every Outlook version. We deliberately
 * avoid enhancedLocation.addAsync because on classic Outlook desktop it also
 * adds the room as a resource (a second attendee), which would break the
 * "location only" behavior. Setting the location text gives identical,
 * attendee-free results on all clients.
 */
export async function setRoomAsLocation(
  displayName: string,
  _emailAddress: string
): Promise<void> {
  await setLocation(displayName);
  console.log("[AB Book IQ] Set room as location (text):", displayName);
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
 * Clear the location if it matches the given room name
 */
export async function clearLocation(roomName: string): Promise<void> {
  const currentLocation = await getLocation();
  // Only clear if the location contains the room name
  if (currentLocation.toLowerCase().includes(roomName.toLowerCase())) {
    await setLocation("");
  }
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
 * Handles both requiredAttendees and resources collections
 */
export async function removeRoomAttendee(emailToRemove: string): Promise<void> {
  return new Promise(async (resolve, reject) => {
    const item = getMailboxItem();
    if (!item) {
      reject(new Error("No appointment item available"));
      return;
    }

    try {
      const normalizedEmail = emailToRemove.toLowerCase();
      const organizerEmail = getOrganizerEmail().toLowerCase();

      // Check if resources API is available (for rooms/equipment)
      const hasResources = typeof (item as any).resources?.getAsync === 'function';

      if (hasResources) {
        // Remove from resources
        await new Promise<void>((res) => {
          (item as any).resources.getAsync((result: any) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const currentResources = result.value || [];
              const filtered = currentResources
                .filter((r: any) => r.emailAddress.toLowerCase() !== normalizedEmail)
                .map((r: any) => ({ displayName: r.displayName, emailAddress: r.emailAddress }));

              if (filtered.length < currentResources.length) {
                (item as any).resources.setAsync(filtered, (setResult: any) => {
                  if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("[AB Book IQ] Removed room from resources:", emailToRemove);
                  }
                  res();
                });
              } else {
                res();
              }
            } else {
              res();
            }
          });
        });
      }

      // Remove from required attendees
      await new Promise<void>((res) => {
        item.requiredAttendees.getAsync((result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const currentAttendees = result.value || [];
            const filtered = currentAttendees
              .filter((a: any) => {
                const email = a.emailAddress.toLowerCase();
                // Remove the specified room
                if (email === normalizedEmail) return false;
                // Also filter out the organizer if they're accidentally in the list
                if (email === organizerEmail) return false;
                return true;
              })
              // Re-create objects with only required properties for old Outlook compatibility
              .map((a: any) => ({ displayName: a.displayName, emailAddress: a.emailAddress }));

            if (filtered.length < currentAttendees.length) {
              item.requiredAttendees.setAsync(filtered, (setResult: any) => {
                if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("[AB Book IQ] Removed room from required attendees:", emailToRemove);
                }
                res();
              });
            } else {
              res();
            }
          } else {
            res();
          }
        });
      });

      resolve();
    } catch (err) {
      reject(err);
    }
  });
}

/**
 * Remove ALL rooms from the meeting.
 * Since rooms are now added BOTH as required attendees AND as the location on
 * every Outlook version, removal must clean up both places.
 *
 * Modern Outlook: remove room locations via enhancedLocation AND filter the
 *   room out of requiredAttendees.
 * Old Outlook: clear resources, filter requiredAttendees, and clear the basic
 *   location text field.
 */
export async function removeAllRooms(allRoomEmails: string[]): Promise<void> {
  return new Promise(async (resolve, reject) => {
    const item = getMailboxItem();
    if (!item) {
      reject(new Error("No appointment item available"));
      return;
    }

    try {
      const roomEmailsLower = new Set(allRoomEmails.map(e => e.toLowerCase()));
      const organizerEmail = getOrganizerEmail().toLowerCase();

      // Check which APIs are available
      const hasEnhancedLocation = typeof (item as any).enhancedLocation?.getAsync === 'function';
      const hasResources = typeof (item as any).resources?.setAsync === 'function';

      console.log("[AB Book IQ] removeAllRooms - hasEnhancedLocation:", hasEnhancedLocation, "hasResources:", hasResources);

      // Step 1 (modern Outlook only): remove room locations via enhancedLocation
      if (hasEnhancedLocation) {
        await new Promise<void>((res) => {
          (item as any).enhancedLocation.getAsync((result: any) => {
            if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.length > 0) {
              const locations = result.value;
              console.log("[AB Book IQ] Current enhancedLocations:", locations.map((l: any) => ({ id: l.locationIdentifier?.id, type: l.locationIdentifier?.type })));

              // Find all room locations to remove
              const roomLocationsToRemove = locations
                .filter((loc: any) => {
                  const locId = loc.locationIdentifier?.id?.toLowerCase() || '';
                  const locType = loc.locationIdentifier?.type;
                  return locType === Office.MailboxEnums.LocationType.Room || roomEmailsLower.has(locId);
                })
                .map((loc: any) => loc.locationIdentifier);

              if (roomLocationsToRemove.length > 0) {
                console.log("[AB Book IQ] Removing room locations:", roomLocationsToRemove.length);
                (item as any).enhancedLocation.removeAsync(roomLocationsToRemove, (removeResult: any) => {
                  if (removeResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("[AB Book IQ] Removed rooms via enhancedLocation");
                  } else {
                    console.log("[AB Book IQ] enhancedLocation.removeAsync failed:", removeResult.error?.message);
                  }
                  res();
                });
              } else {
                res();
              }
            } else {
              res();
            }
          });
        });
      }

      // Step 2 (old Outlook only): clear the resources collection
      if (!hasEnhancedLocation && hasResources) {
        await new Promise<void>((res) => {
          (item as any).resources.getAsync((result: any) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const currentResources = result.value || [];
              console.log("[AB Book IQ] Current resources:", currentResources.length);
              if (currentResources.length > 0) {
                (item as any).resources.setAsync([], (setResult: any) => {
                  if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("[AB Book IQ] Cleared all resources");
                  } else {
                    console.log("[AB Book IQ] Failed to clear resources:", setResult.error?.message);
                  }
                  res();
                });
              } else {
                res();
              }
            } else {
              res();
            }
          });
        });
      }

      // Step 3 (all versions): filter rooms out of required attendees
      await new Promise<void>((res) => {
        item.requiredAttendees.getAsync((result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const currentAttendees = result.value || [];
            console.log("[AB Book IQ] Current required attendees:", currentAttendees.map((a: any) => a.emailAddress));

            const filtered = currentAttendees
              .filter((a: any) => {
                const email = a.emailAddress.toLowerCase();
                if (roomEmailsLower.has(email)) return false;
                if (email === organizerEmail) return false;
                return true;
              })
              .map((a: any) => ({ displayName: a.displayName, emailAddress: a.emailAddress }));

            console.log("[AB Book IQ] After filtering, keeping:", filtered.length, "attendees");

            item.requiredAttendees.setAsync(filtered, (setResult: any) => {
              if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("[AB Book IQ] Updated required attendees");
              } else {
                console.log("[AB Book IQ] Failed to update required attendees:", setResult.error?.message);
              }
              res();
            });
          } else {
            res();
          }
        });
      });

      // Step 4 (all versions): clear the basic location text field.
      // We always set the location as text now, so it must always be cleared
      // here regardless of whether enhancedLocation is available.
      await setLocation("");
      console.log("[AB Book IQ] Cleared location");

      resolve();
    } catch (err) {
      console.error("[AB Book IQ] removeAllRooms error:", err);
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
