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
 * Detect classic Outlook desktop (Windows "Outlook" / "Mac").
 *
 * This matters because classic desktop CROSS-POPULATES the attendee and
 * location fields: adding a room as a required attendee also sets it as the
 * location, and adding it via enhancedLocation also adds it as an attendee.
 * New Outlook and Outlook on the web keep the two fields independent.
 *
 * We use this to give identical end-results on every client.
 */
export function isClassicOutlookDesktop(): boolean {
  try {
    const host = (Office.context?.mailbox?.diagnostics as any)?.hostName;
    // Classic Windows desktop reports "Outlook"; classic Mac reports "Mac".
    // New Outlook for Windows reports "newOutlookWindows"; web reports
    // "OutlookWebApp" — both keep attendee/location independent.
    return host === "Outlook" || host === "Mac";
  } catch {
    return false;
  }
}

/**
 * Check if enhancedLocation API is available (Mailbox requirement set 1.8+)
 */
function hasEnhancedLocation(item: Office.AppointmentCompose): boolean {
  return typeof (item as any).enhancedLocation?.addAsync === "function";
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

    // Safety guard: on new Outlook, resources.getAsync sometimes never invokes
    // its callback, which would leave pendingCount > 0 and make this promise
    // hang forever (blocking addRoomAttendee / bookRoom). Always settle: resolve
    // with whatever we've collected once all callbacks fire OR after a timeout.
    let settled = false;
    const settle = () => {
      if (settled) return;
      settled = true;
      console.log("[AB Book IQ] Attendees fetched:", attendees.length, "total");
      resolve(attendees);
    };
    setTimeout(settle, 2500);

    const checkComplete = () => {
      pendingCount--;
      if (pendingCount === 0) {
        settle();
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
 * Add a single recipient to requiredAttendees. Resolves true on success and
 * false on failure (never rejects). Per the Office.js docs, the recipients
 * array may contain either EmailUser objects or plain SMTP address strings.
 */
function tryAddRequiredAttendee(
  item: Office.AppointmentCompose,
  recipient: string | { displayName: string; emailAddress: string }
): Promise<boolean> {
  return new Promise((resolve) => {
    try {
      item.requiredAttendees.addAsync([recipient] as any, (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(true);
        } else {
          console.log("[AB Book IQ] requiredAttendees.addAsync failed:", result.error?.message);
          resolve(false);
        }
      });
    } catch (err) {
      console.log("[AB Book IQ] requiredAttendees.addAsync threw:", err);
      resolve(false);
    }
  });
}

/**
 * Add a room as a required attendee (idempotent — won't add a duplicate).
 *
 * All Outlook clients use requiredAttendees.addAsync — the documented, supported
 * way to add an attendee (including a room mailbox) to an appointment in classic
 * Outlook, new Outlook, and Outlook on the web. Per the docs the recipient can
 * be an EmailUser object OR a plain SMTP string, so we try the object form first
 * and, if the server rejects it, retry with the bare SMTP string (some Exchange
 * setups only resolve a room mailbox from the raw address).
 *
 * On new Outlook / OWA we do NOT reject on failure and we do NOT fall back to
 * enhancedLocation here — routing the room into the location field is what
 * caused it to appear as a location instead of an attendee. Booking the room as
 * a location is handled separately by addRoomLocation (only in "both"/"location"
 * modes), so attendee-only bookings never populate the location field.
 */
export async function addRoomAttendee(
  displayName: string,
  emailAddress: string
): Promise<void> {
  const item = getMailboxItem();
  if (!item) {
    throw new Error("No appointment item available");
  }

  // Idempotency guard: skip if the room is already an attendee/resource.
  const alreadyAdded = await isRoomAlreadyAdded(emailAddress);
  if (alreadyAdded) {
    console.log("[AB Book IQ] Room already added, skipping:", displayName);
    return;
  }

  // Classic desktop: requiredAttendees works reliably (unchanged behavior —
  // reject on failure so callers surface the error).
  if (isClassicOutlookDesktop()) {
    return new Promise((resolve, reject) => {
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

  // New Outlook / OWA: try the object form, then retry with a plain SMTP string.
  if (await tryAddRequiredAttendee(item, { displayName, emailAddress })) {
    console.log("[AB Book IQ] Added room to requiredAttendees:", displayName);
    return;
  }
  console.log("[AB Book IQ] Retrying attendee add with SMTP string:", emailAddress);
  if (await tryAddRequiredAttendee(item, emailAddress)) {
    console.log("[AB Book IQ] Added room to requiredAttendees (string form):", displayName);
    return;
  }

  // Do not throw and do not fall back to location — let addRoomLocation handle
  // location for "both"/"location" modes so attendee-only never lands in the
  // location field.
  console.log("[AB Book IQ] Could not add room to requiredAttendees:", displayName);
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
 * Add a room as the meeting LOCATION (the actual room resource, not plain text).
 *
 * - New Outlook / web: uses enhancedLocation.addAsync with type Room so the
 *   room resolves to a real resource (avoids the plain-text "Unknown" label)
 *   and supports multiple location entries.
 * - Classic desktop: enhancedLocation also adds the room as an attendee, which
 *   would break "location only". So on classic we set the plain text location
 *   instead. (Classic resolves/displays room text natively without "Unknown".)
 *
 * This function never adds the room as an attendee.
 */
export async function addRoomLocation(
  displayName: string,
  emailAddress: string
): Promise<void> {
  const item = getMailboxItem();
  if (!item) {
    throw new Error("No appointment item available");
  }

  const classic = isClassicOutlookDesktop();

  if (!classic && hasEnhancedLocation(item)) {
    await new Promise<void>((resolve) => {
      const locationIdentifier = {
        id: emailAddress,
        type: Office.MailboxEnums.LocationType.Room,
      };
      (item as any).enhancedLocation.addAsync([locationIdentifier], (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("[AB Book IQ] Added room as location resource:", displayName);
        } else {
          console.log("[AB Book IQ] enhancedLocation.addAsync failed:", result.error?.message);
        }
        resolve();
      });
    });
  } else {
    // Classic desktop (or no enhancedLocation): set plain text location.
    await setLocation(displayName);
    console.log("[AB Book IQ] Set room as location (text):", displayName);
  }
}

/**
 * Remove a single room from the LOCATION only (keeps it as an attendee).
 *
 * Removes the matching enhancedLocation entry (by room email) and clears the
 * plain text location if it matches the room name. Used to undo classic
 * Outlook's auto-population of the location when adding an attendee, and when
 * unbooking a room.
 */
export async function removeRoomLocation(
  displayName: string,
  emailAddress: string
): Promise<void> {
  const item = getMailboxItem();
  if (!item) {
    throw new Error("No appointment item available");
  }

  const normalizedEmail = emailAddress.toLowerCase();

  // Remove the matching enhancedLocation entry (specific room only).
  if (hasEnhancedLocation(item) && typeof (item as any).enhancedLocation?.getAsync === "function") {
    await new Promise<void>((resolve) => {
      (item as any).enhancedLocation.getAsync((result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value?.length > 0) {
          const toRemove = result.value
            .filter((loc: any) => (loc.locationIdentifier?.id?.toLowerCase() || "") === normalizedEmail)
            .map((loc: any) => loc.locationIdentifier);

          if (toRemove.length > 0) {
            (item as any).enhancedLocation.removeAsync(toRemove, () => resolve());
            return;
          }
        }
        resolve();
      });
    });
  }

  // Clear the plain text location if it matches this room.
  const currentLocation = await getLocation();
  if (currentLocation && currentLocation.toLowerCase().includes(displayName.toLowerCase())) {
    await setLocation("");
    console.log("[AB Book IQ] Cleared text location for room:", displayName);
  }
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
 * Determine, for the given rooms, which are present as ATTENDEES and which are
 * present as LOCATIONS. A room is considered "fully booked" only when it is in
 * BOTH sets.
 *
 * Location detection checks enhancedLocation entries (by room email) and the
 * plain text location (by room name), so it works on every Outlook version.
 */
export async function getRoomPresence(
  allRoomEmails: string[],
  allRoomNames: string[]
): Promise<{ attendees: Set<string>; locations: Set<string> }> {
  const attendees = new Set<string>();
  const locations = new Set<string>();

  const item = getMailboxItem();
  if (!item) {
    return { attendees, locations };
  }

  // Attendees (required + optional + resources). Match by email OR display
  // name: on classic Outlook desktop a room is a resolved resource whose
  // attendee email returned by the API doesn't always match the Graph room
  // email, so we also match on the room's display name as a reliable fallback.
  // (New Outlook / OWA already match by email, so this only adds matches.)
  const currentAttendees = await getCurrentAttendees();
  const attendeeEmails = new Set(currentAttendees.map((a) => a.emailAddress.toLowerCase()));
  const attendeeNames = new Set(currentAttendees.map((a) => (a.displayName || "").toLowerCase()));
  for (let i = 0; i < allRoomEmails.length; i++) {
    const email = allRoomEmails[i].toLowerCase();
    const name = (allRoomNames[i] || "").toLowerCase();
    if (attendeeEmails.has(email) || (name && attendeeNames.has(name))) {
      attendees.add(email);
    }
  }

  // Locations via enhancedLocation (match by room email)
  if (typeof (item as any).enhancedLocation?.getAsync === "function") {
    await new Promise<void>((resolve) => {
      (item as any).enhancedLocation.getAsync((result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value?.length > 0) {
          const locIds = result.value.map((l: any) => (l.locationIdentifier?.id || "").toLowerCase());
          for (const email of allRoomEmails) {
            if (locIds.includes(email.toLowerCase())) {
              locations.add(email.toLowerCase());
            }
          }
        }
        resolve();
      });
    });
  }

  // Locations via plain text location (match by room name)
  const locationText = await getLocation();
  if (locationText) {
    const lowerLocation = locationText.toLowerCase();
    for (let i = 0; i < allRoomNames.length; i++) {
      const name = allRoomNames[i];
      const email = allRoomEmails[i];
      if (name && lowerLocation.includes(name.toLowerCase())) {
        locations.add(email.toLowerCase());
      }
    }
  }

  return { attendees, locations };
}

/**
 * Per-item persistence of "fully booked" rooms (added via the Book button as
 * BOTH attendee + location). Stored in the appointment's custom properties so
 * the Booked state survives refresh / reopening the add-in.
 *
 * This is required for classic Outlook desktop, whose Office.js attendee reads
 * don't reliably return a freshly-added (unsaved) room resource — so we can't
 * rely solely on getRoomPresence to redetect Booked rooms there. New Outlook /
 * OWA detect presence reliably, so for them this is just a redundant safety net
 * and their behavior is unchanged.
 */
const BOOKED_ROOMS_PROP_KEY = "abBookedRoomEmails";

// Custom-property calls can occasionally hang or throw in some Outlook clients
// (notably on unsaved compose items in new Outlook). This safety timeout
// guarantees the persistence helpers ALWAYS settle so they can never block or
// break the booking flow.
const CUSTOM_PROPS_TIMEOUT_MS = 3000;

export async function getPersistedBookedRooms(): Promise<Set<string>> {
  return new Promise((resolve) => {
    let settled = false;
    const done = (value: Set<string>) => {
      if (settled) return;
      settled = true;
      resolve(value);
    };

    // Never hang: resolve with an empty set if the Office callback never fires.
    setTimeout(() => done(new Set()), CUSTOM_PROPS_TIMEOUT_MS);

    try {
      const item = getMailboxItem();
      if (!item || typeof (item as any).loadCustomPropertiesAsync !== "function") {
        done(new Set());
        return;
      }
      (item as any).loadCustomPropertiesAsync((result: any) => {
        try {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
            const raw = result.value.get(BOOKED_ROOMS_PROP_KEY);
            const emails = raw
              ? String(raw).split(",").map((e) => e.trim().toLowerCase()).filter(Boolean)
              : [];
            done(new Set(emails));
          } else {
            done(new Set());
          }
        } catch {
          done(new Set());
        }
      });
    } catch (err) {
      console.log("[AB Book IQ] getPersistedBookedRooms error (non-critical):", err);
      done(new Set());
    }
  });
}

export async function markRoomBooked(emailAddress: string, booked: boolean): Promise<void> {
  return new Promise((resolve) => {
    let settled = false;
    const done = () => {
      if (settled) return;
      settled = true;
      resolve();
    };

    // Persistence is a best-effort safety net; never let it block or fail the
    // booking flow. Always resolve (even on timeout/error).
    setTimeout(done, CUSTOM_PROPS_TIMEOUT_MS);

    try {
      const item = getMailboxItem();
      if (!item || typeof (item as any).loadCustomPropertiesAsync !== "function") {
        done();
        return;
      }
      (item as any).loadCustomPropertiesAsync((result: any) => {
        try {
          if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
            done();
            return;
          }
          const props = result.value;
          const raw = props.get(BOOKED_ROOMS_PROP_KEY);
          const set = new Set(
            raw ? String(raw).split(",").map((e) => e.trim().toLowerCase()).filter(Boolean) : []
          );
          const normalized = emailAddress.toLowerCase();
          if (booked) {
            set.add(normalized);
          } else {
            set.delete(normalized);
          }
          props.set(BOOKED_ROOMS_PROP_KEY, Array.from(set).join(","));
          props.saveAsync(() => done());
        } catch (err) {
          console.log("[AB Book IQ] markRoomBooked save error (non-critical):", err);
          done();
        }
      });
    } catch (err) {
      console.log("[AB Book IQ] markRoomBooked error (non-critical):", err);
      done();
    }
  });
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
