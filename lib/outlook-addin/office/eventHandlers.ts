type EventCallback = () => void;

let appointmentChangedCallbacks: EventCallback[] = [];
let isListening = false;
let officeLoaded = false;

/**
 * Load Office.js from CDN dynamically
 */
function loadOfficeJs(): Promise<void> {
  return new Promise((resolve, reject) => {
    // Already loaded
    if (typeof Office !== "undefined") {
      resolve();
      return;
    }

    // Check if script already exists
    if (document.querySelector('script[src*="office.js"]')) {
      // Wait for it to load
      const checkInterval = setInterval(() => {
        if (typeof Office !== "undefined") {
          clearInterval(checkInterval);
          resolve();
        }
      }, 100);
      return;
    }

    // Load the script
    const script = document.createElement("script");
    script.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("Failed to load Office.js"));
    document.head.appendChild(script);
  });
}

/**
 * Initialize Office.js and wait for it to be ready
 */
export async function initializeOffice(): Promise<void> {
  if (officeLoaded) return;

  // Try to load Office.js from CDN
  try {
    await loadOfficeJs();
  } catch {
    console.log("[AB Book IQ] Office.js not available - running in preview mode");
    officeLoaded = true;
    return;
  }

  return new Promise((resolve, reject) => {
    if (typeof Office === "undefined") {
      // Office.js not loaded - we're in development/preview mode
      console.log("[AB Book IQ] Office.js not available - running in preview mode");
      officeLoaded = true;
      resolve();
      return;
    }

    Office.onReady((info) => {
      officeLoaded = true;
      if (info.host === Office.HostType.Outlook) {
        console.log("[AB Book IQ] Office.js ready in Outlook");
        resolve();
      } else if (!info.host) {
        // Running outside of Office - preview mode
        console.log("[AB Book IQ] Running in preview mode (no Office host)");
        resolve();
      } else {
        reject(new Error(`Unsupported host: ${info.host}`));
      }
    });
  });
}

/**
 * Check if we're running inside Outlook
 */
export function isInOutlook(): boolean {
  return typeof Office !== "undefined" &&
    Office.context !== undefined &&
    Office.context.mailbox !== undefined;
}

/**
 * Check if we're in appointment compose mode
 */
export function isAppointmentComposeMode(): boolean {
  if (!isInOutlook()) return false;

  const item = Office.context.mailbox.item;
  if (!item) return false;

  // Check if we're in appointment compose mode
  // itemType may not be available on all item types, so we also check for start
  const itemAsAny = item as { itemType?: Office.MailboxEnums.ItemType; start?: unknown };
  return (
    itemAsAny.itemType === Office.MailboxEnums.ItemType.Appointment ||
    itemAsAny.start !== undefined
  );
}

/**
 * Subscribe to appointment changes (time changes, etc.)
 */
export function onAppointmentChanged(callback: EventCallback): () => void {
  appointmentChangedCallbacks.push(callback);

  if (!isListening && isInOutlook()) {
    startListeningForChanges();
  }

  return () => {
    appointmentChangedCallbacks = appointmentChangedCallbacks.filter(
      (cb) => cb !== callback
    );
  };
}

/**
 * Start listening for appointment changes
 */
function startListeningForChanges(): void {
  if (isListening) return;
  isListening = true;

  const item = Office.context.mailbox.item;
  if (!item) return;

  // Listen for recurrence changes if supported
  if (item.addHandlerAsync) {
    try {
      item.addHandlerAsync(
        Office.EventType.AppointmentTimeChanged,
        handleAppointmentChanged
      );
    } catch {
      console.log("[AB Book IQ] AppointmentTimeChanged event not supported");
    }
  }
}

/**
 * Handle appointment changed event
 */
function handleAppointmentChanged(): void {
  for (const callback of appointmentChangedCallbacks) {
    try {
      callback();
    } catch (error) {
      console.error("[AB Book IQ] Error in appointment change callback:", error);
    }
  }
}

/**
 * Show a notification in Outlook
 */
export function showNotification(
  message: string,
  type: "informational" | "error" = "informational"
): void {
  if (!isInOutlook()) {
    console.log(`[EA BookIQ Notification] ${type}: ${message}`);
    return;
  }

  const item = Office.context.mailbox.item;
  if (item?.notificationMessages) {
    const key = `room-assist-${Date.now()}`;
    item.notificationMessages.addAsync(key, {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message,
      icon: "Icon.80x80",
      persistent: false,
    });

    // Auto-remove after 5 seconds
    setTimeout(() => {
      item.notificationMessages.removeAsync(key);
    }, 5000);
  }
}
