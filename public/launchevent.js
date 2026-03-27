/**
 * Room Assist - Launch Event Handler
 * 
 * This file handles event-based activation for the Outlook add-in.
 * It runs when a new appointment is created and can pre-initialize
 * the add-in experience.
 * 
 * Note: Event-based handlers must be lightweight and call event.completed()
 * promptly. Heavy operations should be deferred to the task pane.
 */

/* global Office */

Office.onReady(function() {
  // Register the event handler
  Office.actions.associate("onNewAppointmentOrganizer", onNewAppointmentOrganizer);
});

/**
 * Handler for new appointment organizer event
 * Called automatically when user creates a new meeting
 */
function onNewAppointmentOrganizer(event) {
  try {
    // Show a notification that Room Assist is available
    Office.context.mailbox.item.notificationMessages.addAsync(
      "room-assist-available",
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Room Assist is ready. Click 'Book Room' to find available conference rooms.",
        icon: "Icon.80x80",
        persistent: false
      },
      function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log("[Room Assist] Failed to show notification:", asyncResult.error.message);
        }
      }
    );

    // Auto-remove notification after 8 seconds
    setTimeout(function() {
      Office.context.mailbox.item.notificationMessages.removeAsync("room-assist-available");
    }, 8000);

  } catch (error) {
    console.error("[Room Assist] Launch event error:", error);
  }

  // Must call completed to signal the event handler is done
  event.completed();
}
