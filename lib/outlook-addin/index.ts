// Room Assist - Outlook Add-in
// Main exports for the add-in library

// Auth
export {
  initializeMsal,
  acquireGraphToken,
  signIn,
  signOut,
  isSignedIn,
  getAccount,
} from "./auth/msal";

// Config
export { OFFICE_CONFIGS, MSAL_CONFIG, GRAPH_SCOPES, STORAGE_KEYS } from "./config/offices";
export type { OfficeConfig } from "./config/offices";

// Graph Services
export { getGraphClient, resetGraphClient } from "./graph/graphClient";
export { getTransitiveGroupMemberships, resolveUserOffices, isUserInOffice } from "./graph/groups";
export { getRoomLists, getRoomsForOffice, getRoomsByBuilding } from "./graph/places";
export type { Room, RoomList } from "./graph/places";
export { getSchedule, checkRoomAvailability, filterAvailableRooms, sortRoomsByPreference } from "./graph/schedule";
export type { ScheduleItem, ScheduleInfo, RoomAvailability, TimeSlot } from "./graph/schedule";

// Office.js Services
export {
  getMeetingWindow,
  getCurrentAttendees,
  addRoomAttendee,
  setLocation,
  getLocation,
  isRoomAlreadyAdded,
  getOrganizerEmail,
  getOrganizerDisplayName,
} from "./office/appointment";
export type { MeetingWindow, Attendee } from "./office/appointment";
export {
  initializeOffice,
  isInOutlook,
  isAppointmentComposeMode,
  onAppointmentChanged,
  showNotification,
} from "./office/eventHandlers";

// Domain Logic
export {
  resolveOffice,
  getCachedOfficePreference,
  setCachedOfficePreference,
  clearCachedOfficePreference,
  getOfficeById,
  getAllOffices,
} from "./domain/officeResolver";
export type { OfficeResolutionResult } from "./domain/officeResolver";
export { bookRoom, removeRoom } from "./domain/booking";
export type { BookingResult } from "./domain/booking";
export { rankRooms, groupRoomsByAvailability, getRoomFeatures } from "./domain/roomRanker";
export type { RankingOptions } from "./domain/roomRanker";
