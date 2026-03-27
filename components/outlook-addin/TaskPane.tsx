"use client";

import { useState, useEffect, useCallback } from "react";
import { RefreshCw, LogIn, Building2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Spinner } from "@/components/ui/spinner";
import { ScrollArea } from "@/components/ui/scroll-area";
import { RoomCard } from "./RoomCard";
import { EmptyState } from "./EmptyState";
import { ErrorState } from "./ErrorState";
import { OfficeSelector, OfficeToggle } from "./OfficeSelector";

import { initializeMsal, signIn, isSignedIn, getAccount } from "@/lib/outlook-addin/auth/msal";
import { initializeOffice, isInOutlook, onAppointmentChanged } from "@/lib/outlook-addin/office/eventHandlers";
import { getMeetingWindow, type MeetingWindow } from "@/lib/outlook-addin/office/appointment";
import { resolveOffice, setCachedOfficePreference, getAllOffices, type OfficeResolutionResult } from "@/lib/outlook-addin/domain/officeResolver";
import { getRoomsForOffice, type Room } from "@/lib/outlook-addin/graph/places";
import { checkRoomAvailability, type RoomAvailability } from "@/lib/outlook-addin/graph/schedule";
import { rankRooms } from "@/lib/outlook-addin/domain/roomRanker";
import { bookRoom } from "@/lib/outlook-addin/domain/booking";
import { isMsalConfigured, type OfficeConfig } from "@/lib/outlook-addin/config/offices";

type AppState = "initializing" | "sign-in" | "select-office" | "loading" | "ready" | "error" | "not-configured";

export function TaskPane() {
  const [appState, setAppState] = useState<AppState>("initializing");
  const [error, setError] = useState<string | null>(null);
  const [meetingWindow, setMeetingWindow] = useState<MeetingWindow | null>(null);
  const [officeResult, setOfficeResult] = useState<OfficeResolutionResult | null>(null);
  const [selectedOffice, setSelectedOffice] = useState<OfficeConfig | null>(null);
  const [rooms, setRooms] = useState<RoomAvailability[]>([]);
  const [bookingRoomId, setBookingRoomId] = useState<string | null>(null);
  const [bookedRoomIds, setBookedRoomIds] = useState<Set<string>>(new Set());
  const [isRefreshing, setIsRefreshing] = useState(false);

  // Initialize the add-in
  useEffect(() => {
    async function initialize() {
      try {
        // Check if Azure AD is configured
        if (!isMsalConfigured()) {
          setAppState("not-configured");
          return;
        }

        // Initialize Office.js
        await initializeOffice();

        // Initialize MSAL
        await initializeMsal();

        // Check if signed in
        if (!isSignedIn()) {
          setAppState("sign-in");
          return;
        }

        // Proceed to load office and rooms
        await loadOfficeAndRooms();
      } catch (err) {
        console.error("[EA BookIQ] Initialization error:", err);
        setError(err instanceof Error ? err.message : "Failed to initialize");
        setAppState("error");
      }
    }

    initialize();
  }, []);

  // Subscribe to appointment changes
  useEffect(() => {
    if (!isInOutlook()) return;

    const unsubscribe = onAppointmentChanged(() => {
      // Refresh rooms when appointment time changes
      if (selectedOffice && appState === "ready") {
        loadRooms(selectedOffice);
      }
    });

    return unsubscribe;
  }, [selectedOffice, appState]);

  const loadOfficeAndRooms = useCallback(async () => {
    setAppState("loading");
    setError(null);

    try {
      // Get meeting window
      const window = isInOutlook() ? await getMeetingWindow() : getMockMeetingWindow();
      setMeetingWindow(window);

      if (!window.complete) {
        setAppState("ready");
        return;
      }

      // Resolve office
      const result = await resolveOffice();
      setOfficeResult(result);

      if (result.type === "none") {
        setAppState("select-office");
        return;
      }

      // Determine which office to use
      let office: OfficeConfig;
      if (result.type === "single") {
        office = result.office;
      } else if (result.cached) {
        office = result.cached;
      } else {
        setAppState("select-office");
        return;
      }

      setSelectedOffice(office);
      await loadRooms(office);
    } catch (err) {
      console.error("[EA BookIQ] Load error:", err);
      setError(err instanceof Error ? err.message : "Failed to load rooms");
      setAppState("error");
    }
  }, []);

  const loadRooms = async (office: OfficeConfig) => {
    if (!meetingWindow?.complete) {
      setRooms([]);
      setAppState("ready");
      return;
    }

    setIsRefreshing(true);
    try {
      // Get rooms for the office
      const officeRooms = await getRoomsForOffice(office);

      // Check availability
      const availability = await checkRoomAvailability(
        officeRooms,
        meetingWindow.start!,
        meetingWindow.end!,
        meetingWindow.timeZone
      );

      // Rank rooms
      const rankedRooms = rankRooms(availability);
      setRooms(rankedRooms);
      setAppState("ready");
    } catch (err) {
      console.error("[EA BookIQ] Room load error:", err);
      setError(err instanceof Error ? err.message : "Failed to load rooms");
      setAppState("error");
    } finally {
      setIsRefreshing(false);
    }
  };

  const handleSignIn = async () => {
    try {
      setAppState("loading");
      await signIn();
      await loadOfficeAndRooms();
    } catch (err) {
      console.error("[EA BookIQ] Sign in error:", err);
      setError(err instanceof Error ? err.message : "Sign in failed");
      setAppState("sign-in");
    }
  };

  const handleOfficeSelect = async (office: OfficeConfig) => {
    setSelectedOffice(office);
    setCachedOfficePreference(office.id);
    setAppState("loading");

    // Need to get meeting window first if we don't have it
    if (!meetingWindow) {
      const window = isInOutlook() ? await getMeetingWindow() : getMockMeetingWindow();
      setMeetingWindow(window);
    }

    await loadRooms(office);
  };

  const handleBookRoom = async (room: Room) => {
    setBookingRoomId(room.id);
    try {
      const result = await bookRoom(room);
      if (result.success) {
        setBookedRoomIds((prev) => new Set(prev).add(room.id));
      } else {
        setError(result.message);
      }
    } catch (err) {
      console.error("[EA BookIQ] Booking error:", err);
      setError(err instanceof Error ? err.message : "Failed to book room");
    } finally {
      setBookingRoomId(null);
    }
  };

  const handleRefresh = () => {
    if (selectedOffice) {
      loadRooms(selectedOffice);
    } else {
      loadOfficeAndRooms();
    }
  };

  // Render based on state
  if (appState === "not-configured") {
    return (
      <div className="flex flex-col items-center justify-center h-full p-8">
        <Building2 className="size-12 text-muted-foreground" />
        <h3 className="mt-4 font-semibold">Configuration Required</h3>
        <p className="mt-2 text-sm text-muted-foreground text-center max-w-[280px]">
          Azure AD credentials are not configured. Please set the following environment variables:
        </p>
        <div className="mt-4 p-3 bg-muted rounded-lg text-xs font-mono space-y-1">
          <p>NEXT_PUBLIC_AZURE_CLIENT_ID</p>
          <p>NEXT_PUBLIC_AZURE_TENANT_ID</p>
        </div>
        <p className="mt-3 text-xs text-muted-foreground text-center max-w-[280px]">
          Add these in the Vars section of the settings menu (gear icon, top right).
        </p>
      </div>
    );
  }

  if (appState === "initializing" || appState === "loading") {
    return (
      <div className="flex flex-col items-center justify-center h-full p-8">
        <Spinner className="size-8" />
        <p className="mt-4 text-sm text-muted-foreground">
          {appState === "initializing" ? "Initializing..." : "Loading rooms..."}
        </p>
      </div>
    );
  }

  if (appState === "sign-in") {
    return (
      <div className="flex flex-col items-center justify-center h-full p-8">
        <LogIn className="size-12 text-muted-foreground" />
        <h3 className="mt-4 font-semibold">Sign In Required</h3>
        <p className="mt-2 text-sm text-muted-foreground text-center max-w-[250px]">
          Please sign in to access conference room booking.
        </p>
        <Button onClick={handleSignIn} className="mt-4">
          Sign In with Microsoft
        </Button>
        {error && (
          <p className="mt-2 text-xs text-destructive">{error}</p>
        )}
      </div>
    );
  }

  if (appState === "select-office") {
    const offices = officeResult?.type === "multiple" 
      ? officeResult.offices 
      : getAllOffices();
    
    return (
      <div className="p-4">
        <OfficeSelector
          offices={offices}
          selectedOffice={selectedOffice || undefined}
          onSelect={handleOfficeSelect}
        />
      </div>
    );
  }

  if (appState === "error") {
    return (
      <div className="p-4">
        <ErrorState
          message={error || "An unexpected error occurred"}
          onRetry={handleRefresh}
        />
      </div>
    );
  }

  // Ready state
  const availableRooms = rooms.filter((r) => r.isAvailable);
  const unavailableRooms = rooms.filter((r) => !r.isAvailable);

  return (
    <div className="flex flex-col h-full">
      {/* Header */}
      <div className="p-4 border-b space-y-3">
        <div className="flex items-center justify-between">
          <div>
            <h2 className="font-semibold text-foreground">EA BookIQ</h2>
            {getAccount() && (
              <p className="text-xs text-muted-foreground">
                {getAccount()?.username}
              </p>
            )}
          </div>
          <Button
            variant="ghost"
            size="icon-sm"
            onClick={handleRefresh}
            disabled={isRefreshing}
          >
            <RefreshCw className={isRefreshing ? "animate-spin" : ""} />
          </Button>
        </div>

        {/* Office toggle (if multiple offices) */}
        {officeResult?.type === "multiple" && selectedOffice && (
          <OfficeToggle
            offices={officeResult.offices}
            selectedOffice={selectedOffice}
            onSelect={handleOfficeSelect}
          />
        )}

        {/* Office indicator */}
        {selectedOffice && officeResult?.type !== "multiple" && (
          <div className="flex items-center gap-2 text-sm text-muted-foreground">
            <Building2 className="size-4" />
            <span>{selectedOffice.displayName}</span>
          </div>
        )}

        {/* Meeting time indicator */}
        {meetingWindow?.complete && (
          <div className="text-xs text-muted-foreground">
            {formatTimeRange(meetingWindow.start!, meetingWindow.end!)}
          </div>
        )}
      </div>

      {/* Content */}
      <ScrollArea className="flex-1">
        <div className="p-4 space-y-4">
          {!meetingWindow?.complete ? (
            <EmptyState type="no-time" />
          ) : rooms.length === 0 ? (
            <EmptyState
              type="no-rooms"
              onAction={handleRefresh}
              actionLabel="Refresh"
            />
          ) : (
            <>
              {/* Available rooms */}
              {availableRooms.length > 0 && (
                <div>
                  <div className="flex items-center gap-2 mb-2">
                    <h3 className="text-sm font-medium">Available</h3>
                    <Badge variant="secondary" className="text-xs">
                      {availableRooms.length}
                    </Badge>
                  </div>
                  <div className="space-y-2">
                    {availableRooms.map((roomAvail) => (
                      <RoomCard
                        key={roomAvail.room.id}
                        roomAvailability={roomAvail}
                        onBook={handleBookRoom}
                        isBooking={bookingRoomId === roomAvail.room.id}
                        isBooked={bookedRoomIds.has(roomAvail.room.id)}
                      />
                    ))}
                  </div>
                </div>
              )}

              {/* Unavailable rooms */}
              {unavailableRooms.length > 0 && (
                <div>
                  <div className="flex items-center gap-2 mb-2">
                    <h3 className="text-sm font-medium text-muted-foreground">
                      Unavailable
                    </h3>
                    <Badge variant="outline" className="text-xs">
                      {unavailableRooms.length}
                    </Badge>
                  </div>
                  <div className="space-y-2">
                    {unavailableRooms.map((roomAvail) => (
                      <RoomCard
                        key={roomAvail.room.id}
                        roomAvailability={roomAvail}
                        onBook={handleBookRoom}
                        isBooking={bookingRoomId === roomAvail.room.id}
                        isBooked={bookedRoomIds.has(roomAvail.room.id)}
                      />
                    ))}
                  </div>
                </div>
              )}
            </>
          )}
        </div>
      </ScrollArea>
    </div>
  );
}

// Helper function to format time range
function formatTimeRange(start: Date, end: Date): string {
  const options: Intl.DateTimeFormatOptions = {
    weekday: "short",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  };

  const startStr = start.toLocaleString(undefined, options);
  const endStr = end.toLocaleTimeString(undefined, {
    hour: "numeric",
    minute: "2-digit",
  });

  return `${startStr} - ${endStr}`;
}

// Mock meeting window for development/preview mode
function getMockMeetingWindow(): MeetingWindow {
  const start = new Date();
  start.setMinutes(0, 0, 0);
  start.setHours(start.getHours() + 1);

  const end = new Date(start);
  end.setHours(end.getHours() + 1);

  return {
    start,
    end,
    complete: true,
    timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
  };
}
