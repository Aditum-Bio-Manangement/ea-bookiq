"use client";

import { Users, MapPin, Check, ChevronDown } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Spinner } from "@/components/ui/spinner";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { cn } from "@/lib/utils";
import type { RoomAvailability } from "@/lib/outlook-addin/graph/schedule";
import type { BookingMode } from "@/lib/outlook-addin/domain/booking";

interface RoomCardProps {
  roomAvailability: RoomAvailability;
  onBook: (room: RoomAvailability["room"], mode: BookingMode) => void;
  isBooking?: boolean;
  isBooked?: boolean;
}

export function RoomCard({
  roomAvailability,
  onBook,
  isBooking = false,
  isBooked = false,
}: RoomCardProps) {
  const { room, isAvailable } = roomAvailability;

  return (
    <Card
      className={cn(
        "transition-all w-full",
        !isAvailable && "opacity-60",
        isBooked && "border-primary bg-primary/5"
      )}
    >
      <CardContent className="p-3">
        <div className="flex items-center justify-between gap-2 w-full">
          {/* Room info - takes available space, truncates if needed */}
          <div className="min-w-0 flex-1">
            <h3 className="font-medium text-sm truncate">{room.displayName}</h3>
            <div className="flex items-center gap-2 text-xs text-muted-foreground mt-0.5">
              {room.capacity > 0 && (
                <span className="flex items-center gap-1">
                  <Users className="size-3" />
                  {room.capacity}
                </span>
              )}
              {room.floorLabel && (
                <span className="flex items-center gap-1 truncate">
                  <MapPin className="size-3 shrink-0" />
                  <span className="truncate">{room.floorLabel}</span>
                </span>
              )}
            </div>
          </div>

          {/* Action buttons - fixed width, never wraps */}
          <div className="shrink-0 flex items-center gap-1">
            {isBooked ? (
              <Badge variant="default" className="text-xs">
                <Check className="size-3 mr-1" />
                Booked
              </Badge>
            ) : isAvailable ? (
              <>
                {/* Main book button - adds as attendee + location */}
                <Button
                  size="sm"
                  onClick={() => onBook(room, "both")}
                  disabled={isBooking}
                  className="h-7 px-2 text-xs"
                >
                  {isBooking ? <Spinner className="size-3" /> : "Book"}
                </Button>
                {/* Dropdown for additional options */}
                <DropdownMenu>
                  <DropdownMenuTrigger asChild>
                    <Button
                      size="sm"
                      variant="outline"
                      className="h-7 px-1"
                      disabled={isBooking}
                    >
                      <ChevronDown className="size-3" />
                    </Button>
                  </DropdownMenuTrigger>
                  <DropdownMenuContent align="end" className="w-40">
                    <DropdownMenuItem onClick={() => onBook(room, "attendee")}>
                      Add as Attendee
                    </DropdownMenuItem>
                    <DropdownMenuItem onClick={() => onBook(room, "location")}>
                      Set as Location
                    </DropdownMenuItem>
                  </DropdownMenuContent>
                </DropdownMenu>
              </>
            ) : (
              <Badge variant="destructive" className="text-xs">
                Busy
              </Badge>
            )}
          </div>
        </div>
      </CardContent>
    </Card>
  );
}
