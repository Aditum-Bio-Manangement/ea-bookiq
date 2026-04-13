"use client";

import { Users, MapPin, Check } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Spinner } from "@/components/ui/spinner";
import { cn } from "@/lib/utils";
import type { RoomAvailability } from "@/lib/outlook-addin/graph/schedule";

interface RoomCardProps {
  roomAvailability: RoomAvailability;
  onBook: (room: RoomAvailability["room"]) => void;
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

          {/* Action button - fixed width, never wraps */}
          <div className="shrink-0">
            {isBooked ? (
              <Badge variant="default" className="text-xs">
                <Check className="size-3 mr-1" />
                Booked
              </Badge>
            ) : isAvailable ? (
              <Button
                size="sm"
                onClick={() => onBook(room)}
                disabled={isBooking}
                className="h-7 px-3 text-xs"
              >
                {isBooking ? <Spinner className="size-3" /> : "Book"}
              </Button>
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
