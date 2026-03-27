"use client";

import { Users, Video, Monitor, Accessibility, MapPin, Check } from "lucide-react";
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
        "transition-all",
        !isAvailable && "opacity-60",
        isBooked && "border-primary bg-primary/5"
      )}
    >
      <CardContent className="p-4">
        <div className="flex items-start justify-between gap-3">
          <div className="flex-1 min-w-0">
            <div className="flex items-center gap-2 mb-1">
              <h3 className="font-medium text-sm truncate">{room.displayName}</h3>
              {isBooked && (
                <Badge variant="default" className="shrink-0">
                  <Check className="size-3 mr-1" />
                  Booked
                </Badge>
              )}
            </div>

            <div className="flex flex-wrap items-center gap-2 text-xs text-muted-foreground mb-2">
              {room.capacity > 0 && (
                <span className="flex items-center gap-1">
                  <Users className="size-3" />
                  {room.capacity}
                </span>
              )}
              {room.floorLabel && (
                <span className="flex items-center gap-1">
                  <MapPin className="size-3" />
                  {room.floorLabel}
                </span>
              )}
              {room.videoDeviceName && (
                <span className="flex items-center gap-1">
                  <Video className="size-3" />
                  Video
                </span>
              )}
              {room.displayDeviceName && (
                <span className="flex items-center gap-1">
                  <Monitor className="size-3" />
                  Display
                </span>
              )}
              {room.isWheelChairAccessible && (
                <span className="flex items-center gap-1">
                  <Accessibility className="size-3" />
                </span>
              )}
            </div>

            {room.tags.length > 0 && (
              <div className="flex flex-wrap gap-1">
                {room.tags.slice(0, 3).map((tag) => (
                  <Badge key={tag} variant="outline" className="text-xs py-0">
                    {tag}
                  </Badge>
                ))}
                {room.tags.length > 3 && (
                  <Badge variant="outline" className="text-xs py-0">
                    +{room.tags.length - 3}
                  </Badge>
                )}
              </div>
            )}
          </div>

          <div className="shrink-0">
            {isBooked ? (
              <Badge variant="secondary" className="text-xs">
                Added
              </Badge>
            ) : isAvailable ? (
              <Button
                size="sm"
                onClick={() => onBook(room)}
                disabled={isBooking}
                className="min-w-[70px]"
              >
                {isBooking ? <Spinner className="size-4" /> : "Book"}
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
