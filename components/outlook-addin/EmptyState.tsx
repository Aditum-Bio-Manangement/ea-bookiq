"use client";

import { CalendarX, Clock, Building2, AlertCircle } from "lucide-react";
import { Button } from "@/components/ui/button";

type EmptyStateType = "no-rooms" | "no-time" | "no-office" | "error";

interface EmptyStateProps {
  type: EmptyStateType;
  message?: string;
  onAction?: () => void;
  actionLabel?: string;
}

const icons: Record<EmptyStateType, React.ReactNode> = {
  "no-rooms": <CalendarX className="size-12 text-muted-foreground" />,
  "no-time": <Clock className="size-12 text-muted-foreground" />,
  "no-office": <Building2 className="size-12 text-muted-foreground" />,
  error: <AlertCircle className="size-12 text-destructive" />,
};

const titles: Record<EmptyStateType, string> = {
  "no-rooms": "No Rooms Available",
  "no-time": "Set Meeting Time",
  "no-office": "Select Office",
  error: "Something Went Wrong",
};

const descriptions: Record<EmptyStateType, string> = {
  "no-rooms":
    "All conference rooms are busy during this time. Try a different time slot.",
  "no-time":
    "Please set the meeting start and end time to see available rooms.",
  "no-office":
    "Unable to determine your office location. Please select your office manually.",
  error: "An error occurred while loading rooms. Please try again.",
};

export function EmptyState({
  type,
  message,
  onAction,
  actionLabel,
}: EmptyStateProps) {
  return (
    <div className="flex flex-col items-center justify-center p-8 text-center">
      {icons[type]}
      <h3 className="mt-4 font-semibold text-foreground">{titles[type]}</h3>
      <p className="mt-2 text-sm text-muted-foreground max-w-[250px]">
        {message || descriptions[type]}
      </p>
      {onAction && actionLabel && (
        <Button onClick={onAction} variant="outline" className="mt-4">
          {actionLabel}
        </Button>
      )}
    </div>
  );
}
