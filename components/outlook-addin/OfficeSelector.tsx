"use client";

import { Building2, Check } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import type { OfficeConfig } from "@/lib/outlook-addin/config/offices";
import { cn } from "@/lib/utils";

interface OfficeSelectorProps {
  offices: OfficeConfig[];
  selectedOffice?: OfficeConfig;
  onSelect: (office: OfficeConfig) => void;
  showHeader?: boolean;
}

export function OfficeSelector({
  offices,
  selectedOffice,
  onSelect,
  showHeader = true,
}: OfficeSelectorProps) {
  return (
    <div>
      {showHeader && (
        <div className="mb-4">
          <h3 className="font-semibold text-foreground flex items-center gap-2">
            <Building2 className="size-4" />
            Select Office
          </h3>
          <p className="text-sm text-muted-foreground mt-1">
            Choose which office location to book rooms from.
          </p>
        </div>
      )}

      <div className="flex flex-col gap-2">
        {offices.map((office) => {
          const isSelected = selectedOffice?.id === office.id;
          return (
            <Card
              key={office.id}
              className={cn(
                "cursor-pointer transition-all hover:border-primary/50",
                isSelected && "border-primary bg-primary/5"
              )}
              onClick={() => onSelect(office)}
            >
              <CardContent className="p-3">
                <div className="flex items-center justify-between">
                  <div>
                    <h4 className="font-medium text-sm">{office.displayName}</h4>
                    <p className="text-xs text-muted-foreground">
                      {office.securityGroupEmail}
                    </p>
                  </div>
                  {isSelected && (
                    <Check className="size-4 text-primary shrink-0" />
                  )}
                </div>
              </CardContent>
            </Card>
          );
        })}
      </div>
    </div>
  );
}

interface OfficeToggleProps {
  offices: OfficeConfig[];
  selectedOffice: OfficeConfig;
  onSelect: (office: OfficeConfig) => void;
}

export function OfficeToggle({
  offices,
  selectedOffice,
  onSelect,
}: OfficeToggleProps) {
  return (
    <div className="flex gap-1 p-1 bg-muted rounded-lg">
      {offices.map((office) => {
        const isSelected = selectedOffice.id === office.id;
        return (
          <Button
            key={office.id}
            variant={isSelected ? "default" : "ghost"}
            size="sm"
            className={cn(
              "flex-1 text-xs",
              !isSelected && "hover:bg-background"
            )}
            onClick={() => onSelect(office)}
          >
            {office.name}
          </Button>
        );
      })}
    </div>
  );
}
