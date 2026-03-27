"use client";

import { TaskPane } from "@/components/outlook-addin/TaskPane";

export default function TaskPanePage() {
  return (
    <main className="h-screen overflow-hidden bg-background">
      <TaskPane />
    </main>
  );
}
