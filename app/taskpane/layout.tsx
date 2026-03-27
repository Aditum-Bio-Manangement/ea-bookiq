import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "EA BookIQ - Aditum Bio",
  description: "Intelligent conference room booking for executive assistants",
};

export default function TaskPaneLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return children;
}
