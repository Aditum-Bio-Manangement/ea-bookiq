import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "EA Book IQ - Aditum Bio",
  description: "Intelligent conference room booking for executive assistants",
};

export default function TaskPaneLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return children;
}
