import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "AB Book IQ - Aditum Bio",
  description: "Intelligent conference room booking for Aditum Bio employees. An Outlook add-in by Aditum Bio.",
};

export default function TaskPaneLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return children;
}
