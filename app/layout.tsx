import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Relay - Calendar Operations",
  description: "Self-hosted Schoolbox to Google Workspace calendar synchronization.",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="en"><body>{children}</body></html>;
}
