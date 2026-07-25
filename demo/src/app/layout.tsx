import type { Metadata } from "next";
import { Analytics } from "@vercel/analytics/next";
import { SiteFooter } from "@/components/SiteFooter";
import { SiteHeader } from "@/components/SiteHeader";
import { SITE_URL } from "@/lib/constants";
import "./globals.css";
const TITLE = "pptx-glimpse Demo - Render, Edit, and Save PPTX in Your Browser";
const DESCRIPTION =
  "Open, render, edit, and save PowerPoint files locally in your browser without uploading them.";

export const metadata: Metadata = {
  metadataBase: new URL(SITE_URL),
  title: {
    default: TITLE,
    template: "%s | pptx-glimpse",
  },
  description: DESCRIPTION,
  keywords: ["PPTX", "PowerPoint", "renderer", "editor", "TypeScript", "presentation", "slides"],
  openGraph: {
    title: TITLE,
    description: DESCRIPTION,
    url: SITE_URL,
    siteName: "pptx-glimpse",
    type: "website",
    locale: "en_US",
  },
  twitter: {
    card: "summary_large_image",
    title: TITLE,
    description: DESCRIPTION,
  },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>
        <div className="site-shell">
          <SiteHeader />
          {children}
          <SiteFooter />
        </div>
        <Analytics />
      </body>
    </html>
  );
}
