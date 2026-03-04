import type { Metadata } from "next";
import { Analytics } from "@vercel/analytics/next";
import { SITE_URL } from "@/lib/constants";
import "./globals.css";
const TITLE = "pptx-glimpse Demo – PPTX to SVG Converter";
const DESCRIPTION =
  "Upload a PPTX file and instantly preview slides as SVG. A free, open-source TypeScript library for converting PowerPoint presentations.";

export const metadata: Metadata = {
  metadataBase: new URL(SITE_URL),
  title: TITLE,
  description: DESCRIPTION,
  keywords: ["PPTX", "SVG", "PowerPoint", "converter", "TypeScript", "presentation", "slides"],
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
  alternates: {
    canonical: "/",
  },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>
        {children}
        <Analytics />
      </body>
    </html>
  );
}
