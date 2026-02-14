import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "pptx-glimpse Demo",
  description: "PPTX to SVG converter — server-side demo",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ja">
      <body>{children}</body>
    </html>
  );
}
