import type { Metadata } from "next";
import { UploadViewer } from "@/components/UploadViewer";
import { SITE_URL } from "@/lib/constants";

export const metadata: Metadata = {
  alternates: {
    canonical: "/",
  },
};

const jsonLd = {
  "@context": "https://schema.org",
  "@type": "SoftwareApplication",
  name: "pptx-glimpse",
  description:
    "A TypeScript toolkit for rendering, editing, and saving PowerPoint (PPTX) files. Try the complete workflow locally in your browser.",
  url: SITE_URL,
  applicationCategory: "DeveloperApplication",
  operatingSystem: "Any",
  offers: {
    "@type": "Offer",
    price: "0",
    priceCurrency: "USD",
  },
};

export default function Home() {
  return (
    <main className="app demo-page">
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLd) }}
      />
      <h1 className="visually-hidden">pptx-glimpse browser editor demo</h1>
      <UploadViewer />
    </main>
  );
}
