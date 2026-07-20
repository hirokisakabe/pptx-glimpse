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
    "A TypeScript library that converts PowerPoint (PPTX) slides to SVG. Upload a file and preview it locally in the browser.",
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
    <div className="app demo-page">
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLd) }}
      />
      <header className="demo-intro">
        <p className="eyebrow">Browser rendering demo</p>
        <h1>See your PowerPoint as SVG.</h1>
        <p className="description">
          Preview PowerPoint slides as SVG in the browser. No upload service, no LibreOffice.
        </p>
      </header>
      <main>
        <UploadViewer />
      </main>
    </div>
  );
}
