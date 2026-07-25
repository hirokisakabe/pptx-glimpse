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
    <div className="app demo-page">
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLd) }}
      />
      <header className="demo-intro">
        <p className="eyebrow">Browser rendering + editing demo</p>
        <h1>Open it. Edit it. Save it back to PPTX.</h1>
        <p className="description">
          View, edit, and resave PowerPoint files entirely in your browser. Your files are never
          sent to a server, and no LibreOffice installation is required.
        </p>
        <ol className="capability-flow" aria-label="Demo workflow">
          <li>
            <span>01</span> View
          </li>
          <li>
            <span>02</span> Edit
          </li>
          <li>
            <span>03</span> Save PPTX
          </li>
        </ol>
      </header>
      <main>
        <UploadViewer />
      </main>
    </div>
  );
}
