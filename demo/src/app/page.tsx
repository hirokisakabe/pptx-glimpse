import { UploadViewer } from "@/components/UploadViewer";
import { SITE_URL } from "@/lib/constants";

const jsonLd = {
  "@context": "https://schema.org",
  "@type": "SoftwareApplication",
  name: "pptx-glimpse",
  description:
    "A TypeScript library that converts PowerPoint (PPTX) slides to SVG and PNG. Upload a file and preview instantly.",
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
    <div className="app">
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLd) }}
      />
      <header>
        <h1>pptx-glimpse</h1>
        <p>Upload a PPTX file and preview slides as SVG.</p>
        <p className="description">
          A lightweight JavaScript library for rendering PowerPoint (.pptx) files as SVG or PNG. No
          LibreOffice required.
        </p>
      </header>
      <main>
        <UploadViewer />
      </main>
      <footer>
        <a href="https://github.com/hirokisakabe/pptx-glimpse">GitHub</a>
        <span> | </span>
        <a href="https://www.npmjs.com/package/pptx-glimpse">npm</a>
      </footer>
    </div>
  );
}
