import { UploadViewer } from "@/components/UploadViewer";
import { SITE_URL } from "@/lib/constants";

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
    <div className="app">
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLd) }}
      />
      <header>
        <h1>pptx-glimpse</h1>
        <p className="description">
          Preview PowerPoint slides as SVG in the browser. No upload service, no LibreOffice.
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
