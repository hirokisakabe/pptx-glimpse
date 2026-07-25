import type { Metadata } from "next";
import { DocsPage } from "@/components/docs/DocsPage";

export const metadata: Metadata = {
  title: "Why pptx-glimpse",
  description: "Understand the problems pptx-glimpse is designed to solve.",
  alternates: { canonical: "/docs/why" },
};

const toc = [
  { href: "#use-cases", label: "Where it fits" },
  { href: "#office-process", label: "No office process" },
  { href: "#tradeoffs", label: "Rendering tradeoffs" },
] as const;

export default function WhyPage() {
  return (
    <DocsPage
      title="Why pptx-glimpse"
      description="Use PowerPoint content inside TypeScript applications without managing an office installation or a separate conversion service."
      filePath="src/app/docs/why/page.tsx"
      toc={toc}
    >
      <section id="use-cases">
        <h2>Where it fits</h2>
        <p>
          pptx-glimpse is designed for applications that need to inspect, preview, render, or edit
          presentations as part of their own workflow.
        </p>
        <ul>
          <li>
            <strong>Browser previews:</strong> turn local PPTX bytes into slide thumbnails or
            embeddable SVG without first uploading the presentation.
          </li>
          <li>
            <strong>Application editing:</strong> apply supported changes, rerender affected slides,
            maintain history, and save updated PPTX bytes.
          </li>
          <li>
            <strong>Image pipelines:</strong> produce PNG slides for exports, indexing, or
            vision-capable models.
          </li>
        </ul>
      </section>

      <section id="office-process">
        <h2>No office process</h2>
        <p>
          Conversion runs inside the Node.js or browser process. Your application does not need to
          install Microsoft Office, bundle LibreOffice into a container, spawn a converter process,
          or coordinate a separate conversion server.
        </p>
        <div className="docs-comparison">
          <div>
            <strong>Office-based conversion</strong>
            <p>Requires an office installation and process lifecycle around each conversion.</p>
          </div>
          <div>
            <strong>pptx-glimpse</strong>
            <p>Installs as npm packages and works with bytes already owned by your application.</p>
          </div>
        </div>
      </section>

      <section id="tradeoffs">
        <h2>Rendering tradeoffs</h2>
        <p>
          pptx-glimpse focuses on common static slide content such as text, shapes, images, tables,
          charts, and spatial layout. It does not run PowerPoint itself, so it cannot promise
          pixel-identical output for every presentation or support every PowerPoint feature.
        </p>
        <p>
          Review the <a href="/docs/feature-support">feature support reference</a> and inspect the
          diagnostics returned for each conversion before choosing it for a fidelity-sensitive
          workflow.
        </p>
      </section>
    </DocsPage>
  );
}
