import type { Metadata } from "next";
import Link from "next/link";
import { DocsCallout, DocsPage, DocsPager } from "@/components/docs/DocsPage";

export const metadata: Metadata = {
  title: "Getting started",
  description: "Install pptx-glimpse and render your first PowerPoint slide.",
  alternates: { canonical: "/docs/getting-started" },
};

const toc = [
  { href: "#requirements", label: "Requirements" },
  { href: "#install", label: "Install" },
  { href: "#first-render", label: "Render your first slide" },
  { href: "#next", label: "Choose what comes next" },
] as const;

export default function GettingStartedPage() {
  return (
    <DocsPage
      eyebrow="Start here"
      title="Render your first presentation."
      description="Install the high-level package, pass it PPTX bytes, and receive SVG or PNG output together with diagnostics and support coverage."
      toc={toc}
    >
      <section id="requirements">
        <h2>Requirements</h2>
        <ul>
          <li>Node.js 22 or later for Node.js applications and package tooling.</li>
          <li>A modern browser bundler for browser applications.</li>
          <li>Font files or known font directories when presentation text must match closely.</li>
        </ul>
        <DocsCallout title="No office installation required">
          <p>
            pptx-glimpse runs in-process. Rendering does not start PowerPoint or LibreOffice and
            does not require a conversion service.
          </p>
        </DocsCallout>
      </section>

      <section id="install">
        <h2>Install</h2>
        <pre>
          <code>npm install pptx-glimpse</code>
        </pre>
        <p>
          The high-level package includes compatible document and editor packages as runtime
          dependencies. Install a lower-level package directly only when your application imports
          it.
        </p>
      </section>

      <section id="first-render">
        <h2>Render your first slide</h2>
        <p>In Node.js, read the file as bytes and pass it directly to the converter.</p>
        <pre>
          <code>{`import { readFile, writeFile } from "node:fs/promises";
import { convertPptxToPng } from "pptx-glimpse";

const pptx = await readFile("presentation.pptx");
const report = await convertPptxToPng(pptx, {
  slides: [1],
  width: 1920,
});

const firstSlide = report.slides[0];
if (firstSlide) {
  await writeFile("slide-1.png", firstSlide.png);
}`}</code>
        </pre>
        <p>
          Each conversion returns <code>slides</code>, <code>diagnostics</code>, and{" "}
          <code>supportCoverage</code>. Check diagnostics when content is skipped or rendered with a
          fallback.
        </p>
      </section>

      <section id="next">
        <h2>Choose what comes next</h2>
        <div className="docs-link-list">
          <Link href="/docs/rendering">
            <strong>Control rendering</strong>
            <span>Selected slides, output size, fonts, SVG text, and browser PNG.</span>
          </Link>
          <Link href="/docs/editing">
            <strong>Build an editing flow</strong>
            <span>Inspect shapes, apply commands, manage history, and save PPTX bytes.</span>
          </Link>
          <Link href="/docs/packages">
            <strong>Work below the high-level API</strong>
            <span>Read, author, or edit documents with the lower-level packages.</span>
          </Link>
        </div>
      </section>

      <DocsPager
        previous={{ href: "/docs", label: "Overview" }}
        next={{ href: "/docs/rendering", label: "Render presentations" }}
      />
    </DocsPage>
  );
}
