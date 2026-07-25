import type { Metadata } from "next";
import Link from "next/link";
import { DocsPage, DocsPager } from "@/components/docs/DocsPage";

export const metadata: Metadata = {
  title: "Documentation",
  description:
    "Learn how to render, edit, and save PowerPoint files with pptx-glimpse in Node.js and the browser.",
  alternates: { canonical: "/docs" },
};

const toc = [
  { href: "#choose-a-path", label: "Choose a path" },
  { href: "#workflow", label: "The high-level workflow" },
  { href: "#packages", label: "Packages" },
] as const;

const paths = [
  {
    marker: ".svg",
    title: "Render presentations",
    description: "Convert a complete deck or selected slides to SVG or PNG.",
    href: "/docs/rendering",
    link: "Rendering guide",
  },
  {
    marker: "edit",
    title: "Build an editing flow",
    description: "Read, inspect, edit, rerender, undo, redo, and save through one session.",
    href: "/docs/editing",
    link: "Editing guide",
  },
  {
    marker: "API",
    title: "Look up an API",
    description: "Find the high-level exports, important types, runtime notes, and error contract.",
    href: "/docs/api",
    link: "API overview",
  },
] as const;

export default function DocsOverviewPage() {
  return (
    <DocsPage
      eyebrow="Documentation / Overview"
      title="From PowerPoint bytes to an application."
      description="Use pptx-glimpse to render presentations or build an editing workflow without running PowerPoint, LibreOffice, or a conversion server."
      toc={toc}
    >
      <section id="choose-a-path">
        <h2>Choose a path</h2>
        <p>
          Start with the job your application needs to perform. The high-level package covers the
          common render and edit lifecycles in both Node.js and browser bundles.
        </p>
        <div className="docs-path-grid">
          {paths.map((path) => (
            <Link className="docs-path" href={path.href} key={path.href}>
              <span>{path.marker}</span>
              <h3>{path.title}</h3>
              <p>{path.description}</p>
              <strong>{path.link} →</strong>
            </Link>
          ))}
        </div>
      </section>

      <section id="workflow">
        <h2>The high-level workflow</h2>
        <p>
          The browser demo follows the same lifecycle exposed by the public API. PPTX input, PNG
          output, and saved presentations use <code>Uint8Array</code>; SVG output is a string. Your
          application owns file selection, UI, and storage.
        </p>
        <ol className="docs-process">
          <li>
            <span>01</span>
            <div>
              <strong>Open</strong>
              <p>Read PPTX bytes from a file, request, object store, or other source.</p>
            </div>
          </li>
          <li>
            <span>02</span>
            <div>
              <strong>Render or edit</strong>
              <p>Convert slides directly, or keep an editor session for repeated changes.</p>
            </div>
          </li>
          <li>
            <span>03</span>
            <div>
              <strong>Use the output</strong>
              <p>Embed SVG, write PNG, or save updated PPTX bytes.</p>
            </div>
          </li>
        </ol>
      </section>

      <section id="packages">
        <h2>Three packages, one dependency direction</h2>
        <p>
          Most applications should begin with <code>pptx-glimpse</code>. The document and editor
          packages are lower-level building blocks for applications that need to own more of the
          workflow.
        </p>
        <ol className="docs-package-stack" aria-label="Packages from highest to lowest level">
          <li>
            <code>pptx-glimpse</code>
            <span>render + high-level editing</span>
          </li>
          <li>
            <code>@pptx-glimpse/editor</code>
            <span>commands + selection + history</span>
          </li>
          <li>
            <code>@pptx-glimpse/document</code>
            <span>OOXML source + computed view + writing</span>
          </li>
        </ol>
        <p>
          See <Link href="/docs/packages">Choose a package</Link> for package boundaries, stability,
          and links to the detailed lower-level guides.
        </p>
      </section>

      <DocsPager next={{ href: "/docs/getting-started", label: "Getting started" }} />
    </DocsPage>
  );
}
