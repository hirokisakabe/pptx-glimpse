import type { Metadata } from "next";

const REPOSITORY_URL = "https://github.com/hirokisakabe/pptx-glimpse";
const RENDERING_DOCS_URL = `${REPOSITORY_URL}/blob/main/README.md`;
const DOCUMENT_DOCS_URL = `${REPOSITORY_URL}/blob/main/packages/document/README.md`;

export const metadata: Metadata = {
  title: "Documentation",
  description:
    "Choose the pptx-glimpse package and documentation for rendering or document workflows.",
  keywords: ["pptx-glimpse", "PPTX rendering", "PPTX editing", "PowerPoint", "OOXML", "TypeScript"],
  alternates: {
    canonical: "/docs",
  },
  openGraph: {
    title: "pptx-glimpse Documentation",
    description:
      "Choose the pptx-glimpse package and documentation for rendering or document workflows.",
    url: "/docs",
  },
  twitter: {
    card: "summary_large_image",
    title: "pptx-glimpse Documentation",
    description:
      "Choose the pptx-glimpse package and documentation for rendering or document workflows.",
  },
};

const documentWorkflows = [
  "Read typed PPTX source data",
  "Resolve a non-mutating computed view",
  "Author presentations from scratch",
  "Edit supported content in an existing PPTX",
  "Write PPTX bytes with round-trip preservation",
];

export default function DocsPage() {
  return (
    <main className="docs-page">
      <section className="docs-intro" aria-labelledby="docs-heading">
        <p className="eyebrow">Project documentation</p>
        <h1 id="docs-heading">Choose the path that matches your PPTX task.</h1>
        <p>
          pptx-glimpse has two public packages. One turns slides into visual output; the other
          provides lower-level document semantics. Start with the package boundary, then follow its
          maintained guide for API details.
        </p>
      </section>

      <section className="package-routes" aria-label="Documentation routes">
        <article className="package-route rendering-route">
          <div className="route-output" aria-hidden="true">
            .svg / .png
          </div>
          <p className="route-label">Render and preview</p>
          <h2>pptx-glimpse</h2>
          <p className="package-summary">
            The high-level rendering package for converting PPTX slides to SVG or PNG in Node.js and
            browsers.
          </p>
          <ul>
            <li>Render a complete deck or selected slides</li>
            <li>Build browser previews and thumbnails</li>
            <li>Provide fonts and choose SVG text output</li>
          </ul>
          <div className="route-links">
            <a className="primary-doc-link" href={RENDERING_DOCS_URL}>
              Read the rendering guide <span aria-hidden="true">→</span>
            </a>
            <a href="https://www.npmjs.com/package/pptx-glimpse">View package on npm</a>
          </div>
        </article>

        <article className="package-route document-route">
          <div className="route-output" aria-hidden="true">
            .pptx
          </div>
          <p className="route-label">Read, author, edit, and write</p>
          <h2>@pptx-glimpse/document</h2>
          <p className="package-summary">
            The lower-level OOXML document foundation. It owns the editable source model and a
            derived computed view; it does not render SVG or PNG.
          </p>
          <ul>
            {documentWorkflows.map((workflow) => (
              <li key={workflow}>{workflow}</li>
            ))}
          </ul>
          <div className="route-links">
            <a className="primary-doc-link" href={DOCUMENT_DOCS_URL}>
              Choose a document workflow <span aria-hidden="true">→</span>
            </a>
            <a href="https://www.npmjs.com/package/@pptx-glimpse/document">View package on npm</a>
          </div>
        </article>
      </section>

      <aside className="demo-boundary" aria-label="Demo scope">
        <strong>About this site’s demo</strong>
        <p>
          The home page opens with browser-based SVG rendering. Its focused editor demonstrates a
          supported subset of editing and download operations, but it is not an interactive
          reference for every @pptx-glimpse/document workflow. Use the document package guide for
          current API coverage and constraints.
        </p>
      </aside>
    </main>
  );
}
