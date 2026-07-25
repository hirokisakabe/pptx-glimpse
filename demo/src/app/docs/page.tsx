import type { Metadata } from "next";

const REPOSITORY_URL = "https://github.com/hirokisakabe/pptx-glimpse";
const TOOLKIT_DOCS_URL = `${REPOSITORY_URL}/blob/main/packages/core/README.md`;
const DOCUMENT_DOCS_URL = `${REPOSITORY_URL}/blob/main/packages/document/README.md`;
const EDITOR_DOCS_URL = `${REPOSITORY_URL}/blob/main/packages/editor/README.md`;

export const metadata: Metadata = {
  title: "Documentation",
  description:
    "Choose among the pptx-glimpse toolkit and its lower-level document and headless editor packages.",
  keywords: ["pptx-glimpse", "PPTX rendering", "PPTX editing", "PowerPoint", "OOXML", "TypeScript"],
  alternates: {
    canonical: "/docs",
  },
  openGraph: {
    title: "pptx-glimpse Documentation",
    description:
      "Choose among the pptx-glimpse toolkit and its lower-level document and headless editor packages.",
    url: "/docs",
  },
  twitter: {
    card: "summary_large_image",
    title: "pptx-glimpse Documentation",
    description:
      "Choose among the pptx-glimpse toolkit and its lower-level document and headless editor packages.",
  },
};

const documentWorkflows = [
  "Read typed PPTX source data",
  "Resolve a non-mutating computed view",
  "Author presentations from scratch",
  "Edit supported content in an existing PPTX",
  "Write PPTX bytes with round-trip preservation",
];

const toolkitWorkflows = [
  "Render complete decks or selected slides as SVG or PNG",
  "Open one editor session for read, edit, rerender, and save",
  "Run the same high-level workflow in Node.js and browsers",
];

const editorWorkflows = [
  "Apply validated editing commands to a document",
  "Track selection and editing state",
  "Integrate undo and redo into your own application",
];

export default function DocsPage() {
  return (
    <main className="docs-page">
      <section className="docs-intro" aria-labelledby="docs-heading">
        <p className="eyebrow">Project documentation</p>
        <h1 id="docs-heading">Choose the path that matches your PPTX task.</h1>
        <p>
          The project has three public packages. Start with <code>pptx-glimpse</code> for an
          integrated render-and-edit workflow. Reach for the lower-level packages when you need
          document semantics or a headless editing state machine.
        </p>
      </section>

      <section className="package-routes" aria-label="Documentation routes">
        <article className="package-route toolkit-route">
          <div className="route-output" aria-hidden="true">
            .pptx ↔ .svg
          </div>
          <p className="route-label">Recommended starting point</p>
          <h2>pptx-glimpse</h2>
          <p className="package-summary">
            The high-level toolkit that brings PPTX rendering, editing, SVG rerendering, and saving
            together in one browser- and Node.js-friendly API.
          </p>
          <ul>
            {toolkitWorkflows.map((workflow) => (
              <li key={workflow}>{workflow}</li>
            ))}
          </ul>
          <div className="route-links">
            <a className="primary-doc-link" href={TOOLKIT_DOCS_URL}>
              Start with the toolkit <span aria-hidden="true">→</span>
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

        <article className="package-route editor-route">
          <div className="route-output" aria-hidden="true">
            commands
          </div>
          <p className="route-label">Headless editing state</p>
          <h2>@pptx-glimpse/editor</h2>
          <p className="package-summary">
            The lower-level editing engine for validated commands, selection, and undo/redo. It
            supplies no UI and does not render slides by itself.
          </p>
          <ul>
            {editorWorkflows.map((workflow) => (
              <li key={workflow}>{workflow}</li>
            ))}
          </ul>
          <div className="route-links">
            <a className="primary-doc-link" href={EDITOR_DOCS_URL}>
              Build a headless editor <span aria-hidden="true">→</span>
            </a>
            <a href="https://www.npmjs.com/package/@pptx-glimpse/editor">View package on npm</a>
          </div>
        </article>
      </section>

      <aside className="demo-boundary" aria-label="Demo scope">
        <strong>About this site’s demo</strong>
        <p>
          The home page demonstrates the recommended <code>pptx-glimpse</code> workflow: open,
          render, edit, rerender, and save. Its focused editor covers a supported subset of
          operations; use the document and editor guides for lower-level API coverage and
          constraints.
        </p>
      </aside>
    </main>
  );
}
