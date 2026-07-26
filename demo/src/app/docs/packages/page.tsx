import type { Metadata } from "next";
import { DocsPage, DocsPager } from "@/components/docs/DocsPage";
import { GITHUB_REPOSITORY_URL } from "@/lib/constants";

export const metadata: Metadata = {
  title: "Choosing a package",
  description: "Choose between pptx-glimpse, @pptx-glimpse/document, and @pptx-glimpse/editor.",
  alternates: { canonical: "/docs/packages" },
};

const toc = [
  { href: "#decision", label: "Decision guide" },
  { href: "#toolkit", label: "pptx-glimpse" },
  { href: "#document", label: "@pptx-glimpse/document" },
  { href: "#editor", label: "@pptx-glimpse/editor" },
] as const;

const repository = `${GITHUB_REPOSITORY_URL}/blob/main`;

export default function PackagesPage() {
  return (
    <DocsPage
      title="Choosing a package"
      description="Begin with the integrated toolkit. Move to a lower-level package only when your application needs to own document semantics or the command lifecycle directly."
      filePath="src/app/docs/packages/page.tsx"
      toc={toc}
    >
      <section id="decision">
        <h2>Decision guide</h2>
        <div className="docs-decision-table" role="table" aria-label="Package decision guide">
          <div role="row">
            <strong role="columnheader">Your application needs to…</strong>
            <strong role="columnheader">Start with</strong>
          </div>
          <div role="row">
            <span role="cell">Render, preview, or run an integrated editor session</span>
            <a href="#toolkit" role="cell">
              <code>pptx-glimpse</code>
            </a>
          </div>
          <div role="row">
            <span role="cell">Read OOXML semantics, author, edit, or write directly</span>
            <a href="#document" role="cell">
              <code>@pptx-glimpse/document</code>
            </a>
          </div>
          <div role="row">
            <span role="cell">Own rendering and writing, but add command history</span>
            <a href="#editor" role="cell">
              <code>@pptx-glimpse/editor</code>
            </a>
          </div>
        </div>
      </section>

      <section id="toolkit">
        <h2 className="docs-package-heading">pptx-glimpse</h2>
        <p>
          This is the recommended starting point for most applications. It converts PPTX to SVG or
          PNG and provides the integrated edit, rerender, history, and save session used by this
          site’s demo. It supports Node.js and browser bundles.
        </p>
        <h3 className="docs-resources-heading">Resources</h3>
        <ul className="docs-resource-links">
          <li>
            <a href="https://www.npmjs.com/package/pptx-glimpse">View on npm</a>
          </li>
          <li>
            <a href={`${repository}/packages/core/README.md`}>Read the package README</a>
          </li>
        </ul>
      </section>

      <section id="document">
        <h2 className="docs-package-heading">@pptx-glimpse/document</h2>
        <p>
          Reads typed source data, derives effective values through a computed view, authors new
          presentations, edits supported existing content, and writes PPTX with round-trip
          preservation. It does not render SVG or PNG.
        </p>
        <h3 className="docs-resources-heading">Resources</h3>
        <ul className="docs-resource-links">
          <li>
            <a href="https://www.npmjs.com/package/@pptx-glimpse/document">View on npm</a>
          </li>
          <li>
            <a href={`${repository}/packages/document/README.md`}>Read the workflow guides</a>
          </li>
          <li>
            <a href={`${repository}/packages/document/docs/feature-support.md`}>
              Review feature support
            </a>
          </li>
        </ul>
      </section>

      <section id="editor">
        <h2 className="docs-package-heading">@pptx-glimpse/editor</h2>
        <p>
          Adds validated commands, selection, warnings, and undo/redo history to a document source
          model. It has no DOM, UI framework, renderer, or file I/O and is currently a{" "}
          <code>0.x</code> package.
        </p>
        <h3 className="docs-resources-heading">Resources</h3>
        <ul className="docs-resource-links">
          <li>
            <a href="https://www.npmjs.com/package/@pptx-glimpse/editor">View on npm</a>
          </li>
          <li>
            <a href={`${repository}/packages/editor/README.md`}>Read the package README</a>
          </li>
        </ul>
      </section>

      <DocsPager
        previous={{ href: "/docs/feature-support", label: "Feature support" }}
        next={{ href: "/docs/api", label: "API Reference" }}
      />
    </DocsPage>
  );
}
