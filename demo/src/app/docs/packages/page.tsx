import type { Metadata } from "next";
import { DocsPage, DocsPager } from "@/components/docs/DocsPage";

export const metadata: Metadata = {
  title: "Choose a package",
  description: "Choose between pptx-glimpse, @pptx-glimpse/document, and @pptx-glimpse/editor.",
  alternates: { canonical: "/docs/packages" },
};

const toc = [
  { href: "#toolkit", label: "pptx-glimpse" },
  { href: "#document", label: "Document package" },
  { href: "#editor", label: "Editor package" },
  { href: "#decision", label: "Decision guide" },
] as const;

const repository = "https://github.com/hirokisakabe/pptx-glimpse/blob/main";

export default function PackagesPage() {
  return (
    <DocsPage
      eyebrow="Reference / Packages"
      title="Choose the highest useful layer."
      description="Begin with the integrated toolkit. Move to a lower-level package only when your application needs to own document semantics or the command lifecycle directly."
      toc={toc}
    >
      <section id="toolkit">
        <p className="docs-package-kicker">Recommended starting point</p>
        <h2>
          <code>pptx-glimpse</code>
        </h2>
        <p>
          Converts PPTX to SVG or PNG and provides the integrated edit, rerender, history, and save
          session used by this site’s demo. It supports Node.js and browser bundles.
        </p>
        <p>
          <a href="https://www.npmjs.com/package/pptx-glimpse">npm package</a>
          {" · "}
          <a href={`${repository}/packages/core/README.md`}>package README</a>
        </p>
      </section>

      <section id="document">
        <p className="docs-package-kicker">OOXML foundation</p>
        <h2>
          <code>@pptx-glimpse/document</code>
        </h2>
        <p>
          Reads typed source data, derives effective values through a computed view, authors new
          presentations, edits supported existing content, and writes PPTX with round-trip
          preservation. It does not render SVG or PNG.
        </p>
        <p>
          <a href="https://www.npmjs.com/package/@pptx-glimpse/document">npm package</a>
          {" · "}
          <a href={`${repository}/packages/document/README.md`}>workflow guides</a>
          {" · "}
          <a href={`${repository}/packages/document/docs/feature-support.md`}>
            feature support matrix
          </a>
        </p>
      </section>

      <section id="editor">
        <p className="docs-package-kicker">Headless command state</p>
        <h2>
          <code>@pptx-glimpse/editor</code>
        </h2>
        <p>
          Adds validated commands, selection, warnings, and undo/redo history to a document source
          model. It has no DOM, UI framework, renderer, or file I/O and is currently a{" "}
          <code>0.x</code> package.
        </p>
        <p>
          <a href="https://www.npmjs.com/package/@pptx-glimpse/editor">npm package</a>
          {" · "}
          <a href={`${repository}/packages/editor/README.md`}>package README</a>
        </p>
      </section>

      <section id="decision">
        <h2>Decision guide</h2>
        <div className="docs-decision-table" role="table" aria-label="Package decision guide">
          <div role="row">
            <strong role="columnheader">Your application needs to…</strong>
            <strong role="columnheader">Start with</strong>
          </div>
          <div role="row">
            <span role="cell">Render, preview, or run an integrated editor session</span>
            <code role="cell">pptx-glimpse</code>
          </div>
          <div role="row">
            <span role="cell">Read OOXML semantics, author, edit, or write directly</span>
            <code role="cell">@pptx-glimpse/document</code>
          </div>
          <div role="row">
            <span role="cell">Own rendering and writing, but add command history</span>
            <code role="cell">@pptx-glimpse/editor</code>
          </div>
        </div>
      </section>

      <DocsPager previous={{ href: "/docs/api", label: "High-level API" }} />
    </DocsPage>
  );
}
