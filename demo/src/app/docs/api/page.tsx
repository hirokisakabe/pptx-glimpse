import type { Metadata } from "next";
import Link from "next/link";
import { DocsPage, DocsPager } from "@/components/docs/DocsPage";

export const metadata: Metadata = {
  title: "High-level API",
  description: "A practical reference for the public pptx-glimpse exports.",
  alternates: { canonical: "/docs/api" },
};

const toc = [
  { href: "#conversion", label: "Conversion" },
  { href: "#editor-session", label: "Editor session" },
  { href: "#fonts", label: "Fonts and PNG" },
  { href: "#errors", label: "Errors and warnings" },
] as const;

const conversionApis = [
  ["convertPptxToSvg(input, options?)", "Convert PPTX bytes to SVG slide results."],
  ["convertPptxToPng(input, options?)", "Convert PPTX bytes to PNG slide results."],
  ["renderPptxSourceModelToSvg(source, options?)", "Render an already parsed source model."],
  ["collectUsedFonts(input)", "Inspect font names without performing a full render."],
] as const;

const sessionApis = [
  ["slides", "The current SVG results."],
  ["shapes(slideNumber)", "Editable shape information for one slide."],
  ["apply(command)", "Apply one command and create one history entry."],
  ["applyAll(commands)", "Apply commands atomically as one history entry."],
  ["selectShape(handle)", "Select one editable shape."],
  ["undo() / redo()", "Move through session history and rerender."],
  ["save()", "Validate and serialize the current document to PPTX bytes."],
] as const;

export default function ApiPage() {
  return (
    <DocsPage
      title="High-level API"
      description="This page covers the stable entry points most applications need. TypeScript remains the source of truth for complete signatures and exported types."
      filePath="src/app/docs/api/page.tsx"
      toc={toc}
    >
      <p>
        For complete TypeScript signatures, options, return values, and related types, open the{" "}
        <Link href="/docs/api-reference">generated API reference</Link>. It is generated from the
        public Node.js and browser entry points and their JSDoc.
      </p>

      <section id="conversion">
        <h2>Conversion</h2>
        <dl className="docs-api-list">
          {conversionApis.map(([signature, description]) => (
            <div key={signature}>
              <dt>
                <code>{signature}</code>
              </dt>
              <dd>{description}</dd>
            </div>
          ))}
        </dl>
        <p>
          Conversion reports expose slide output, diagnostics, and support coverage. Inputs and
          binary outputs use <code>Uint8Array</code>.
        </p>
      </section>

      <section id="editor-session">
        <h2>Editor session</h2>
        <p>
          Create a session with <code>createPptxEditorSession(input, renderOptions?)</code>. The
          same API is available from the Node.js and browser entry points.
        </p>
        <dl className="docs-api-list">
          {sessionApis.map(([signature, description]) => (
            <div key={signature}>
              <dt>
                <code>{signature}</code>
              </dt>
              <dd>{description}</dd>
            </div>
          ))}
        </dl>
      </section>

      <section id="fonts">
        <h2>Fonts and PNG</h2>
        <dl className="docs-api-list">
          <div>
            <dt>
              <code>createFontMapping(overrides)</code>
            </dt>
            <dd>Merge application mappings with the default font mapping.</dd>
          </div>
          <div>
            <dt>
              <code>getMappedFont(name, mapping)</code>
            </dt>
            <dd>Resolve a presentation font name case-insensitively.</dd>
          </div>
          <div>
            <dt>
              <code>initResvgWasm(wasm)</code>
            </dt>
            <dd>Initialize PNG rasterization explicitly in browser-like environments.</dd>
          </div>
        </dl>
      </section>

      <section id="errors">
        <h2>Errors and warnings</h2>
        <p>
          Expected high-level editor failures throw <code>PptxEditorError</code>. Narrow unknown
          errors with <code>isPptxEditorError</code> and branch on its machine-readable code.
        </p>
        <pre>
          <code>{`try {
  await editor.undo();
} catch (error) {
  if (isPptxEditorError(error)) {
    console.error(error.code, error.message);
  } else {
    throw error;
  }
}`}</code>
        </pre>
        <p>
          <strong>Warnings do not throw.</strong> Successful commands can return warnings, such as
          replacing a media part referenced by multiple images. Inspect the response before
          discarding it.
        </p>
      </section>

      <DocsPager
        previous={{ href: "/docs/nodejs", label: "Run in Node.js" }}
        next={{ href: "/docs/packages", label: "Choose a package" }}
      />
    </DocsPage>
  );
}
