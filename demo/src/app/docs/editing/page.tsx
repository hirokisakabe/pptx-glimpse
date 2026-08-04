import type { Metadata } from "next";
import { Callout } from "nextra/components";
import { DocsPage, DocsPager } from "@/components/docs/DocsPage";

export const metadata: Metadata = {
  title: "Editing presentations",
  description: "Inspect, edit, rerender, and save PPTX files through PptxEditorSession.",
  alternates: { canonical: "/docs/editing" },
};

const toc = [
  { href: "#create-session", label: "Create a session" },
  { href: "#inspect", label: "Inspect editable shapes" },
  { href: "#template-preview", label: "Preview masters and layouts" },
  { href: "#apply", label: "Apply a command" },
  { href: "#history", label: "Selection and history" },
  { href: "#save", label: "Save the presentation" },
] as const;

export default function EditingPage() {
  return (
    <DocsPage
      title="Editing presentations"
      description="PptxEditorSession coordinates reading, editable shape information, commands, SVG rerendering, selection, undo and redo, and PPTX serialization."
      filePath="src/app/docs/editing/page.tsx"
      toc={toc}
    >
      <section id="create-session">
        <h2>Create a session</h2>
        <pre>
          <code>{`import { createPptxEditorSession } from "pptx-glimpse";

const editor = await createPptxEditorSession(pptxBytes, {
  fonts: [{ name: "Inter", data: interFontBytes }],
  textOutput: "text",
});`}</code>
        </pre>
        <p>
          Session input and output are byte-oriented. The API does not read file inputs, update the
          DOM, or start a download for you.
        </p>
      </section>

      <section id="inspect">
        <h2>Inspect editable shapes</h2>
        <p>
          <code>editor.slides</code> contains the current SVG output. Use{" "}
          <code>editor.shapes(slideNumber)</code> to inspect shape bounds, text bodies, image
          replacement information, and stable source handles.
        </p>
        <pre>
          <code>{`const shapes = editor.shapes(1);
const run = shapes
  .flatMap((shape) => shape.textBody?.paragraphs ?? [])
  .flatMap((paragraph) => paragraph.runs)
  .find((candidate) => candidate.handle !== undefined);`}</code>
        </pre>
      </section>

      <section id="template-preview">
        <h2>Preview masters and layouts</h2>
        <p>
          <code>layoutCatalog</code> stays a metadata-only ordered view. Pass a master or layout
          handle to <code>previewLayoutCatalogTarget</code> when a UI needs one SVG thumbnail.
          Previewing does not change the document, selection, undo/redo history, rendered slides, or
          saved PPTX bytes.
        </p>
        <pre>
          <code>{`const layout = editor.layoutCatalog[0]?.layouts[0];
if (layout) {
  const result = await editor.previewLayoutCatalogTarget(layout.handle);
  if (result.ok) {
    thumbnail.innerHTML = result.svg;
    console.log(result.diagnostics);
  } else {
    // preview-handle-not-found | preview-handle-ambiguous
    showThumbnailFallback(result.code);
  }
}`}</code>
        </pre>
        <p>
          Layout previews resolve template backgrounds, normal master/layout shapes, and compatible
          placeholder inheritance without including real slide user content. Unsupported elements
          produce diagnostics. Cache, scheduling, cancellation, PNG conversion, and fallback UI are
          application responsibilities.
        </p>
      </section>

      <section id="apply">
        <h2>Apply a command</h2>
        <pre>
          <code>{`if (run?.handle) {
  const response = await editor.apply({
    kind: "replaceTextRunPlainText",
    handle: run.handle,
    text: "Edited with pptx-glimpse",
  });

  const updatedSvg = response.slides[0]?.svg;
}`}</code>
        </pre>
        <p>
          Successful state-changing commands rerender the current document and create history. Use{" "}
          <code>applyAll</code> when several commands must succeed atomically as one undo entry.
        </p>
        <Callout type="warning">
          <p>
            <strong>Editing support is intentionally constrained.</strong> Shapes without stable
            source handles and operations that cannot be preserved safely are rejected. Check the
            feature support documentation before designing a general-purpose editor.
          </p>
        </Callout>
      </section>

      <section id="history">
        <h2>Selection and history</h2>
        <p>
          Use <code>selectShape</code> and <code>selection</code> to keep a single shape selection.
          The session reconciles selection after edits and history changes.
        </p>
        <pre>
          <code>{`const selectableShape = editor.shapes(1)
  .find((shape) => shape.handle !== undefined);

if (selectableShape?.handle) {
  editor.selectShape(selectableShape.handle);
}

if (editor.history.canUndo) {
  await editor.undo();
}

if (editor.history.canRedo) {
  await editor.redo();
}`}</code>
        </pre>
      </section>

      <section id="save">
        <h2>Save the presentation</h2>
        <pre>
          <code>{`const { pptx } = editor.save();
const copy = new Uint8Array(pptx.byteLength);
copy.set(pptx);

const blob = new Blob([copy.buffer], {
  type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
});`}</code>
        </pre>
        <p>
          <code>save()</code> returns updated PPTX bytes. Your application decides whether to write
          them to disk, upload them, or offer a browser download.
        </p>
      </section>

      <DocsPager
        previous={{ href: "/docs/rendering", label: "Render presentations" }}
        next={{ href: "/docs/browser", label: "Run in the browser" }}
      />
    </DocsPage>
  );
}
