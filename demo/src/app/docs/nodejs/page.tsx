import type { Metadata } from "next";
import { DocsCallout, DocsPage, DocsPager } from "@/components/docs/DocsPage";

export const metadata: Metadata = {
  title: "Run in Node.js",
  description: "Render and edit PPTX files in Node.js with explicit font and file handling.",
  alternates: { canonical: "/docs/nodejs" },
};

const toc = [
  { href: "#files", label: "Read and write files" },
  { href: "#fonts", label: "Resolve fonts" },
  { href: "#determinism", label: "Keep output deterministic" },
  { href: "#containers", label: "Containers and services" },
] as const;

export default function NodejsPage() {
  return (
    <DocsPage
      eyebrow="Runtime / Node.js"
      title="Render inside the Node.js process."
      description="Use Node.js 22 or later to convert and edit PPTX bytes without spawning an office process. File and storage operations remain explicit in your application."
      toc={toc}
    >
      <section id="files">
        <h2>Read and write files</h2>
        <pre>
          <code>{`import { readFile, writeFile } from "node:fs/promises";
import { createPptxEditorSession } from "pptx-glimpse";

const input = await readFile("input.pptx");
const editor = await createPptxEditorSession(input, {
  skipSystemFonts: false,
});

await editor.apply(command);
await writeFile("edited.pptx", editor.save().pptx);`}</code>
        </pre>
        <p>
          Node.js <code>Buffer</code> is accepted because it is a subclass of{" "}
          <code>Uint8Array</code>. Public APIs remain byte-oriented and do not depend on Node.js
          file paths.
        </p>
      </section>

      <section id="fonts">
        <h2>Resolve fonts</h2>
        <p>
          Set <code>skipSystemFonts: false</code> to let Node.js search known OS font directories.
          Use <code>fontDirs</code> to add application-owned directories, or supply{" "}
          <code>fonts</code> as bytes for the most portable setup.
        </p>
        <pre>
          <code>{`const report = await convertPptxToPng(pptx, {
  skipSystemFonts: true,
  fontDirs: ["/app/fonts"],
  fontMapping: {
    "Corporate Sans": "Inter",
  },
});`}</code>
        </pre>
      </section>

      <section id="determinism">
        <h2>Keep output deterministic</h2>
        <p>
          System fonts vary across developer machines, CI runners, and production images. For stable
          output, bundle the exact font files you need, set <code>skipSystemFonts: true</code>, and
          use only explicit font bytes or directories.
        </p>
        <DocsCallout title="Cache repeated work">
          <p>
            For repeated slide renders, read the PPTX once with <code>@pptx-glimpse/document</code>{" "}
            and pass its source model to <code>renderPptxSourceModelToSvg</code>.
          </p>
        </DocsCallout>
      </section>

      <section id="containers">
        <h2>Containers and services</h2>
        <p>
          pptx-glimpse runs in the event loop and does not manage job queues, concurrency limits, or
          storage. Bound input size and concurrency at the application layer when processing
          untrusted presentations in a service.
        </p>
      </section>

      <DocsPager
        previous={{ href: "/docs/browser", label: "Run in the browser" }}
        next={{ href: "/docs/api", label: "High-level API" }}
      />
    </DocsPage>
  );
}
