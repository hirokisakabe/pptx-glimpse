import type { Metadata } from "next";
import { DocsCallout, DocsPage, DocsPager } from "@/components/docs/DocsPage";

export const metadata: Metadata = {
  title: "Run in the browser",
  description: "Load PPTX and font bytes, render SVG, edit, and download files in a browser.",
  alternates: { canonical: "/docs/browser" },
};

const toc = [
  { href: "#load-bytes", label: "Load bytes" },
  { href: "#fonts", label: "Provide fonts" },
  { href: "#svg", label: "Embed SVG safely" },
  { href: "#png", label: "Initialize PNG rendering" },
] as const;

export default function BrowserPage() {
  return (
    <DocsPage
      eyebrow="Runtime / Browser"
      title="Keep the presentation in the browser."
      description="Browser bundles use the same Uint8Array-oriented APIs. Your application supplies file and font bytes and decides how rendered or edited output enters the page."
      toc={toc}
    >
      <section id="load-bytes">
        <h2>Load presentation bytes</h2>
        <pre>
          <code>{`const file = input.files?.[0];
if (!file) throw new Error("Choose a PPTX file");

const pptx = new Uint8Array(await file.arrayBuffer());`}</code>
        </pre>
        <DocsCallout title="Local by default">
          <p>
            Reading a <code>File</code> and passing its bytes to pptx-glimpse does not upload it.
            Network behavior is entirely controlled by your application.
          </p>
        </DocsCallout>
      </section>

      <section id="fonts">
        <h2>Provide fonts</h2>
        <p>
          Browsers cannot scan OS font directories. Fetch or bundle the fonts your application is
          allowed to use and provide their bytes through the <code>fonts</code> option.
        </p>
        <pre>
          <code>{`import { convertPptxToSvg } from "pptx-glimpse";

const fontResponse = await fetch("/fonts/Inter-Regular.ttf");
if (!fontResponse.ok) {
  throw new Error(\`Failed to load font: HTTP \${fontResponse.status}\`);
}
const inter = await fontResponse.arrayBuffer();

const report = await convertPptxToSvg(pptx, {
  fonts: [{ name: "Inter", data: inter }],
  textOutput: "text",
});`}</code>
        </pre>
      </section>

      <section id="svg">
        <h2>Embed SVG safely</h2>
        <p>
          SVG output is a string. Insert it only where your application intentionally accepts
          rendered presentation content. Review your sanitization and Content Security Policy when
          presentations can come from untrusted users.
        </p>
        <p>
          Native <code>&lt;text&gt;</code> output works best as inline SVG. Embedded font rules may
          not survive sanitizers or render when the SVG is loaded through an{" "}
          <code>&lt;img&gt;</code> element.
        </p>
      </section>

      <section id="png">
        <h2>Initialize PNG rendering</h2>
        <p>
          Browser PNG conversion requires explicit resvg WASM initialization. SVG rendering and the
          high-level editor session do not.
        </p>
        <pre>
          <code>{`import { initResvgWasm, convertPptxToPng } from "pptx-glimpse";

await initResvgWasm(await fetch("/resvg.wasm"));
const report = await convertPptxToPng(pptx);`}</code>
        </pre>
      </section>

      <DocsPager
        previous={{ href: "/docs/editing", label: "Build an editing flow" }}
        next={{ href: "/docs/nodejs", label: "Run in Node.js" }}
      />
    </DocsPage>
  );
}
