import type { Metadata } from "next";
import { DocsCallout, DocsPage, DocsPager } from "@/components/docs/DocsPage";

export const metadata: Metadata = {
  title: "Render presentations",
  description: "Convert PPTX slides to SVG or PNG with pptx-glimpse.",
  alternates: { canonical: "/docs/rendering" },
};

const toc = [
  { href: "#svg-or-png", label: "Choose SVG or PNG" },
  { href: "#options", label: "Conversion options" },
  { href: "#fonts", label: "Fonts and text output" },
  { href: "#reports", label: "Reports and diagnostics" },
] as const;

export default function RenderingPage() {
  return (
    <DocsPage
      eyebrow="Guide / Rendering"
      title="Turn slides into output you can use."
      description="Render complete presentations or selected slides as embeddable SVG or PNG bytes. The conversion report also describes unsupported and fallback content."
      toc={toc}
    >
      <section id="svg-or-png">
        <h2>Choose SVG or PNG</h2>
        <div className="docs-comparison">
          <div>
            <code>convertPptxToSvg</code>
            <p>
              Best for browser previews, thumbnails, selectable text, and responsive interfaces.
            </p>
          </div>
          <div>
            <code>convertPptxToPng</code>
            <p>Best for image pipelines, exports, vision models, and fixed raster output.</p>
          </div>
        </div>
        <pre>
          <code>{`import { convertPptxToSvg } from "pptx-glimpse";

const { slides } = await convertPptxToSvg(pptx, {
  slides: [1, 3],
  textOutput: "text",
});

const firstSlideSvg = slides[0]?.svg;`}</code>
        </pre>
        <p>
          SVG is active document markup. Before inserting SVG from an untrusted presentation into
          the DOM, apply an SVG-aware sanitizer and a restrictive Content Security Policy. See the
          browser runtime guide for details.
        </p>
      </section>

      <section id="options">
        <h2>Conversion options</h2>
        <dl className="docs-definitions">
          <div>
            <dt>
              <code>slides</code>
            </dt>
            <dd>One-based slide numbers to render. Omit it to render the complete deck.</dd>
          </div>
          <div>
            <dt>
              <code>width</code>
            </dt>
            <dd>PNG output width in pixels. The default is 960.</dd>
          </div>
          <div>
            <dt>
              <code>fonts</code>
            </dt>
            <dd>Font bytes supplied by the application. This is the portable browser option.</dd>
          </div>
          <div>
            <dt>
              <code>fontDirs</code>
            </dt>
            <dd>Directories scanned for fonts in Node.js environments.</dd>
          </div>
          <div>
            <dt>
              <code>skipSystemFonts</code>
            </dt>
            <dd>Prevents scanning OS font directories for deterministic or browser-safe output.</dd>
          </div>
          <div>
            <dt>
              <code>fontMapping</code>
            </dt>
            <dd>Maps presentation font names to fonts available in the rendering environment.</dd>
          </div>
        </dl>
      </section>

      <section id="fonts">
        <h2>Fonts and text output</h2>
        <p>
          SVG defaults to path output for predictable rendering. Set <code>textOutput: "text"</code>{" "}
          to emit native SVG text with embedded subset fonts. Native text is selectable and often
          smaller for CJK-heavy slides, but the viewing environment can affect its rasterization.
        </p>
        <DocsCallout title="PNG always uses paths">
          <p>
            <code>convertPptxToPng</code> ignores SVG text mode and converts text to paths before
            rasterization.
          </p>
        </DocsCallout>
      </section>

      <section id="reports">
        <h2>Reports and diagnostics</h2>
        <pre>
          <code>{`const { slides, diagnostics, supportCoverage } =
  await convertPptxToSvg(pptx);

for (const diagnostic of diagnostics) {
  console.warn(diagnostic.code, diagnostic.message);
}

console.log(supportCoverage.overall);`}</code>
        </pre>
        <p>
          Support coverage counts input, output, skipped, unresolved, and fallback elements. It is
          not a pixel-accuracy score or a promise that PowerPoint would render the slide
          identically.
        </p>
      </section>

      <DocsPager
        previous={{ href: "/docs/getting-started", label: "Getting started" }}
        next={{ href: "/docs/editing", label: "Build an editing flow" }}
      />
    </DocsPage>
  );
}
