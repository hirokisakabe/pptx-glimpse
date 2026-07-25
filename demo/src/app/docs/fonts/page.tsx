import type { Metadata } from "next";
import { DocsPage } from "@/components/docs/DocsPage";

export const metadata: Metadata = {
  title: "Using fonts",
  description: "Load, map, inspect, and render presentation fonts with pptx-glimpse.",
  alternates: { canonical: "/docs/fonts" },
};

const toc = [
  { href: "#loading", label: "Choose a loading strategy" },
  { href: "#browser", label: "Load font bytes" },
  { href: "#nodejs", label: "Load fonts in Node.js" },
  { href: "#mapping", label: "Map presentation fonts" },
  { href: "#inspect", label: "Inspect used fonts" },
  { href: "#text-output", label: "Choose SVG text output" },
] as const;

export default function FontsPage() {
  return (
    <DocsPage
      title="Using fonts"
      description="Make font availability explicit when presentation text must render consistently across browsers, servers, and containers."
      filePath="src/app/docs/fonts/page.tsx"
      toc={toc}
    >
      <section id="loading">
        <h2>Choose a loading strategy</h2>
        <p>pptx-glimpse can resolve fonts from one of three sources:</p>
        <ul>
          <li>
            Pass <code>fonts</code> when your application already has font bytes. This is the
            portable choice for browsers, edge runtimes, and deterministic servers.
          </li>
          <li>
            Pass <code>fontDirs</code> to scan application-controlled directories in Node.js.
          </li>
          <li>
            Set <code>skipSystemFonts: false</code> when a Node.js application should also use fonts
            installed by the operating system.
          </li>
        </ul>
        <p>
          When <code>fonts</code> is provided, directory and system font scanning are not used.
        </p>
      </section>

      <section id="browser">
        <h2>Load font bytes</h2>
        <pre>
          <code>{`import { convertPptxToSvg } from "pptx-glimpse";

const [pptx, inter] = await Promise.all([
  fetch("/slides/report.pptx").then((response) => response.arrayBuffer()),
  fetch("/fonts/Inter-Regular.ttf").then((response) => response.arrayBuffer()),
]);

const { slides } = await convertPptxToSvg(new Uint8Array(pptx), {
  fonts: [{ name: "Inter", data: inter }],
  fontMapping: {
    Arial: "Inter",
    Calibri: "Inter",
  },
});`}</code>
        </pre>
      </section>

      <section id="nodejs">
        <h2>Load fonts in Node.js</h2>
        <p>
          By default, high-level conversion can scan common system font directories on Linux, macOS,
          and Windows. Use explicit directories when running in a container or when output must not
          depend on the host:
        </p>
        <pre>
          <code>{`const { slides } = await convertPptxToSvg(pptx, {
  fontDirs: ["/app/fonts"],
  skipSystemFonts: true,
});`}</code>
        </pre>
      </section>

      <section id="mapping">
        <h2>Map presentation fonts</h2>
        <p>
          A PPTX can refer to fonts that are not available in the rendering environment. The
          built-in mapping covers common Microsoft fonts, including Calibri, Arial, Cambria, Meiryo,
          and Yu Gothic, with metrically compatible or broadly available alternatives. Override it
          when your application supplies its own font family:
        </p>
        <pre>
          <code>{`const { slides } = await convertPptxToSvg(pptx, {
  fontMapping: {
    "Custom Corp Font": "Inter",
    Arial: "Inter",
  },
});`}</code>
        </pre>
      </section>

      <section id="inspect">
        <h2>Inspect used fonts</h2>
        <p>
          <code>collectUsedFonts</code> reads the font names referenced by a presentation without
          performing a full render. Use it to decide which font files to load.
        </p>
        <pre>
          <code>{`import { collectUsedFonts } from "pptx-glimpse";

const { theme, fonts } = collectUsedFonts(pptx);
console.log(theme.majorFont, theme.minorFont);
console.log(fonts);`}</code>
        </pre>
      </section>

      <section id="text-output">
        <h2>Choose SVG text output</h2>
        <p>
          SVG conversion uses path outlines by default for predictable rendering. Set{" "}
          <code>textOutput: "text"</code> to emit selectable native SVG text with embedded subset
          fonts. PNG conversion always uses path output.
        </p>
        <p>
          Native text can be smaller and smoother in a browser, especially for CJK-heavy slides, but
          its rasterization depends on the viewer. See{" "}
          <a href="/docs/rendering#fonts">Rendering presentations</a> for the output tradeoffs.
        </p>
      </section>
    </DocsPage>
  );
}
