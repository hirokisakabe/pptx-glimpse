import type { Metadata } from "next";
import { DocsPage, DocsPager } from "@/components/docs/DocsPage";

export const metadata: Metadata = {
  title: "Feature support",
  description: "Review the PowerPoint content currently rendered by pptx-glimpse.",
  alternates: { canonical: "/docs/feature-support" },
};

const toc = [
  { href: "#scope", label: "How to read this page" },
  { href: "#content", label: "Slide content" },
  { href: "#appearance", label: "Appearance and layout" },
  { href: "#not-supported", label: "Not supported" },
] as const;

const contentRows = [
  ["Shapes", "136 preset shapes, custom paths, connectors, groups, rotation, and flipping"],
  [
    "Text",
    "Character and paragraph formatting, bullets, wrapping, autofit, inheritance, tabs, and fields",
  ],
  ["Images", "PNG, JPEG, and GIF elements and fills, including cropping and transforms"],
  ["Tables", "Rows, columns, merged cells, text, fills, and cell borders"],
  [
    "Charts",
    "Bar, line, pie, scatter, area, doughnut, bubble, and radar charts with common chart elements",
  ],
  ["SmartArt", "PowerPoint pre-rendered drawing shapes and AlternateContent fallbacks"],
] as const;

const appearanceRows = [
  ["Fills", "Solid, linear and radial gradient, image, pattern, group, and no-fill"],
  ["Lines", "Width, color, transparency, caps, joins, dashes, and arrow endpoints"],
  ["Colors", "RGB, theme, and system colors with common transforms and transparency"],
  ["Effects", "Outer and inner shadows, glow, and soft edges"],
  ["Backgrounds", "Solid, gradient, and image backgrounds with slide → layout → master fallback"],
  ["Slide settings", "16:9, 4:3, and custom sizes, themes, and master-shape visibility"],
] as const;

const unsupportedRows = [
  ["Effects", "Reflection, 3D rotation and extrusion, and artistic effects"],
  ["Charts", "Stock, combo, histogram, box plot, waterfall, treemap, and sunburst charts"],
  ["Chart details", "Data labels, axis titles, tick marks, grid lines, error bars, and trendlines"],
  ["Text", "Per-run shadow or glow effects and text columns"],
  ["Tables", "Table style template application and diagonal borders"],
  ["Shapes", "Union, subtract, intersect, and fragment operations"],
  ["Media", "Embedded video and audio"],
  ["Motion", "Object animations and slide transitions"],
  ["Slide metadata", "Notes, comments, headers, footers, slide numbers, and dates"],
  ["Image formats", "EMF and WMF are parsed but not rendered"],
  ["Other", "Macros, VBA, sections, and zoom slides"],
] as const;

function FeatureTable({
  rows,
}: {
  readonly rows: readonly (readonly [category: string, details: string])[];
}) {
  return (
    <div className="docs-feature-table" role="table">
      {rows.map(([category, details]) => (
        <div role="row" key={category}>
          <strong role="rowheader">{category}</strong>
          <span role="cell">{details}</span>
        </div>
      ))}
    </div>
  );
}

export default function FeatureSupportPage() {
  return (
    <DocsPage
      title="Feature support"
      description="pptx-glimpse covers common static presentation content. Check conversion diagnostics for the specific files your application handles."
      filePath="src/app/docs/feature-support/page.tsx"
      toc={toc}
    >
      <section id="scope">
        <h2>How to read this page</h2>
        <p>
          This is a rendering overview, not a promise of pixel-identical PowerPoint output.
          Individual presentations can still depend on unavailable fonts, uncommon OOXML variants,
          or fallback behavior. Use <code>diagnostics</code> and <code>supportCoverage</code> from
          the conversion result to inspect each file.
        </p>
        <p>
          The lower-level document package has a separate, evidence-linked{" "}
          <a href="https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/feature-support.md">
            document API support matrix
          </a>
          .
        </p>
      </section>

      <section id="content">
        <h2>Slide content</h2>
        <FeatureTable rows={contentRows} />
      </section>

      <section id="appearance">
        <h2>Appearance and layout</h2>
        <FeatureTable rows={appearanceRows} />
      </section>

      <section id="not-supported">
        <h2>Not supported</h2>
        <FeatureTable rows={unsupportedRows} />
      </section>

      <DocsPager
        previous={{ href: "/docs/nodejs", label: "Run in Node.js" }}
        next={{ href: "/docs/packages", label: "Choose a package" }}
      />
    </DocsPage>
  );
}
