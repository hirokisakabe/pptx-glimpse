/**
 * Document path VRT は public snapshot を更新せず、current parser path の PNG を
 * 参照画像としてその場で比較する full parity harness。
 *
 * mismatchTolerance は #489 時点の実測 gap を blocker issue 単位で固定した上限。
 * blocker 解消 PR では対象ケースを再測定し、可能なら 0 まで締める。
 */

export const DOCUMENT_PATH_VRT_RENDER_WIDTH = 320;

export const DOCUMENT_PATH_VRT_SNAPSHOT_POLICY =
  "No committed snapshot update is required: document path VRT compares against the current parser path in-memory until the public default path changes.";

export const DOCUMENT_PATH_VRT_BLOCKER_ISSUES = {
  tables: 491,
  chartsAndSmartArt: 492,
  shapeTreeAndGeometry: 493,
  fillsAndBackgrounds: 494,
  textAndFonts: 496,
} as const;

export type DocumentPathVrtFixtureGroup = "shared" | "generated";

type DocumentPathVrtDiagnosticCode =
  | "cleandoc-adapter.raw-element-skipped"
  | "cleandoc-adapter.raw-fill-ignored";

interface DocumentPathVrtCase {
  readonly group: DocumentPathVrtFixtureGroup;
  readonly name: string;
  readonly fixture: string;
  readonly mismatchTolerance: number;
  readonly blockerIssues: readonly number[];
  readonly expectedDiagnosticCodes: readonly DocumentPathVrtDiagnosticCode[];
}

const B = DOCUMENT_PATH_VRT_BLOCKER_ISSUES;

export const DOCUMENT_PATH_VRT_SHARED_CASES = [
  shared("real-basic-theme", "real-basic-theme.pptx", 0, []),
  shared("real-product-page", "real-product-page.pptx", 0, []),
  shared("real-financial-report", "real-financial-report.pptx", 0, [
    B.tables,
    B.shapeTreeAndGeometry,
  ]),
  shared("sample", "sample.pptx", 0, []),
  shared("sample-issue-387", "sample-issue-387.pptx", 0, []),
] as const satisfies readonly DocumentPathVrtCase[];

export const DOCUMENT_PATH_VRT_GENERATED_CASES = [
  generated("shapes", "shapes.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("fill-and-lines", "fill-and-lines.pptx", 0, [], []),
  generated("text", "text.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("transform", "transform.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("background", "background.pptx", 0, [], []),
  generated("groups", "groups.pptx", 0, []),
  generated("charts", "charts.pptx", 0, [], []),
  generated("connectors", "connectors.pptx", 0, []),
  generated("custom-geometry", "custom-geometry.pptx", 0, []),
  generated("image", "image.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("tables", "tables.pptx", 0, [], []),
  generated("bullets", "bullets.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("flowchart", "flowchart.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("callouts-arcs", "callouts-arcs.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("arrows-stars", "arrows-stars.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("math-other", "math-other.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("word-wrap", "word-wrap.pptx", 0, [B.shapeTreeAndGeometry]),
  generated("background-blipfill", "background-blipfill.pptx", 0, [], []),
  generated("composite", "composite.pptx", 0, [], []),
  generated("text-decoration", "text-decoration.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("slide-size-4-3", "slide-size-4-3.pptx", 0, [], []),
  generated("effects", "effects.pptx", 0, [], []),
  generated("hyperlinks", "hyperlinks.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("pattern-image-fill", "pattern-image-fill.pptx", 0, [], []),
  generated("smartart", "smartart.pptx", 0, [], []),
  generated("theme-fonts", "theme-fonts.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated(
    "text-style-inheritance",
    "text-style-inheritance.pptx",
    0,
    [B.shapeTreeAndGeometry],
    [],
  ),
  generated("z-order-mixed", "z-order-mixed.pptx", 0, []),
  generated("paragraph-spacing", "paragraph-spacing.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("placeholder-overlap", "placeholder-overlap.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("image-crop", "image-crop.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("text-advanced", "text-advanced.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("shrink-to-fit", "shrink-to-fit.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("sp-autofit", "sp-autofit.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("style-reference", "style-reference.pptx", 0, [], []),
  generated("blip-effects", "blip-effects.pptx", 0, [], []),
  generated("image-stretch-tile", "image-stretch-tile.pptx", 0, [], []),
  generated("vertical-text", "vertical-text.pptx", 0, [B.shapeTreeAndGeometry]),
  generated(
    "shape-hyperlink-text-outline",
    "shape-hyperlink-text-outline.pptx",
    0,
    [B.shapeTreeAndGeometry],
    [],
  ),
  generated("charts-3d-fallback", "charts-3d-fallback.pptx", 0, [], []),
  generated("color-transforms", "color-transforms.pptx", 0, [B.shapeTreeAndGeometry], []),
  generated("table-complex-merge", "table-complex-merge.pptx", 0, [], []),
  generated("multi-lang-font", "multi-lang-font.pptx", 0, [B.shapeTreeAndGeometry]),
  generated(
    "placeholder-inheritance-extended",
    "placeholder-inheritance-extended.pptx",
    0,
    [B.shapeTreeAndGeometry],
    [],
  ),
  generated("table-style-border", "table-style-border.pptx", 0, [], []),
  generated(
    "placeholder-geometry-inheritance",
    "placeholder-geometry-inheritance.pptx",
    0,
    [B.shapeTreeAndGeometry],
    [],
  ),
  generated(
    "placeholder-empty-on-slide",
    "placeholder-empty-on-slide.pptx",
    0,
    [B.shapeTreeAndGeometry],
    [],
  ),
  generated(
    "interleaved-bullet-ppr",
    "interleaved-bullet-ppr.pptx",
    0,
    [B.shapeTreeAndGeometry],
    [],
  ),
] as const satisfies readonly DocumentPathVrtCase[];

export const DOCUMENT_PATH_VRT_CASES = [
  ...DOCUMENT_PATH_VRT_SHARED_CASES,
  ...DOCUMENT_PATH_VRT_GENERATED_CASES,
] as const satisfies readonly DocumentPathVrtCase[];

function shared(
  name: string,
  fixture: string,
  mismatchTolerance: number,
  blockerIssues: readonly number[],
  expectedDiagnosticCodes: readonly DocumentPathVrtDiagnosticCode[] = [],
): DocumentPathVrtCase {
  return {
    group: "shared",
    name,
    fixture,
    mismatchTolerance,
    blockerIssues,
    expectedDiagnosticCodes,
  };
}

function generated(
  name: string,
  fixture: string,
  mismatchTolerance: number,
  blockerIssues: readonly number[],
  expectedDiagnosticCodes: readonly DocumentPathVrtDiagnosticCode[] = [],
): DocumentPathVrtCase {
  return {
    group: "generated",
    name,
    fixture,
    mismatchTolerance,
    blockerIssues,
    expectedDiagnosticCodes,
  };
}
