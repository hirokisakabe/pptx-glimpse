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
  effects: 495,
  textAndFonts: 496,
} as const;

export type DocumentPathVrtFixtureGroup = "shared" | "generated";

type DocumentPathVrtDiagnosticCode =
  | "cleandoc-adapter.raw-element-skipped"
  | "cleandoc-adapter.raw-fill-ignored"
  | "document-render.cjk-font-context-unsupported";

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
  shared(
    "real-basic-theme",
    "real-basic-theme.pptx",
    0.02,
    [B.shapeTreeAndGeometry, B.textAndFonts],
    ["cleandoc-adapter.raw-element-skipped", "document-render.cjk-font-context-unsupported"],
  ),
  shared("real-product-page", "real-product-page.pptx", 0.04, [B.textAndFonts]),
  shared(
    "real-financial-report",
    "real-financial-report.pptx",
    0.16,
    [B.tables, B.chartsAndSmartArt, B.shapeTreeAndGeometry, B.textAndFonts],
    ["cleandoc-adapter.raw-element-skipped", "document-render.cjk-font-context-unsupported"],
  ),
  shared(
    "sample",
    "sample.pptx",
    0,
    [B.textAndFonts],
    ["document-render.cjk-font-context-unsupported"],
  ),
  shared("sample-issue-387", "sample-issue-387.pptx", 0, []),
] as const satisfies readonly DocumentPathVrtCase[];

export const DOCUMENT_PATH_VRT_GENERATED_CASES = [
  generated(
    "shapes",
    "shapes.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "fill-and-lines",
    "fill-and-lines.pptx",
    0.39,
    [B.shapeTreeAndGeometry, B.fillsAndBackgrounds],
    ["cleandoc-adapter.raw-element-skipped", "cleandoc-adapter.raw-fill-ignored"],
  ),
  generated(
    "text",
    "text.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "transform",
    "transform.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "background",
    "background.pptx",
    0.93,
    [B.shapeTreeAndGeometry, B.fillsAndBackgrounds],
    ["cleandoc-adapter.raw-element-skipped", "cleandoc-adapter.raw-fill-ignored"],
  ),
  generated(
    "groups",
    "groups.pptx",
    0.46,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "charts",
    "charts.pptx",
    0.64,
    [B.chartsAndSmartArt],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "connectors",
    "connectors.pptx",
    0.02,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "custom-geometry",
    "custom-geometry.pptx",
    0.39,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "image",
    "image.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated("tables", "tables.pptx", 0.21, [B.tables], ["cleandoc-adapter.raw-element-skipped"]),
  generated(
    "bullets",
    "bullets.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "flowchart",
    "flowchart.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "callouts-arcs",
    "callouts-arcs.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "arrows-stars",
    "arrows-stars.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "math-other",
    "math-other.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "word-wrap",
    "word-wrap.pptx",
    0,
    [B.shapeTreeAndGeometry, B.textAndFonts],
    ["cleandoc-adapter.raw-element-skipped", "document-render.cjk-font-context-unsupported"],
  ),
  generated(
    "background-blipfill",
    "background-blipfill.pptx",
    0.98,
    [B.shapeTreeAndGeometry, B.fillsAndBackgrounds],
    ["cleandoc-adapter.raw-element-skipped", "cleandoc-adapter.raw-fill-ignored"],
  ),
  generated(
    "composite",
    "composite.pptx",
    0.24,
    [B.shapeTreeAndGeometry, B.fillsAndBackgrounds],
    ["cleandoc-adapter.raw-element-skipped", "cleandoc-adapter.raw-fill-ignored"],
  ),
  generated(
    "text-decoration",
    "text-decoration.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "slide-size-4-3",
    "slide-size-4-3.pptx",
    0.99,
    [B.shapeTreeAndGeometry, B.fillsAndBackgrounds],
    ["cleandoc-adapter.raw-element-skipped", "cleandoc-adapter.raw-fill-ignored"],
  ),
  generated("effects", "effects.pptx", 0.19, [B.effects], ["cleandoc-adapter.raw-element-skipped"]),
  generated(
    "hyperlinks",
    "hyperlinks.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "pattern-image-fill",
    "pattern-image-fill.pptx",
    0.56,
    [B.shapeTreeAndGeometry, B.fillsAndBackgrounds],
    ["cleandoc-adapter.raw-element-skipped", "cleandoc-adapter.raw-fill-ignored"],
  ),
  generated(
    "smartart",
    "smartart.pptx",
    0.51,
    [B.chartsAndSmartArt],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "theme-fonts",
    "theme-fonts.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "text-style-inheritance",
    "text-style-inheritance.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "z-order-mixed",
    "z-order-mixed.pptx",
    0.2,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "paragraph-spacing",
    "paragraph-spacing.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "placeholder-overlap",
    "placeholder-overlap.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "image-crop",
    "image-crop.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "text-advanced",
    "text-advanced.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "shrink-to-fit",
    "shrink-to-fit.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "sp-autofit",
    "sp-autofit.pptx",
    0.01,
    [B.shapeTreeAndGeometry, B.textAndFonts],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "style-reference",
    "style-reference.pptx",
    0.35,
    [B.shapeTreeAndGeometry, B.fillsAndBackgrounds],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "blip-effects",
    "blip-effects.pptx",
    0.65,
    [B.effects],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "image-stretch-tile",
    "image-stretch-tile.pptx",
    0.34,
    [B.shapeTreeAndGeometry, B.fillsAndBackgrounds],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "vertical-text",
    "vertical-text.pptx",
    0,
    [B.shapeTreeAndGeometry, B.textAndFonts],
    ["cleandoc-adapter.raw-element-skipped", "document-render.cjk-font-context-unsupported"],
  ),
  generated(
    "shape-hyperlink-text-outline",
    "shape-hyperlink-text-outline.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "charts-3d-fallback",
    "charts-3d-fallback.pptx",
    0.54,
    [B.chartsAndSmartArt],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "color-transforms",
    "color-transforms.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "table-complex-merge",
    "table-complex-merge.pptx",
    0.31,
    [B.tables],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "multi-lang-font",
    "multi-lang-font.pptx",
    0,
    [B.shapeTreeAndGeometry, B.textAndFonts],
    ["cleandoc-adapter.raw-element-skipped", "document-render.cjk-font-context-unsupported"],
  ),
  generated(
    "placeholder-inheritance-extended",
    "placeholder-inheritance-extended.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "table-style-border",
    "table-style-border.pptx",
    0.14,
    [B.tables],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "placeholder-geometry-inheritance",
    "placeholder-geometry-inheritance.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "placeholder-empty-on-slide",
    "placeholder-empty-on-slide.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
  ),
  generated(
    "interleaved-bullet-ppr",
    "interleaved-bullet-ppr.pptx",
    0,
    [B.shapeTreeAndGeometry],
    ["cleandoc-adapter.raw-element-skipped"],
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
