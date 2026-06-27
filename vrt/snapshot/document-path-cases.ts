/**
 * Document path VRT は public snapshot を更新せず、明示的な parser path の PNG を
 * 参照画像としてその場で比較する full parity harness。
 *
 * CleanDoc default switch 後も parser path oracle との zero-diff gate として残す。
 * 以前の migration blocker issue はすべて解消済みなので、ケースごとの issue 番号は
 * ここでは保持しない。
 */

export const DOCUMENT_PATH_VRT_RENDER_WIDTH = 320;

export const DOCUMENT_PATH_VRT_SNAPSHOT_POLICY =
  "No committed snapshot update is required: document path VRT compares against the explicit parser path oracle in-memory.";

export type DocumentPathVrtFixtureGroup = "shared" | "generated";

interface DocumentPathVrtCase {
  readonly group: DocumentPathVrtFixtureGroup;
  readonly name: string;
  readonly fixture: string;
  readonly mismatchTolerance: number;
}

export const DOCUMENT_PATH_VRT_SHARED_CASES = [
  shared("real-basic-theme", "real-basic-theme.pptx", 0),
  shared("real-product-page", "real-product-page.pptx", 0),
  shared("real-financial-report", "real-financial-report.pptx", 0),
  shared("sample", "sample.pptx", 0),
  shared("sample-issue-387", "sample-issue-387.pptx", 0),
] as const satisfies readonly DocumentPathVrtCase[];

export const DOCUMENT_PATH_VRT_GENERATED_CASES = [
  generated("shapes", "shapes.pptx", 0),
  generated("fill-and-lines", "fill-and-lines.pptx", 0),
  generated("text", "text.pptx", 0),
  generated("transform", "transform.pptx", 0),
  generated("background", "background.pptx", 0),
  generated("groups", "groups.pptx", 0),
  generated("charts", "charts.pptx", 0),
  generated("connectors", "connectors.pptx", 0),
  generated("custom-geometry", "custom-geometry.pptx", 0),
  generated("image", "image.pptx", 0),
  generated("tables", "tables.pptx", 0),
  generated("bullets", "bullets.pptx", 0),
  generated("flowchart", "flowchart.pptx", 0),
  generated("callouts-arcs", "callouts-arcs.pptx", 0),
  generated("arrows-stars", "arrows-stars.pptx", 0),
  generated("math-other", "math-other.pptx", 0),
  generated("word-wrap", "word-wrap.pptx", 0),
  generated("background-blipfill", "background-blipfill.pptx", 0),
  generated("composite", "composite.pptx", 0),
  generated("text-decoration", "text-decoration.pptx", 0),
  generated("slide-size-4-3", "slide-size-4-3.pptx", 0),
  generated("effects", "effects.pptx", 0),
  generated("hyperlinks", "hyperlinks.pptx", 0),
  generated("pattern-image-fill", "pattern-image-fill.pptx", 0),
  generated("smartart", "smartart.pptx", 0),
  generated("theme-fonts", "theme-fonts.pptx", 0),
  generated("text-style-inheritance", "text-style-inheritance.pptx", 0),
  generated("z-order-mixed", "z-order-mixed.pptx", 0),
  generated("paragraph-spacing", "paragraph-spacing.pptx", 0),
  generated("placeholder-overlap", "placeholder-overlap.pptx", 0),
  generated("image-crop", "image-crop.pptx", 0),
  generated("text-advanced", "text-advanced.pptx", 0),
  generated("shrink-to-fit", "shrink-to-fit.pptx", 0),
  generated("sp-autofit", "sp-autofit.pptx", 0),
  generated("style-reference", "style-reference.pptx", 0),
  generated("blip-effects", "blip-effects.pptx", 0),
  generated("image-stretch-tile", "image-stretch-tile.pptx", 0),
  generated("vertical-text", "vertical-text.pptx", 0),
  generated("shape-hyperlink-text-outline", "shape-hyperlink-text-outline.pptx", 0),
  generated("charts-3d-fallback", "charts-3d-fallback.pptx", 0),
  generated("color-transforms", "color-transforms.pptx", 0),
  generated("table-complex-merge", "table-complex-merge.pptx", 0),
  generated("multi-lang-font", "multi-lang-font.pptx", 0),
  generated("placeholder-inheritance-extended", "placeholder-inheritance-extended.pptx", 0),
  generated("table-style-border", "table-style-border.pptx", 0),
  generated("placeholder-geometry-inheritance", "placeholder-geometry-inheritance.pptx", 0),
  generated("placeholder-empty-on-slide", "placeholder-empty-on-slide.pptx", 0),
  generated("interleaved-bullet-ppr", "interleaved-bullet-ppr.pptx", 0),
] as const satisfies readonly DocumentPathVrtCase[];

export const DOCUMENT_PATH_VRT_CASES = [
  ...DOCUMENT_PATH_VRT_SHARED_CASES,
  ...DOCUMENT_PATH_VRT_GENERATED_CASES,
] as const satisfies readonly DocumentPathVrtCase[];

function shared(name: string, fixture: string, mismatchTolerance: number): DocumentPathVrtCase {
  return {
    group: "shared",
    name,
    fixture,
    mismatchTolerance,
  };
}

function generated(name: string, fixture: string, mismatchTolerance: number): DocumentPathVrtCase {
  return {
    group: "generated",
    name,
    fixture,
    mismatchTolerance,
  };
}
