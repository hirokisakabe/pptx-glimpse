/**
 * VRT テストケース定義（共通）
 *
 * create-fixtures.ts と regression.test.ts の両方から参照される。
 * 新しいフィクスチャを追加する際はここに追記すること。
 */
export const VRT_CASES = [
  { name: "shapes", fixture: "shapes.pptx" },
  { name: "fill-and-lines", fixture: "fill-and-lines.pptx" },
  { name: "text", fixture: "text.pptx" },
  { name: "transform", fixture: "transform.pptx" },
  { name: "background", fixture: "background.pptx" },
  { name: "groups", fixture: "groups.pptx" },
  { name: "charts", fixture: "charts.pptx" },
  { name: "connectors", fixture: "connectors.pptx" },
  { name: "custom-geometry", fixture: "custom-geometry.pptx" },
  { name: "image", fixture: "image.pptx" },
  { name: "tables", fixture: "tables.pptx" },
  { name: "bullets", fixture: "bullets.pptx" },
  { name: "flowchart", fixture: "flowchart.pptx" },
  { name: "callouts-arcs", fixture: "callouts-arcs.pptx" },
  { name: "arrows-stars", fixture: "arrows-stars.pptx" },
  { name: "math-other", fixture: "math-other.pptx" },
  { name: "word-wrap", fixture: "word-wrap.pptx" },
  { name: "background-blipfill", fixture: "background-blipfill.pptx" },
  { name: "composite", fixture: "composite.pptx" },
  { name: "text-decoration", fixture: "text-decoration.pptx" },
  { name: "slide-size-4-3", fixture: "slide-size-4-3.pptx" },
  { name: "effects", fixture: "effects.pptx" },
  { name: "hyperlinks", fixture: "hyperlinks.pptx" },
  { name: "pattern-image-fill", fixture: "pattern-image-fill.pptx" },
  { name: "smartart", fixture: "smartart.pptx" },
  { name: "theme-fonts", fixture: "theme-fonts.pptx" },
  { name: "text-style-inheritance", fixture: "text-style-inheritance.pptx" },
  { name: "z-order-mixed", fixture: "z-order-mixed.pptx" },
  { name: "paragraph-spacing", fixture: "paragraph-spacing.pptx" },
  { name: "placeholder-overlap", fixture: "placeholder-overlap.pptx" },
  { name: "image-crop", fixture: "image-crop.pptx" },
  { name: "text-advanced", fixture: "text-advanced.pptx" },
  { name: "shrink-to-fit", fixture: "shrink-to-fit.pptx" },
  { name: "sp-autofit", fixture: "sp-autofit.pptx" },
  { name: "blip-effects", fixture: "blip-effects.pptx" },
  { name: "image-stretch-tile", fixture: "image-stretch-tile.pptx" },
] as const;

export type VrtCase = (typeof VRT_CASES)[number];
