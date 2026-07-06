/**
 * VRT test case definition (common)
 *
 * Referenced by both create-fixtures.ts and regression.test.ts.
 * When adding new fixtures, add them here.
 */

/**
 * A VRT case using PPTX created by an actual application in shared-fixtures/.
 * The fixtures are already committed, so there is no need to create them using create-fixtures.ts.
 */
export const SHARED_FIXTURE_CASES = [
  { name: "real-basic-theme", fixture: "real-basic-theme.pptx" },
  { name: "real-product-page", fixture: "real-product-page.pptx" },
  { name: "real-financial-report", fixture: "real-financial-report.pptx" },
  { name: "sample", fixture: "sample.pptx" },
  { name: "sample-issue-387", fixture: "sample-issue-387.pptx" },
] as const;

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
  { name: "style-reference", fixture: "style-reference.pptx" },
  { name: "blip-effects", fixture: "blip-effects.pptx" },
  { name: "image-stretch-tile", fixture: "image-stretch-tile.pptx" },
  { name: "vertical-text", fixture: "vertical-text.pptx" },
  { name: "shape-hyperlink-text-outline", fixture: "shape-hyperlink-text-outline.pptx" },
  { name: "charts-3d-fallback", fixture: "charts-3d-fallback.pptx" },
  { name: "color-transforms", fixture: "color-transforms.pptx" },
  { name: "table-complex-merge", fixture: "table-complex-merge.pptx" },
  { name: "multi-lang-font", fixture: "multi-lang-font.pptx" },
  { name: "placeholder-inheritance-extended", fixture: "placeholder-inheritance-extended.pptx" },
  { name: "table-style-border", fixture: "table-style-border.pptx" },
  {
    name: "placeholder-geometry-inheritance",
    fixture: "placeholder-geometry-inheritance.pptx",
  },
  { name: "placeholder-empty-on-slide", fixture: "placeholder-empty-on-slide.pptx" },
  { name: "interleaved-bullet-ppr", fixture: "interleaved-bullet-ppr.pptx" },
] as const;

const SNAPSHOT_CASES = [...VRT_CASES, ...SHARED_FIXTURE_CASES] as const;

type SnapshotCase = (typeof SNAPSHOT_CASES)[number];

function formatCaseNameList(caseNames: readonly string[]): string {
  return caseNames.map((name) => `"${name}"`).join(", ");
}

export function resolveSnapshotCases(caseNames: readonly string[]): SnapshotCase[] {
  const casesByName = new Map<string, SnapshotCase>(
    SNAPSHOT_CASES.map((vrtCase) => [vrtCase.name, vrtCase] as const),
  );
  const unknownNames = caseNames.filter((caseName) => !casesByName.has(caseName));

  if (unknownNames.length > 0) {
    throw new Error(
      `Unknown VRT snapshot case name(s): ${formatCaseNameList(unknownNames)}. Check vrt/snapshot/vrt-cases.ts for valid case names.`,
    );
  }

  const selectedCases: SnapshotCase[] = [];
  const seenNames = new Set<string>();

  for (const caseName of caseNames) {
    if (seenNames.has(caseName)) {
      continue;
    }
    seenNames.add(caseName);
    const vrtCase = casesByName.get(caseName);
    if (vrtCase !== undefined) {
      selectedCases.push(vrtCase);
    }
  }

  return selectedCases;
}

export function resolveGeneratedVrtCases(
  caseNames: readonly string[],
): (typeof VRT_CASES)[number][] {
  if (caseNames.length === 0) {
    return [...VRT_CASES];
  }

  const selectedNames = new Set(resolveSnapshotCases(caseNames).map(({ name }) => name));
  return VRT_CASES.filter(({ name }) => selectedNames.has(name));
}
