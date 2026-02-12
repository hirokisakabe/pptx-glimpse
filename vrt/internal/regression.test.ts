import { readFileSync, existsSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { convertPptxToPng } from "../../src/converter.js";
import { compareImages } from "../compare-utils.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SNAPSHOT_DIR = join(__dirname, "snapshots");
const DIFF_DIR = join(__dirname, "diffs");
const FIXTURE_DIR = join(__dirname, "fixtures");

const PIXEL_THRESHOLD = 0.1;
const MISMATCH_TOLERANCE = 0.015;

const VRT_CASES = [
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
] as const;

describe("Visual Regression Tests", { timeout: 60000 }, () => {
  for (const { name, fixture } of VRT_CASES) {
    describe(name, () => {
      it("should match reference snapshot", async () => {
        const fixturePath = join(FIXTURE_DIR, fixture);
        if (!existsSync(fixturePath)) {
          throw new Error(
            `Fixture not found: ${fixturePath}. Run "npm run vrt:internal:fixtures" first.`,
          );
        }

        const input = readFileSync(fixturePath);
        const results = await convertPptxToPng(input);

        for (const result of results) {
          const refPath = join(SNAPSHOT_DIR, `${name}-slide${result.slideNumber}.png`);
          const diffPath = join(DIFF_DIR, `${name}-slide${result.slideNumber}-diff.png`);
          const comparison = await compareImages(result.png, refPath, diffPath, {
            pixelThreshold: PIXEL_THRESHOLD,
            mismatchTolerance: MISMATCH_TOLERANCE,
          });

          expect(
            comparison.passed,
            `${name} slide ${result.slideNumber}: ${(comparison.mismatchPercentage * 100).toFixed(2)}% pixels differ (${comparison.mismatchedPixels}/${comparison.totalPixels})`,
          ).toBe(true);
        }
      });
    });
  }
});
