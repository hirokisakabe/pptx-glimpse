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
  { name: "shapes", fixture: "vrt-shapes.pptx" },
  { name: "fill-and-lines", fixture: "vrt-fill-and-lines.pptx" },
  { name: "text", fixture: "vrt-text.pptx" },
  { name: "transform", fixture: "vrt-transform.pptx" },
  { name: "background", fixture: "vrt-background.pptx" },
  { name: "groups", fixture: "vrt-groups.pptx" },
  { name: "charts", fixture: "vrt-charts.pptx" },
  { name: "connectors", fixture: "vrt-connectors.pptx" },
  { name: "custom-geometry", fixture: "vrt-custom-geometry.pptx" },
  { name: "image", fixture: "vrt-image.pptx" },
  { name: "tables", fixture: "vrt-tables.pptx" },
  { name: "bullets", fixture: "vrt-bullets.pptx" },
  { name: "flowchart", fixture: "vrt-flowchart.pptx" },
  { name: "callouts-arcs", fixture: "vrt-callouts-arcs.pptx" },
  { name: "arrows-stars", fixture: "vrt-arrows-stars.pptx" },
  { name: "math-other", fixture: "vrt-math-other.pptx" },
  { name: "word-wrap", fixture: "vrt-word-wrap.pptx" },
  { name: "background-blipfill", fixture: "vrt-background-blipfill.pptx" },
  { name: "composite", fixture: "vrt-composite.pptx" },
  { name: "text-decoration", fixture: "vrt-text-decoration.pptx" },
  { name: "slide-size-4-3", fixture: "vrt-slide-size-4-3.pptx" },
  { name: "effects", fixture: "vrt-effects.pptx" },
] as const;

describe("Visual Regression Tests", { timeout: 60000 }, () => {
  for (const { name, fixture } of VRT_CASES) {
    describe(name, () => {
      it("should match reference snapshot", async () => {
        const fixturePath = join(FIXTURE_DIR, fixture);
        if (!existsSync(fixturePath)) {
          throw new Error(
            `Fixture not found: ${fixturePath}. Run "npm run test:vrt:fixtures" first.`,
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
