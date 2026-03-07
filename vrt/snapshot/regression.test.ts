import { existsSync, readFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { convertPptxToPng } from "../../src/converter.js";
import { compareImages } from "../compare-utils.js";
import { SHARED_FIXTURE_CASES, VRT_CASES } from "./vrt-cases.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SNAPSHOT_DIR = join(__dirname, "snapshots");
const DIFF_DIR = join(__dirname, "diffs");
const FIXTURE_DIR = join(__dirname, "fixtures");

const SHARED_FIXTURE_DIR = join(__dirname, "..", "..", "shared-fixtures");

const PIXEL_THRESHOLD = 0.1;
const MISMATCH_TOLERANCE = 0.015;

const SHARED_PIXEL_THRESHOLD = 0.3;
const SHARED_MISMATCH_TOLERANCE = 0.02;

describe("Visual Regression Tests", { timeout: 60000 }, () => {
  for (const { name, fixture } of VRT_CASES) {
    describe(name, () => {
      it("should match reference snapshot", async () => {
        const fixturePath = join(FIXTURE_DIR, fixture);
        if (!existsSync(fixturePath)) {
          throw new Error(
            `Fixture not found: ${fixturePath}. Run "npm run vrt:snapshot:fixtures" first.`,
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

describe("Visual Regression Tests (shared fixtures)", { timeout: 60000 }, () => {
  for (const { name, fixture } of SHARED_FIXTURE_CASES) {
    describe(name, () => {
      it("should match reference snapshot", async () => {
        const fixturePath = join(SHARED_FIXTURE_DIR, fixture);
        if (!existsSync(fixturePath)) {
          throw new Error(`Shared fixture not found: ${fixturePath}.`);
        }

        const input = readFileSync(fixturePath);
        const results = await convertPptxToPng(input);

        for (const result of results) {
          const refPath = join(SNAPSHOT_DIR, `${name}-slide${result.slideNumber}.png`);
          const diffPath = join(DIFF_DIR, `${name}-slide${result.slideNumber}-diff.png`);
          const comparison = await compareImages(result.png, refPath, diffPath, {
            pixelThreshold: SHARED_PIXEL_THRESHOLD,
            mismatchTolerance: SHARED_MISMATCH_TOLERANCE,
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
