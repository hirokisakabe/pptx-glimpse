import { readFileSync, existsSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import sharp from "sharp";
import { convertPptxToPng } from "../../src/converter.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SNAPSHOT_DIR = join(__dirname, "snapshots");
const FIXTURE_DIR = join(__dirname, "../fixtures");

const PIXEL_THRESHOLD = 5;
const MISMATCH_TOLERANCE = 0.005;

interface CompareResult {
  totalPixels: number;
  mismatchedPixels: number;
  mismatchPercentage: number;
  passed: boolean;
}

async function compareImages(actualPng: Buffer, referencePath: string): Promise<CompareResult> {
  if (!existsSync(referencePath)) {
    throw new Error(
      `Reference snapshot not found: ${referencePath}. Run "npm run test:vrt:update" to generate.`,
    );
  }

  const refPng = readFileSync(referencePath);

  const [actualMeta, refMeta] = await Promise.all([
    sharp(actualPng).metadata(),
    sharp(refPng).metadata(),
  ]);

  if (actualMeta.width !== refMeta.width || actualMeta.height !== refMeta.height) {
    const total = (actualMeta.width ?? 0) * (actualMeta.height ?? 0);
    return { totalPixels: total, mismatchedPixels: total, mismatchPercentage: 1, passed: false };
  }

  const [actualRaw, refRaw] = await Promise.all([
    sharp(actualPng).raw().toBuffer(),
    sharp(refPng).raw().toBuffer(),
  ]);

  const channels = actualMeta.channels ?? 4;
  const totalPixels = actualRaw.length / channels;
  let mismatched = 0;

  for (let i = 0; i < actualRaw.length; i += channels) {
    let pixelDiff = false;
    for (let c = 0; c < channels; c++) {
      if (Math.abs(actualRaw[i + c] - refRaw[i + c]) > PIXEL_THRESHOLD) {
        pixelDiff = true;
        break;
      }
    }
    if (pixelDiff) mismatched++;
  }

  const mismatchPercentage = mismatched / totalPixels;

  return {
    totalPixels,
    mismatchedPixels: mismatched,
    mismatchPercentage,
    passed: mismatchPercentage <= MISMATCH_TOLERANCE,
  };
}

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
          const comparison = await compareImages(result.png, refPath);

          expect(
            comparison.passed,
            `${name} slide ${result.slideNumber}: ${(comparison.mismatchPercentage * 100).toFixed(2)}% pixels differ (${comparison.mismatchedPixels}/${comparison.totalPixels})`,
          ).toBe(true);
        }
      });
    });
  }
});
