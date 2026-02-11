import { readFileSync, existsSync, mkdirSync, writeFileSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import pixelmatch from "pixelmatch";
import sharp from "sharp";
import { convertPptxToPng } from "../../src/converter.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SNAPSHOT_DIR = join(__dirname, "snapshots");
const DIFF_DIR = join(__dirname, "diffs");
const FIXTURE_DIR = join(__dirname, "../fixtures");

const PIXEL_THRESHOLD = 0.1;
const MISMATCH_TOLERANCE = 0.01;

interface CompareResult {
  totalPixels: number;
  mismatchedPixels: number;
  mismatchPercentage: number;
  passed: boolean;
}

async function compareImages(
  actualPng: Buffer,
  referencePath: string,
  diffPath: string,
): Promise<CompareResult> {
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

  const width = actualMeta.width ?? 0;
  const height = actualMeta.height ?? 0;

  if (width !== refMeta.width || height !== refMeta.height) {
    const total = width * height;
    return { totalPixels: total, mismatchedPixels: total, mismatchPercentage: 1, passed: false };
  }

  const [actualRaw, refRaw] = await Promise.all([
    sharp(actualPng).ensureAlpha().raw().toBuffer(),
    sharp(refPng).ensureAlpha().raw().toBuffer(),
  ]);

  const totalPixels = width * height;
  const diffBuf = new Uint8Array(totalPixels * 4);

  const mismatched = pixelmatch(actualRaw, refRaw, diffBuf, width, height, {
    threshold: PIXEL_THRESHOLD,
    includeAA: false,
  });

  const mismatchPercentage = mismatched / totalPixels;
  const passed = mismatchPercentage <= MISMATCH_TOLERANCE;

  if (!passed) {
    mkdirSync(dirname(diffPath), { recursive: true });
    const diffPng = await sharp(Buffer.from(diffBuf), { raw: { width, height, channels: 4 } })
      .png()
      .toBuffer();
    writeFileSync(diffPath, diffPng);
  }

  return { totalPixels, mismatchedPixels: mismatched, mismatchPercentage, passed };
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
  { name: "tables", fixture: "vrt-tables.pptx" },
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
          const comparison = await compareImages(result.png, refPath, diffPath);

          expect(
            comparison.passed,
            `${name} slide ${result.slideNumber}: ${(comparison.mismatchPercentage * 100).toFixed(2)}% pixels differ (${comparison.mismatchedPixels}/${comparison.totalPixels})`,
          ).toBe(true);
        }
      });
    });
  }
});
