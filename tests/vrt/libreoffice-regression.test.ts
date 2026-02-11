import { readFileSync, existsSync, mkdirSync, writeFileSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import pixelmatch from "pixelmatch";
import sharp from "sharp";
import { convertPptxToPng } from "../../src/converter.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SNAPSHOT_DIR = join(__dirname, "libreoffice-snapshots");
const DIFF_DIR = join(__dirname, "libreoffice-diffs");
const FIXTURE_DIR = join(__dirname, "libreoffice-fixtures");

// LibreOffice と pptx-glimpse の比較は高い許容度が必要
// フォントレンダリングやアンチエイリアスの差異を許容する
const PIXEL_THRESHOLD = 0.3;
const MISMATCH_TOLERANCE = 0.05;

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
      `LibreOffice reference snapshot not found: ${referencePath}. ` +
        `Run "npm run vrt:lo:update" to generate.`,
    );
  }

  const refPng = readFileSync(referencePath);

  const actualMeta = await sharp(actualPng).metadata();
  const width = actualMeta.width ?? 0;
  const height = actualMeta.height ?? 0;

  // 参照画像を pptx-glimpse の出力サイズにリサイズして比較
  const [actualRaw, refRaw] = await Promise.all([
    sharp(actualPng).ensureAlpha().raw().toBuffer(),
    sharp(refPng).resize(width, height, { fit: "fill" }).ensureAlpha().raw().toBuffer(),
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

const LO_VRT_CASES = [
  { name: "basic-shapes", fixture: "lo-basic-shapes.pptx" },
  { name: "text-formatting", fixture: "lo-text-formatting.pptx" },
  { name: "fill-and-lines", fixture: "lo-fill-and-lines.pptx" },
] as const;

// フィクスチャとスナップショットの両方が存在する場合のみ実行
const hasFixtures =
  existsSync(FIXTURE_DIR) && LO_VRT_CASES.some((c) => existsSync(join(FIXTURE_DIR, c.fixture)));
const hasSnapshots =
  existsSync(SNAPSHOT_DIR) &&
  LO_VRT_CASES.some((c) => existsSync(join(SNAPSHOT_DIR, `${c.name}-slide1.png`)));

const describeOrSkip = hasFixtures && hasSnapshots ? describe : describe.skip;

describeOrSkip("LibreOffice Visual Regression Tests", { timeout: 60000 }, () => {
  for (const { name, fixture } of LO_VRT_CASES) {
    describe(name, () => {
      it("should match LibreOffice reference", async () => {
        const fixturePath = join(FIXTURE_DIR, fixture);
        if (!existsSync(fixturePath)) {
          throw new Error(`Fixture not found: ${fixturePath}`);
        }

        const input = readFileSync(fixturePath);
        const results = await convertPptxToPng(input);

        for (const result of results) {
          const refPath = join(SNAPSHOT_DIR, `${name}-slide${result.slideNumber}.png`);
          const diffPath = join(DIFF_DIR, `${name}-slide${result.slideNumber}-diff.png`);
          const comparison = await compareImages(result.png, refPath, diffPath);

          expect(
            comparison.passed,
            `${name} slide ${result.slideNumber}: ` +
              `${(comparison.mismatchPercentage * 100).toFixed(2)}% pixels differ ` +
              `(${comparison.mismatchedPixels}/${comparison.totalPixels}). ` +
              `Tolerance: ${MISMATCH_TOLERANCE * 100}%`,
          ).toBe(true);
        }
      });
    });
  }
});
