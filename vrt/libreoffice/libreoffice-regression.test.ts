import { readFileSync, existsSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { convertPptxToPng } from "../../src/converter.js";
import { compareImages } from "../compare-utils.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SNAPSHOT_DIR = join(__dirname, "snapshots");
const DIFF_DIR = join(__dirname, "diffs");
const FIXTURE_DIR = join(__dirname, "fixtures");

// LibreOffice と pptx-glimpse の比較は高い許容度が必要
// フォントレンダリングやアンチエイリアスの差異を許容する
const PIXEL_THRESHOLD = 0.3;
const MISMATCH_TOLERANCE = 0.05;

// tolerance: テストケース固有の許容度（省略時はデフォルト値を使用）
// 複雑な図形やチャートは LibreOffice と pptx-glimpse の描画差異が大きいため、高い許容度を設定
const LO_VRT_CASES = [
  { name: "basic-shapes", fixture: "lo-basic-shapes.pptx" },
  { name: "text-formatting", fixture: "lo-text-formatting.pptx" },
  { name: "fill-and-lines", fixture: "lo-fill-and-lines.pptx" },
  { name: "gradient-fills", fixture: "lo-gradient-fills.pptx" },
  { name: "dash-lines", fixture: "lo-dash-lines.pptx" },
  { name: "text-decoration", fixture: "lo-text-decoration.pptx" },
  { name: "tables", fixture: "lo-tables.pptx" },
  { name: "bullets", fixture: "lo-bullets.pptx" },
  { name: "transforms", fixture: "lo-transforms.pptx" },
  { name: "groups", fixture: "lo-groups.pptx" },
  { name: "slide-background", fixture: "lo-slide-background.pptx" },
  { name: "flowchart-shapes", fixture: "lo-flowchart-shapes.pptx" },
  { name: "arrows-stars", fixture: "lo-arrows-stars.pptx", tolerance: 0.15 },
  { name: "callouts-arcs", fixture: "lo-callouts-arcs.pptx", tolerance: 0.15 },
  { name: "math-other", fixture: "lo-math-other.pptx", tolerance: 0.1 },
  { name: "image", fixture: "lo-image.pptx" },
  { name: "charts", fixture: "lo-charts.pptx", tolerance: 0.15 },
  { name: "connectors", fixture: "lo-connectors.pptx" },
  { name: "custom-geometry", fixture: "lo-custom-geometry.pptx", tolerance: 0.45 },
] as const;

// フィクスチャとスナップショットの両方が存在する場合のみ実行
const hasFixtures =
  existsSync(FIXTURE_DIR) && LO_VRT_CASES.some((c) => existsSync(join(FIXTURE_DIR, c.fixture)));
const hasSnapshots =
  existsSync(SNAPSHOT_DIR) &&
  LO_VRT_CASES.some((c) => existsSync(join(SNAPSHOT_DIR, `${c.name}-slide1.png`)));

const describeOrSkip = hasFixtures && hasSnapshots ? describe : describe.skip;

describeOrSkip("LibreOffice Visual Regression Tests", { timeout: 60000 }, () => {
  for (const testCase of LO_VRT_CASES) {
    const { name, fixture } = testCase;
    const tolerance = "tolerance" in testCase ? testCase.tolerance : MISMATCH_TOLERANCE;
    // フィクスチャまたはスナップショットが存在しないケースはスキップ
    const fixturePath = join(FIXTURE_DIR, fixture);
    const snapshotPath = join(SNAPSHOT_DIR, `${name}-slide1.png`);
    const itOrSkip = existsSync(fixturePath) && existsSync(snapshotPath) ? it : it.skip;

    describe(name, () => {
      itOrSkip("should match LibreOffice reference", async () => {
        const input = readFileSync(fixturePath);
        const results = await convertPptxToPng(input);

        for (const result of results) {
          const refPath = join(SNAPSHOT_DIR, `${name}-slide${result.slideNumber}.png`);
          const diffPath = join(DIFF_DIR, `${name}-slide${result.slideNumber}-diff.png`);
          const comparison = await compareImages(result.png, refPath, diffPath, {
            pixelThreshold: PIXEL_THRESHOLD,
            mismatchTolerance: tolerance,
            resizeRef: true,
          });

          expect(
            comparison.passed,
            `${name} slide ${result.slideNumber}: ` +
              `${(comparison.mismatchPercentage * 100).toFixed(2)}% pixels differ ` +
              `(${comparison.mismatchedPixels}/${comparison.totalPixels}). ` +
              `Tolerance: ${tolerance * 100}%`,
          ).toBe(true);
        }
      });
    });
  }
});
