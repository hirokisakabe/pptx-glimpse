import { existsSync, readFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { convertPptxToPng } from "../../packages/pptx-glimpse/src/converter.js";
import { compareImages } from "../compare-utils.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SNAPSHOT_DIR = join(__dirname, "snapshots");
const DIFF_DIR = join(__dirname, "diffs");
const FIXTURE_DIR = join(__dirname, "fixtures");

// LibreOffice と pptx-glimpse の比較は高い許容度が必要
// フォントレンダリングやアンチエイリアスの差異を許容する
const PIXEL_THRESHOLD = 0.3;
// 新規ケース用のフォールバック。CI で実測値を確認したら個別の tolerance を設定すること
const MISMATCH_TOLERANCE = 0.02;

// tolerance: CI 上の実測 mismatch 率 × 1.2 を 0.1pt 単位で切り上げ（下限 0.3%）
// 実測値はテスト実行時の [lo-vrt] ログで確認できる。
// LibreOffice や CI ランナーのフォント更新で実測値が変わったら、同じ方針で再設定する
const LO_VRT_CASES = [
  { name: "basic-shapes", fixture: "basic-shapes.pptx", tolerance: 0.012 },
  { name: "text-formatting", fixture: "text-formatting.pptx", tolerance: 0.018 },
  { name: "fill-and-lines", fixture: "fill-and-lines.pptx", tolerance: 0.003 },
  { name: "gradient-fills", fixture: "gradient-fills.pptx", tolerance: 0.545 },
  { name: "dash-lines", fixture: "dash-lines.pptx", tolerance: 0.018 },
  { name: "text-decoration", fixture: "text-decoration.pptx", tolerance: 0.013 },
  { name: "tables", fixture: "tables.pptx", tolerance: 0.143 },
  { name: "bullets", fixture: "bullets.pptx", tolerance: 0.01 },
  { name: "transforms", fixture: "transforms.pptx", tolerance: 0.057 },
  { name: "groups", fixture: "groups.pptx", tolerance: 0.158 },
  { name: "slide-background", fixture: "slide-background.pptx", tolerance: 0.005 },
  { name: "flowchart-shapes", fixture: "flowchart-shapes.pptx", tolerance: 0.031 },
  { name: "arrows-stars", fixture: "arrows-stars.pptx", tolerance: 0.159 },
  { name: "callouts-arcs", fixture: "callouts-arcs.pptx", tolerance: 0.168 },
  { name: "math-other", fixture: "math-other.pptx", tolerance: 0.093 },
  { name: "image", fixture: "image.pptx", tolerance: 0.003 },
  { name: "charts", fixture: "charts.pptx", tolerance: 0.145 },
  { name: "chart-legend-position", fixture: "chart-legend-position.pptx", tolerance: 0.207 },
  { name: "connectors", fixture: "connectors.pptx", tolerance: 0.003 },
  { name: "custom-geometry", fixture: "custom-geometry.pptx", tolerance: 0.386 },
  { name: "slide-size-4-3", fixture: "slide-size-4-3.pptx", tolerance: 0.006 },
  { name: "word-wrap", fixture: "word-wrap.pptx", tolerance: 0.031 },
  { name: "background-blipfill", fixture: "background-blipfill.pptx", tolerance: 0.691 },
  { name: "composite", fixture: "composite.pptx", tolerance: 0.128 },
  { name: "effects", fixture: "effects.pptx", tolerance: 0.011 },
  { name: "hyperlinks", fixture: "hyperlinks.pptx", tolerance: 0.009 },
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

          // tolerance 調整用に実測値を常に出力する
          console.log(
            `[lo-vrt] ${name} slide${result.slideNumber}: ` +
              `${(comparison.mismatchPercentage * 100).toFixed(3)}% ` +
              `(tolerance ${(tolerance * 100).toFixed(1)}%)`,
          );

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
