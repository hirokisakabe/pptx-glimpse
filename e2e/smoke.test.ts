import { readFileSync } from "fs";
import { join } from "path";
import { beforeAll, describe, expect, it, vi } from "vitest";

// macOS 上では `/System/Library/Fonts` 等から数百個のフォントを parse すると
// worker メモリが膨れ上がり OOM する。ここでは構造ベースのアサーション
// (`<text|<path>` のいずれかが出ること、`<image>`/`<rect>` が出ること) のみを
// 検証しているので、`DefaultTextMeasurer` フォールバックで十分。
vi.mock("../packages/renderer/src/font/system-font-loader.js", () => ({
  collectFontFilePaths: vi.fn(() => []),
  getSystemFontDirs: vi.fn(() => []),
}));

import { convertPptxToPng, convertPptxToSvg } from "../packages/core/src/converter.js";
import { clearFontCache } from "../packages/renderer/src/font/opentype-helpers.js";

beforeAll(() => {
  clearFontCache();
});

const FIXTURE_DIR = join(import.meta.dirname, "..", "shared-fixtures");

describe("Real PPTX E2E smoke tests", () => {
  describe("real-basic-theme.pptx (Google Slides)", () => {
    it("convertPptxToSvg completes without error", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-basic-theme.pptx"));
      const results = await convertPptxToSvg(input);

      expect(results).toHaveLength(2);
      for (const result of results) {
        expect(result.svg).toContain("<svg");
        expect(result.svg).toContain("</svg>");
      }
    });

    it("slide 1 contains title text", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-basic-theme.pptx"));
      const results = await convertPptxToSvg(input);
      const slide1 = results[0].svg;

      // タイトルとサブタイトルのテキスト（<text> or <path> via opentype）
      expect(slide1).toMatch(/<text|<path/);
    });

    it("slide 2 contains text, table, and image", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-basic-theme.pptx"));
      const results = await convertPptxToSvg(input);
      const slide2 = results[1].svg;

      // テーブル（<rect>でセルが描画される）
      expect(slide2).toContain("<rect");
      // 画像
      expect(slide2).toContain("<image");
      // テキスト
      expect(slide2).toMatch(/<text|<path/);
    });

    it("convertPptxToPng produces valid PNG", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-basic-theme.pptx"));
      const results = await convertPptxToPng(input);

      expect(results).toHaveLength(2);
      for (const result of results) {
        // PNG magic bytes
        expect(result.png[0]).toBe(0x89);
        expect(result.png[1]).toBe(0x50);
        expect(result.png[2]).toBe(0x4e);
        expect(result.png[3]).toBe(0x47);
        expect(result.width).toBeGreaterThan(0);
        expect(result.height).toBeGreaterThan(0);
      }
    });
  });

  describe("real-product-page.pptx (product landing page)", () => {
    it("convertPptxToSvg completes without error", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-product-page.pptx"));
      const results = await convertPptxToSvg(input);

      expect(results).toHaveLength(1);
      expect(results[0].svg).toContain("<svg");
      expect(results[0].svg).toContain("</svg>");
    });

    it("slide contains text elements", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-product-page.pptx"));
      const results = await convertPptxToSvg(input);
      const slide = results[0].svg;

      // タイトルやボタンなどのテキスト（<text> or <path> via opentype）
      expect(slide).toMatch(/<text|<path/);
    });

    it("slide contains shapes (roundRect, ellipse)", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-product-page.pptx"));
      const results = await convertPptxToSvg(input);
      const slide = results[0].svg;

      // カード・ボタン等の角丸矩形やアイコン用の楕円
      expect(slide).toContain("<rect");
      expect(slide).toContain("<ellipse");
    });

    it("convertPptxToPng produces valid PNG", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-product-page.pptx"));
      const results = await convertPptxToPng(input);

      expect(results).toHaveLength(1);
      // PNG magic bytes
      expect(results[0].png[0]).toBe(0x89);
      expect(results[0].png[1]).toBe(0x50);
      expect(results[0].png[2]).toBe(0x4e);
      expect(results[0].png[3]).toBe(0x47);
      expect(results[0].width).toBeGreaterThan(0);
      expect(results[0].height).toBeGreaterThan(0);
    });
  });

  describe("real-financial-report.pptx (financial report with charts)", () => {
    it("convertPptxToSvg completes without error for all 4 slides", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-financial-report.pptx"));
      const results = await convertPptxToSvg(input);

      expect(results).toHaveLength(4);
      for (const result of results) {
        expect(result.svg).toContain("<svg");
        expect(result.svg).toContain("</svg>");
      }
    });

    it("slides contain text elements", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-financial-report.pptx"));
      const results = await convertPptxToSvg(input);

      for (const result of results) {
        expect(result.svg).toMatch(/<text|<path/);
      }
    });

    it("chart slides contain rendered chart elements", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-financial-report.pptx"));
      const results = await convertPptxToSvg(input);

      // チャートはスライド2以降に含まれる（graphicFrame → <rect> / <path> で描画）
      const chartSlides = results.slice(1);
      for (const result of chartSlides) {
        expect(result.svg).toMatch(/<rect|<path/);
      }
    });

    it("convertPptxToPng produces valid PNG for all slides", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-financial-report.pptx"));
      const results = await convertPptxToPng(input);

      expect(results).toHaveLength(4);
      for (const result of results) {
        expect(result.png[0]).toBe(0x89);
        expect(result.png[1]).toBe(0x50);
        expect(result.png[2]).toBe(0x4e);
        expect(result.png[3]).toBe(0x47);
        expect(result.width).toBeGreaterThan(0);
        expect(result.height).toBeGreaterThan(0);
      }
    });
  });

  describe("sample.pptx (md-pptx generated, Japanese text)", () => {
    it("convertPptxToSvg completes without error for all 6 slides", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "sample.pptx"));
      const results = await convertPptxToSvg(input);

      expect(results).toHaveLength(6);
      for (const result of results) {
        expect(result.svg).toContain("<svg");
        expect(result.svg).toContain("</svg>");
      }
    });

    it("slides contain text elements including Japanese", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "sample.pptx"));
      const results = await convertPptxToSvg(input);

      for (const result of results) {
        expect(result.svg).toMatch(/<text|<path/);
      }
    });

    it("convertPptxToPng produces valid PNG for all slides", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "sample.pptx"));
      const results = await convertPptxToPng(input);

      expect(results).toHaveLength(6);
      for (const result of results) {
        expect(result.png[0]).toBe(0x89);
        expect(result.png[1]).toBe(0x50);
        expect(result.png[2]).toBe(0x4e);
        expect(result.png[3]).toBe(0x47);
        expect(result.width).toBeGreaterThan(0);
        expect(result.height).toBeGreaterThan(0);
      }
    });
  });

  describe("sample-issue-387.pptx (inline text formatting)", () => {
    it("convertPptxToSvg completes without error", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "sample-issue-387.pptx"));
      const results = await convertPptxToSvg(input);

      expect(results).toHaveLength(1);
      expect(results[0].svg).toContain("<svg");
      expect(results[0].svg).toContain("</svg>");
    });

    it("slide contains text elements", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "sample-issue-387.pptx"));
      const results = await convertPptxToSvg(input);
      const slide = results[0].svg;

      expect(slide).toMatch(/<text|<path/);
    });

    it("convertPptxToPng produces valid PNG", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "sample-issue-387.pptx"));
      const results = await convertPptxToPng(input);

      expect(results).toHaveLength(1);
      expect(results[0].png[0]).toBe(0x89);
      expect(results[0].png[1]).toBe(0x50);
      expect(results[0].png[2]).toBe(0x4e);
      expect(results[0].png[3]).toBe(0x47);
      expect(results[0].width).toBeGreaterThan(0);
      expect(results[0].height).toBeGreaterThan(0);
    });
  });
});
