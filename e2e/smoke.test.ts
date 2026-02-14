import { readFileSync } from "fs";
import { join } from "path";
import { describe, it, expect } from "vitest";
import { convertPptxToSvg, convertPptxToPng } from "../src/converter.js";

const FIXTURE_DIR = join(import.meta.dirname, "fixtures");

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
});
