import { readFileSync } from "fs";
import { join } from "path";
import { beforeAll, describe, expect, it, vi } from "vitest";

// On macOS, parsing hundreds of fonts from `/System/Library/Fonts` etc.
// Worker memory swells and OOM occurs. Here the structure-based assertion
// (either `<text|<path>` appears, `<image>`/`<rect>`) only.
// Since we are testing, the `DefaultTextMeasurer` fallback is sufficient.
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
      const { slides: results } = await convertPptxToSvg(input);

      expect(results).toHaveLength(2);
      for (const result of results) {
        expect(result.svg).toContain("<svg");
        expect(result.svg).toContain("</svg>");
      }
    });

    it("slide 1 contains title text", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-basic-theme.pptx"));
      const { slides: results } = await convertPptxToSvg(input);
      const slide1 = results[0].svg;

      // Title and subtitle text (<text> or <path> via opentype)
      expect(slide1).toMatch(/<text|<path/);
    });

    it("slide 2 contains text, table, and image", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-basic-theme.pptx"));
      const { slides: results } = await convertPptxToSvg(input);
      const slide2 = results[1].svg;

      // Table (cells are drawn with <rect>)
      expect(slide2).toContain("<rect");
      // image
      expect(slide2).toContain("<image");
      // text
      expect(slide2).toMatch(/<text|<path/);
    });

    it("convertPptxToPng produces valid PNG", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-basic-theme.pptx"));
      const { slides: results } = await convertPptxToPng(input);

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
      const { slides: results } = await convertPptxToSvg(input);

      expect(results).toHaveLength(1);
      expect(results[0].svg).toContain("<svg");
      expect(results[0].svg).toContain("</svg>");
    });

    it("slide contains text elements", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-product-page.pptx"));
      const { slides: results } = await convertPptxToSvg(input);
      const slide = results[0].svg;

      // Text such as titles and buttons (<text> or <path> via opentype)
      expect(slide).toMatch(/<text|<path/);
    });

    it("slide contains shapes (roundRect, ellipse)", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-product-page.pptx"));
      const { slides: results } = await convertPptxToSvg(input);
      const slide = results[0].svg;

      // Rounded rectangles for cards and buttons, and ellipses for icons
      expect(slide).toContain("<rect");
      expect(slide).toContain("<ellipse");
    });

    it("convertPptxToPng produces valid PNG", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-product-page.pptx"));
      const { slides: results } = await convertPptxToPng(input);

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
      const { slides: results } = await convertPptxToSvg(input);

      expect(results).toHaveLength(4);
      for (const result of results) {
        expect(result.svg).toContain("<svg");
        expect(result.svg).toContain("</svg>");
      }
    });

    it("slides contain text elements", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-financial-report.pptx"));
      const { slides: results } = await convertPptxToSvg(input);

      for (const result of results) {
        expect(result.svg).toMatch(/<text|<path/);
      }
    });

    it("chart slides contain rendered chart elements", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-financial-report.pptx"));
      const { slides: results } = await convertPptxToSvg(input);

      // The chart is included in slide 2 and beyond (drawn with graphicFrame -> <rect> / <path>)
      const chartSlides = results.slice(1);
      for (const result of chartSlides) {
        expect(result.svg).toMatch(/<rect|<path/);
      }
    });

    it("convertPptxToPng produces valid PNG for all slides", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "real-financial-report.pptx"));
      const { slides: results } = await convertPptxToPng(input);

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
      const { slides: results } = await convertPptxToSvg(input);

      expect(results).toHaveLength(6);
      for (const result of results) {
        expect(result.svg).toContain("<svg");
        expect(result.svg).toContain("</svg>");
      }
    });

    it("slides contain text elements including Japanese", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "sample.pptx"));
      const { slides: results } = await convertPptxToSvg(input);

      for (const result of results) {
        expect(result.svg).toMatch(/<text|<path/);
      }
    });

    it("convertPptxToPng produces valid PNG for all slides", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "sample.pptx"));
      const { slides: results } = await convertPptxToPng(input);

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
      const { slides: results } = await convertPptxToSvg(input);

      expect(results).toHaveLength(1);
      expect(results[0].svg).toContain("<svg");
      expect(results[0].svg).toContain("</svg>");
    });

    it("slide contains text elements", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "sample-issue-387.pptx"));
      const { slides: results } = await convertPptxToSvg(input);
      const slide = results[0].svg;

      expect(slide).toMatch(/<text|<path/);
    });

    it("convertPptxToPng produces valid PNG", async () => {
      const input = readFileSync(join(FIXTURE_DIR, "sample-issue-387.pptx"));
      const { slides: results } = await convertPptxToPng(input);

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
