import { describe, it, expect } from "vitest";
import { OpentypeTextMeasurer, type OpentypeFont } from "./opentype-text-measurer.js";

function createMockFont(opts: {
  unitsPerEm: number;
  ascender: number;
  descender: number;
  glyphWidths: Record<string, number>;
}): OpentypeFont {
  return {
    unitsPerEm: opts.unitsPerEm,
    ascender: opts.ascender,
    descender: opts.descender,
    stringToGlyphs: (text: string) =>
      [...text].map((ch) => ({
        advanceWidth: opts.glyphWidths[ch] ?? opts.unitsPerEm * 0.6,
      })),
  };
}

describe("OpentypeTextMeasurer", () => {
  it("登録フォントで幅を計測する", () => {
    const font = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 600 },
    });
    const fonts = new Map([["TestFont", font]]);
    const measurer = new OpentypeTextMeasurer(fonts);
    const width = measurer.measureTextWidth("A", 18, false, "TestFont");
    const expected = (600 / 1000) * 18 * (96 / 72);
    expect(width).toBeCloseTo(expected, 1);
  });

  it("太字は BOLD_FACTOR を適用する", () => {
    const font = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 600 },
    });
    const fonts = new Map([["TestFont", font]]);
    const measurer = new OpentypeTextMeasurer(fonts);
    const normalWidth = measurer.measureTextWidth("A", 18, false, "TestFont");
    const boldWidth = measurer.measureTextWidth("A", 18, true, "TestFont");
    expect(boldWidth).toBeCloseTo(normalWidth * 1.05, 1);
  });

  it("フォントが見つからない場合はデフォルト実装にフォールバック", () => {
    const measurer = new OpentypeTextMeasurer(new Map());
    const width = measurer.measureTextWidth("A", 18, false, "Unknown");
    expect(width).toBeGreaterThan(0);
  });

  it("getLineHeightRatio を計算する", () => {
    const font = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: {},
    });
    const fonts = new Map([["TestFont", font]]);
    const measurer = new OpentypeTextMeasurer(fonts);
    expect(measurer.getLineHeightRatio("TestFont")).toBeCloseTo(1.0, 5);
  });

  it("フォントが見つからない場合は 1.2 を返す", () => {
    const measurer = new OpentypeTextMeasurer(new Map());
    expect(measurer.getLineHeightRatio("Unknown")).toBe(1.2);
  });

  it("getAscenderRatio を計算する", () => {
    const font = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: {},
    });
    const fonts = new Map([["TestFont", font]]);
    const measurer = new OpentypeTextMeasurer(fonts);
    expect(measurer.getAscenderRatio("TestFont")).toBeCloseTo(0.8, 5);
  });

  it("getAscenderRatio でフォントが見つからない場合は 1.0 を返す", () => {
    const measurer = new OpentypeTextMeasurer(new Map());
    expect(measurer.getAscenderRatio("Unknown")).toBe(1.0);
  });

  it("defaultFont をフォールバックとして使用する", () => {
    const font = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 500 },
    });
    const measurer = new OpentypeTextMeasurer(new Map(), font);
    const width = measurer.measureTextWidth("A", 18, false, "Unknown");
    const expected = (500 / 1000) * 18 * (96 / 72);
    expect(width).toBeCloseTo(expected, 1);
  });

  it("fontFamilyEa で解決できる場合はそちらを使う", () => {
    const font = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 700 },
    });
    const fonts = new Map([["EaFont", font]]);
    const measurer = new OpentypeTextMeasurer(fonts);
    const width = measurer.measureTextWidth("A", 18, false, null, "EaFont");
    const expected = (700 / 1000) * 18 * (96 / 72);
    expect(width).toBeCloseTo(expected, 1);
  });
});
