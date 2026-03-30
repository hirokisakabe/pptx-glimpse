import { describe, expect, it } from "vitest";

import { type OpentypeFont, OpentypeTextMeasurer } from "./opentype-text-measurer.js";

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

  it("太字でも CJK 文字には BOLD_FACTOR を適用しない", () => {
    const font = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { 漢: 1000 },
    });
    const fonts = new Map([["TestFont", font]]);
    const measurer = new OpentypeTextMeasurer(fonts);
    const normalWidth = measurer.measureTextWidth("漢", 18, false, "TestFont");
    const boldWidth = measurer.measureTextWidth("漢", 18, true, "TestFont");
    expect(boldWidth).toBeCloseTo(normalWidth, 5);
  });

  it("太字の混合テキストではラテン文字のみ BOLD_FACTOR が適用される", () => {
    const font = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 600, 漢: 1000 },
    });
    const fonts = new Map([["TestFont", font]]);
    const measurer = new OpentypeTextMeasurer(fonts);
    const latinNormal = measurer.measureTextWidth("A", 18, false, "TestFont");
    const cjkNormal = measurer.measureTextWidth("漢", 18, false, "TestFont");
    const mixedBold = measurer.measureTextWidth("A漢", 18, true, "TestFont");
    expect(mixedBold).toBeCloseTo(latinNormal * 1.05 + cjkNormal, 1);
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

  it("混在文字列で CJK は fontFamilyEa、ラテンは fontFamily のフォントを使う", () => {
    const latinFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 500, 漢: 300 }, // ラテンフォントの CJK 幅は不正確
    });
    const eaFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 400, 漢: 1000 }, // EA フォントの CJK 幅は正確
    });
    const fonts = new Map<string, OpentypeFont>([
      ["Latin", latinFont],
      ["EA", eaFont],
    ]);
    const measurer = new OpentypeTextMeasurer(fonts);
    const pxPerPt = 96 / 72;

    // "A漢A" → A は latinFont(500), 漢 は eaFont(1000), A は latinFont(500)
    const width = measurer.measureTextWidth("A漢A", 18, false, "Latin", "EA");
    const expectedLatin = (500 / 1000) * 18 * pxPerPt;
    const expectedEa = (1000 / 1000) * 18 * pxPerPt;
    expect(width).toBeCloseTo(expectedLatin + expectedEa + expectedLatin, 1);
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
