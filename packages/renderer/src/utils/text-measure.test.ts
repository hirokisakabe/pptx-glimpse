import { describe, expect, it } from "vitest";

import { getAscenderRatio, getLineHeightRatio, measureTextWidth } from "./text-measure.js";

const PX_PER_PT = 96 / 72;

describe("measureTextWidth", () => {
  it("The width of an empty string is 0", () => {
    expect(measureTextWidth("", 18, false)).toBe(0);
  });

  it("Estimate width of ASCII text", () => {
    const width = measureTextWidth("Hello", 18, false);
    // 'H','e','o' = normal(0.6), 'l','l' = narrow(0.3)
    // (3 * 0.6 + 2 * 0.3) * 18 * (96/72) = (1.8 + 0.6) * 24 = 57.6
    expect(width).toBeCloseTo(57.6, 1);
  });

  it("Estimate the width of CJK text", () => {
    const width = measureTextWidth("漢字", 18, false);
    // 2 * 1.0 * 18 * (96/72) = 48
    expect(width).toBeCloseTo(2 * 1.0 * 18 * PX_PER_PT, 1);
  });

  it("Estimate width of mixed text", () => {
    const width = measureTextWidth("A漢", 18, false);
    // A=normal(0.6) + Chinese=wide(1.0) = 1.6 * 18 * (96/72) = 38.4
    expect(width).toBeCloseTo(1.6 * 18 * PX_PER_PT, 1);
  });

  it("Bold increases the width of latin characters", () => {
    const normalWidth = measureTextWidth("Test", 18, false);
    const boldWidth = measureTextWidth("Test", 18, true);
    expect(boldWidth).toBeCloseTo(normalWidth * 1.05, 1);
  });

  it("Bold does not change the width of CJK characters", () => {
    const normalWidth = measureTextWidth("漢字", 18, false);
    const boldWidth = measureTextWidth("漢字", 18, true);
    expect(boldWidth).toBeCloseTo(normalWidth, 5);
  });

  it("BOLD_FACTOR only applies to Latin characters in bold mixed text", () => {
    const latinNormal = measureTextWidth("A", 18, false);
    const cjkNormal = measureTextWidth("漢", 18, false);
    const mixedBold = measureTextWidth("A漢", 18, true);
    expect(mixedBold).toBeCloseTo(latinNormal * 1.05 + cjkNormal, 1);
  });

  it("Estimate the width of a space", () => {
    const width = measureTextWidth(" ", 18, false);
    // narrow(0.3) * 18 * (96/72) = 7.2
    expect(width).toBeCloseTo(0.3 * 18 * PX_PER_PT, 1);
  });

  it("Estimate hiragana as wide", () => {
    const width = measureTextWidth("あ", 18, false);
    expect(width).toBeCloseTo(1.0 * 18 * PX_PER_PT, 1);
  });

  it("Estimate katakana as wide", () => {
    const width = measureTextWidth("ア", 18, false);
    expect(width).toBeCloseTo(1.0 * 18 * PX_PER_PT, 1);
  });

  it("proportional to font size", () => {
    const width12 = measureTextWidth("A", 12, false);
    const width24 = measureTextWidth("A", 24, false);
    expect(width24).toBeCloseTo(width12 * 2, 1);
  });
});

describe("measureTextWidth with font metrics", () => {
  it("Calculate ASCII text width with Calibri metrics", () => {
    // Carlito: A=1185, unitsPerEm=2048
    // Width = 1185 / 2048 * 18 * (96/72)
    const width = measureTextWidth("A", 18, false, "Calibri");
    const expected = (1185 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("returns a different value than the heuristic", () => {
    const metricsWidth = measureTextWidth("A", 18, false, "Calibri");
    const heuristicWidth = measureTextWidth("A", 18, false);
    expect(metricsWidth).not.toBeCloseTo(heuristicWidth, 0);
  });

  it("Fallback to heuristics on unknown fonts", () => {
    const metricsWidth = measureTextWidth("A", 18, false, "UnknownFont");
    const heuristicWidth = measureTextWidth("A", 18, false);
    expect(metricsWidth).toBeCloseTo(heuristicWidth, 5);
  });

  it("Fallback to heuristic if fontFamily is null", () => {
    const metricsWidth = measureTextWidth("A", 18, false, null);
    const heuristicWidth = measureTextWidth("A", 18, false);
    expect(metricsWidth).toBeCloseTo(heuristicWidth, 5);
  });

  it("Calculate with Arial metrics", () => {
    // LiberationSans: A=1366, unitsPerEm=2048
    const width = measureTextWidth("A", 18, false, "Arial");
    const expected = (1366 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("CJK text uses cjkWidth", () => {
    // Carlito: cjkWidth=2048, unitsPerEm=2048 -> 1.0 * fontSizePx
    const width = measureTextWidth("漢", 18, false, "Calibri");
    const expected = (2048 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("Bold also applies BOLD_FACTOR to metric-based widths", () => {
    const normalWidth = measureTextWidth("Test", 18, false, "Calibri");
    const boldWidth = measureTextWidth("Test", 18, true, "Calibri");
    expect(boldWidth).toBeCloseTo(normalWidth * 1.05, 1);
  });

  it("Add widths of multiple characters correctly", () => {
    // Carlito: H=1276, e=1019, l=470, l=470, o=1080
    const width = measureTextWidth("Hello", 18, false, "Calibri");
    const expected = ((1276 + 1019 + 470 + 470 + 1080) / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("Use defaultWidth for characters not found in metrics", () => {
    // Carlito: defaultWidth=991, unitsPerEm=2048
    // U+0100 (Ā) is a Latin extended character not included in metrics
    const width = measureTextWidth("\u0100", 18, false, "Calibri");
    const expected = (991 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });
});

describe("measureTextWidth with fontFamilyEa", () => {
  it("Use ea font metrics for CJK characters", () => {
    // NotoSansJP: cjkWidth=1000, unitsPerEm=1000
    const width = measureTextWidth("漢", 18, false, "Calibri", "Noto Sans JP");
    const expected = (1000 / 1000) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("Use latin font metrics for latin characters", () => {
    // Carlito: A=1185, unitsPerEm=2048
    const width = measureTextWidth("A", 18, false, "Calibri", "Noto Sans JP");
    const expected = (1185 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("Use different metrics for each character in mixed text", () => {
    // A: Carlito (1185/2048), Chinese: NotoSansJP (1000/1000)
    const width = measureTextWidth("A漢", 18, false, "Calibri", "Noto Sans JP");
    const expectedLatin = (1185 / 2048) * 18 * PX_PER_PT;
    const expectedEa = (1000 / 1000) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expectedLatin + expectedEa, 1);
  });

  it("If fontFamilyEa is null, also measure CJK with latin metrics", () => {
    // Carlito: cjkWidth=2048, unitsPerEm=2048
    const width = measureTextWidth("漢", 18, false, "Calibri", null);
    const expected = (2048 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });
});

describe("getLineHeightRatio", () => {
  it("Returns the row height ratio for Calibri (Carlito)", () => {
    // Carlito: ascender=1950, descender=-550, unitsPerEm=2048
    // (1950 + 550) / 2048 = 1.220703125
    const ratio = getLineHeightRatio("Calibri");
    expect(ratio).toBeCloseTo((1950 + 550) / 2048, 5);
  });

  it("Returns the row height ratio of Arial (LiberationSans)", () => {
    // LiberationSans: ascender=1854, descender=-434, unitsPerEm=2048
    const ratio = getLineHeightRatio("Arial");
    expect(ratio).toBeCloseTo((1854 + 434) / 2048, 5);
  });

  it("Returns the row height ratio of NotoSansJP", () => {
    // NotoSansJP: ascender=1160, descender=-288, unitsPerEm=1000
    const ratio = getLineHeightRatio(null, "Meiryo");
    expect(ratio).toBeCloseTo((1160 + 288) / 1000, 5);
  });

  it("Use fontFamily preferentially", () => {
    const ratio = getLineHeightRatio("Calibri", "Meiryo");
    expect(ratio).toBeCloseTo((1950 + 550) / 2048, 5);
  });

  it("Use fontFamilyEa if fontFamily is null", () => {
    const ratio = getLineHeightRatio(null, "Meiryo");
    expect(ratio).toBeCloseTo((1160 + 288) / 1000, 5);
  });

  it("Returns default value 1.2 for unknown fonts", () => {
    const ratio = getLineHeightRatio("UnknownFont");
    expect(ratio).toBe(1.2);
  });

  it("If both are null, return default value 1.2", () => {
    const ratio = getLineHeightRatio(null, null);
    expect(ratio).toBe(1.2);
  });

  it("If no argument, returns default value 1.2", () => {
    const ratio = getLineHeightRatio();
    expect(ratio).toBe(1.2);
  });
});

describe("getAscenderRatio", () => {
  it("Returns the ascender ratio of Calibri (Carlito)", () => {
    // Carlito: ascender=1950, unitsPerEm=2048
    const ratio = getAscenderRatio("Calibri");
    expect(ratio).toBeCloseTo(1950 / 2048, 5);
  });

  it("Returns the ascender ratio of Arial (LiberationSans)", () => {
    // LiberationSans: ascender=1854, unitsPerEm=2048
    const ratio = getAscenderRatio("Arial");
    expect(ratio).toBeCloseTo(1854 / 2048, 5);
  });

  it("Returns the ascender ratio of NotoSansJP", () => {
    // NotoSansJP: ascender=1160, unitsPerEm=1000
    const ratio = getAscenderRatio(null, "Meiryo");
    expect(ratio).toBeCloseTo(1160 / 1000, 5);
  });

  it("Use fontFamily preferentially", () => {
    const ratio = getAscenderRatio("Calibri", "Meiryo");
    expect(ratio).toBeCloseTo(1950 / 2048, 5);
  });

  it("Use fontFamilyEa if fontFamily is null", () => {
    const ratio = getAscenderRatio(null, "Meiryo");
    expect(ratio).toBeCloseTo(1160 / 1000, 5);
  });

  it("Returns default value 1.0 for unknown fonts", () => {
    const ratio = getAscenderRatio("UnknownFont");
    expect(ratio).toBe(1.0);
  });

  it("If both are null, return default value 1.0", () => {
    const ratio = getAscenderRatio(null, null);
    expect(ratio).toBe(1.0);
  });

  it("If no argument, returns default value 1.0", () => {
    const ratio = getAscenderRatio();
    expect(ratio).toBe(1.0);
  });

  it("Returns a value less than the row height ratio", () => {
    const ascRatio = getAscenderRatio("Calibri");
    const lhRatio = getLineHeightRatio("Calibri");
    expect(ascRatio).toBeLessThan(lhRatio);
  });
});
