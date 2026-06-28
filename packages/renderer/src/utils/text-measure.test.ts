import { describe, expect, it } from "vitest";

import { getAscenderRatio, getLineHeightRatio, measureTextWidth } from "./text-measure.js";

const PX_PER_PT = 96 / 72;

describe("measureTextWidth", () => {
  it("covers text-measure behavior 1", () => {
    expect(measureTextWidth("", 18, false)).toBe(0);
  });

  it("covers text-measure behavior 2", () => {
    const width = measureTextWidth("Hello", 18, false);
    // 'H','e','o' = normal(0.6), 'l','l' = narrow(0.3)
    // (3 * 0.6 + 2 * 0.3) * 18 * (96/72) = (1.8 + 0.6) * 24 = 57.6
    expect(width).toBeCloseTo(57.6, 1);
  });

  it("covers text-measure behavior 3", () => {
    const width = measureTextWidth("漢字", 18, false);
    // 2 * 1.0 * 18 * (96/72) = 48
    expect(width).toBeCloseTo(2 * 1.0 * 18 * PX_PER_PT, 1);
  });

  it("covers text-measure behavior 4", () => {
    const width = measureTextWidth("A漢", 18, false);
    // Test note.
    expect(width).toBeCloseTo(1.6 * 18 * PX_PER_PT, 1);
  });

  it("covers text-measure behavior 5", () => {
    const normalWidth = measureTextWidth("Test", 18, false);
    const boldWidth = measureTextWidth("Test", 18, true);
    expect(boldWidth).toBeCloseTo(normalWidth * 1.05, 1);
  });

  it("covers text-measure behavior 6", () => {
    const normalWidth = measureTextWidth("漢字", 18, false);
    const boldWidth = measureTextWidth("漢字", 18, true);
    expect(boldWidth).toBeCloseTo(normalWidth, 5);
  });

  it("covers text-measure behavior 7", () => {
    const latinNormal = measureTextWidth("A", 18, false);
    const cjkNormal = measureTextWidth("漢", 18, false);
    const mixedBold = measureTextWidth("A漢", 18, true);
    expect(mixedBold).toBeCloseTo(latinNormal * 1.05 + cjkNormal, 1);
  });

  it("covers text-measure behavior 8", () => {
    const width = measureTextWidth(" ", 18, false);
    // narrow(0.3) * 18 * (96/72) = 7.2
    expect(width).toBeCloseTo(0.3 * 18 * PX_PER_PT, 1);
  });

  it("covers text-measure behavior 9", () => {
    const width = measureTextWidth("あ", 18, false);
    expect(width).toBeCloseTo(1.0 * 18 * PX_PER_PT, 1);
  });

  it("covers text-measure behavior 10", () => {
    const width = measureTextWidth("ア", 18, false);
    expect(width).toBeCloseTo(1.0 * 18 * PX_PER_PT, 1);
  });

  it("covers text-measure behavior 11", () => {
    const width12 = measureTextWidth("A", 12, false);
    const width24 = measureTextWidth("A", 24, false);
    expect(width24).toBeCloseTo(width12 * 2, 1);
  });
});

describe("measureTextWidth with font metrics", () => {
  it("covers text-measure behavior 12", () => {
    // Carlito: A=1185, unitsPerEm=2048
    // Test note.
    const width = measureTextWidth("A", 18, false, "Calibri");
    const expected = (1185 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("covers text-measure behavior 13", () => {
    const metricsWidth = measureTextWidth("A", 18, false, "Calibri");
    const heuristicWidth = measureTextWidth("A", 18, false);
    expect(metricsWidth).not.toBeCloseTo(heuristicWidth, 0);
  });

  it("covers text-measure behavior 14", () => {
    const metricsWidth = measureTextWidth("A", 18, false, "UnknownFont");
    const heuristicWidth = measureTextWidth("A", 18, false);
    expect(metricsWidth).toBeCloseTo(heuristicWidth, 5);
  });

  it("covers text-measure behavior 15", () => {
    const metricsWidth = measureTextWidth("A", 18, false, null);
    const heuristicWidth = measureTextWidth("A", 18, false);
    expect(metricsWidth).toBeCloseTo(heuristicWidth, 5);
  });

  it("covers text-measure behavior 16", () => {
    // LiberationSans: A=1366, unitsPerEm=2048
    const width = measureTextWidth("A", 18, false, "Arial");
    const expected = (1366 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("covers text-measure behavior 17", () => {
    // Carlito: cjkWidth=2048, unitsPerEm=2048 → 1.0 * fontSizePx
    const width = measureTextWidth("漢", 18, false, "Calibri");
    const expected = (2048 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("covers text-measure behavior 18", () => {
    const normalWidth = measureTextWidth("Test", 18, false, "Calibri");
    const boldWidth = measureTextWidth("Test", 18, true, "Calibri");
    expect(boldWidth).toBeCloseTo(normalWidth * 1.05, 1);
  });

  it("covers text-measure behavior 19", () => {
    // Carlito: H=1276, e=1019, l=470, l=470, o=1080
    const width = measureTextWidth("Hello", 18, false, "Calibri");
    const expected = ((1276 + 1019 + 470 + 470 + 1080) / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("covers text-measure behavior 20", () => {
    // Carlito: defaultWidth=991, unitsPerEm=2048
    // Test note.
    const width = measureTextWidth("\u0100", 18, false, "Calibri");
    const expected = (991 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });
});

describe("measureTextWidth with fontFamilyEa", () => {
  it("covers text-measure behavior 21", () => {
    // NotoSansJP: cjkWidth=1000, unitsPerEm=1000
    const width = measureTextWidth("漢", 18, false, "Calibri", "Noto Sans JP");
    const expected = (1000 / 1000) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("covers text-measure behavior 22", () => {
    // Carlito: A=1185, unitsPerEm=2048
    const width = measureTextWidth("A", 18, false, "Calibri", "Noto Sans JP");
    const expected = (1185 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });

  it("covers text-measure behavior 23", () => {
    // Test note.
    const width = measureTextWidth("A漢", 18, false, "Calibri", "Noto Sans JP");
    const expectedLatin = (1185 / 2048) * 18 * PX_PER_PT;
    const expectedEa = (1000 / 1000) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expectedLatin + expectedEa, 1);
  });

  it("covers text-measure behavior 24", () => {
    // Carlito: cjkWidth=2048, unitsPerEm=2048
    const width = measureTextWidth("漢", 18, false, "Calibri", null);
    const expected = (2048 / 2048) * 18 * PX_PER_PT;
    expect(width).toBeCloseTo(expected, 1);
  });
});

describe("getLineHeightRatio", () => {
  it("covers text-measure behavior 25", () => {
    // Carlito: ascender=1950, descender=-550, unitsPerEm=2048
    // (1950 + 550) / 2048 = 1.220703125
    const ratio = getLineHeightRatio("Calibri");
    expect(ratio).toBeCloseTo((1950 + 550) / 2048, 5);
  });

  it("covers text-measure behavior 26", () => {
    // LiberationSans: ascender=1854, descender=-434, unitsPerEm=2048
    const ratio = getLineHeightRatio("Arial");
    expect(ratio).toBeCloseTo((1854 + 434) / 2048, 5);
  });

  it("covers text-measure behavior 27", () => {
    // NotoSansJP: ascender=1160, descender=-288, unitsPerEm=1000
    const ratio = getLineHeightRatio(null, "Meiryo");
    expect(ratio).toBeCloseTo((1160 + 288) / 1000, 5);
  });

  it("covers text-measure behavior 28", () => {
    const ratio = getLineHeightRatio("Calibri", "Meiryo");
    expect(ratio).toBeCloseTo((1950 + 550) / 2048, 5);
  });

  it("covers text-measure behavior 29", () => {
    const ratio = getLineHeightRatio(null, "Meiryo");
    expect(ratio).toBeCloseTo((1160 + 288) / 1000, 5);
  });

  it("covers text-measure behavior 30", () => {
    const ratio = getLineHeightRatio("UnknownFont");
    expect(ratio).toBe(1.2);
  });

  it("covers text-measure behavior 31", () => {
    const ratio = getLineHeightRatio(null, null);
    expect(ratio).toBe(1.2);
  });

  it("covers text-measure behavior 32", () => {
    const ratio = getLineHeightRatio();
    expect(ratio).toBe(1.2);
  });
});

describe("getAscenderRatio", () => {
  it("covers text-measure behavior 33", () => {
    // Carlito: ascender=1950, unitsPerEm=2048
    const ratio = getAscenderRatio("Calibri");
    expect(ratio).toBeCloseTo(1950 / 2048, 5);
  });

  it("covers text-measure behavior 34", () => {
    // LiberationSans: ascender=1854, unitsPerEm=2048
    const ratio = getAscenderRatio("Arial");
    expect(ratio).toBeCloseTo(1854 / 2048, 5);
  });

  it("covers text-measure behavior 35", () => {
    // NotoSansJP: ascender=1160, unitsPerEm=1000
    const ratio = getAscenderRatio(null, "Meiryo");
    expect(ratio).toBeCloseTo(1160 / 1000, 5);
  });

  it("covers text-measure behavior 36", () => {
    const ratio = getAscenderRatio("Calibri", "Meiryo");
    expect(ratio).toBeCloseTo(1950 / 2048, 5);
  });

  it("covers text-measure behavior 37", () => {
    const ratio = getAscenderRatio(null, "Meiryo");
    expect(ratio).toBeCloseTo(1160 / 1000, 5);
  });

  it("covers text-measure behavior 38", () => {
    const ratio = getAscenderRatio("UnknownFont");
    expect(ratio).toBe(1.0);
  });

  it("covers text-measure behavior 39", () => {
    const ratio = getAscenderRatio(null, null);
    expect(ratio).toBe(1.0);
  });

  it("covers text-measure behavior 40", () => {
    const ratio = getAscenderRatio();
    expect(ratio).toBe(1.0);
  });

  it("covers text-measure behavior 41", () => {
    const ascRatio = getAscenderRatio("Calibri");
    const lhRatio = getLineHeightRatio("Calibri");
    expect(ascRatio).toBeLessThan(lhRatio);
  });
});
