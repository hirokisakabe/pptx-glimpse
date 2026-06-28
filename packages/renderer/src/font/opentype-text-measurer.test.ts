import { afterEach, describe, expect, it, vi } from "vitest";

import { getWarningEntries, initWarningLogger } from "../warning-logger.js";
import { resetFontMapping, setFontMapping } from "./font-mapping-context.js";
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
  it("covers opentype-text-measurer behavior 1", () => {
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

  it("covers opentype-text-measurer behavior 2", () => {
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

  it("covers opentype-text-measurer behavior 3", () => {
    const regularFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 600 },
    });
    const boldFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 720 },
    });
    const fonts = new Map<string, OpentypeFont>([
      ["TestFont", regularFont],
      ["TestFont Bold", boldFont],
    ]);
    const measurer = new OpentypeTextMeasurer(fonts);
    const pxPerPt = 96 / 72;
    const boldWidth = measurer.measureTextWidth("A", 18, true, "TestFont");
    const expectedBoldWidth = (720 / 1000) * 18 * pxPerPt;
    expect(boldWidth).toBeCloseTo(expectedBoldWidth, 1);
    // Test note.
    const expectedWithFactor = (600 / 1000) * 18 * pxPerPt * 1.05;
    expect(boldWidth).not.toBeCloseTo(expectedWithFactor, 1);
  });

  it("covers opentype-text-measurer behavior 4", () => {
    const regularFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 600 },
    });
    const boldFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 720 },
    });
    const fonts = new Map<string, OpentypeFont>([
      ["TestFont", regularFont],
      ["TestFont-Bold", boldFont],
    ]);
    const measurer = new OpentypeTextMeasurer(fonts);
    const pxPerPt = 96 / 72;
    const boldWidth = measurer.measureTextWidth("A", 18, true, "TestFont");
    const expectedBoldWidth = (720 / 1000) * 18 * pxPerPt;
    expect(boldWidth).toBeCloseTo(expectedBoldWidth, 1);
  });

  it("covers opentype-text-measurer behavior 5", () => {
    const regularFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { 漢: 1000 },
    });
    const boldFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { 漢: 1200 },
    });
    const fonts = new Map<string, OpentypeFont>([
      ["TestFont", regularFont],
      ["TestFont Bold", boldFont],
    ]);
    const measurer = new OpentypeTextMeasurer(fonts);
    const normalWidth = measurer.measureTextWidth("漢", 18, false, "TestFont");
    const boldWidth = measurer.measureTextWidth("漢", 18, true, "TestFont");
    expect(boldWidth).toBeCloseTo(normalWidth, 5);
  });

  it("covers opentype-text-measurer behavior 6", () => {
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

  it("covers opentype-text-measurer behavior 7", () => {
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

  it("covers opentype-text-measurer behavior 8", () => {
    const measurer = new OpentypeTextMeasurer(new Map());
    const width = measurer.measureTextWidth("A", 18, false, "Unknown");
    expect(width).toBeGreaterThan(0);
  });

  it("covers opentype-text-measurer behavior 9", () => {
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

  it("covers opentype-text-measurer behavior 10", () => {
    const measurer = new OpentypeTextMeasurer(new Map());
    expect(measurer.getLineHeightRatio("Unknown")).toBe(1.2);
  });

  it("covers opentype-text-measurer behavior 11", () => {
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

  it("covers opentype-text-measurer behavior 12", () => {
    const measurer = new OpentypeTextMeasurer(new Map());
    expect(measurer.getAscenderRatio("Unknown")).toBe(1.0);
  });

  it("covers opentype-text-measurer behavior 13", () => {
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

  it("covers opentype-text-measurer behavior 14", () => {
    const latinFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 500, 漢: 300 }, // Test note.
    });
    const eaFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 400, 漢: 1000 }, // Test note.
    });
    const fonts = new Map<string, OpentypeFont>([
      ["Latin", latinFont],
      ["EA", eaFont],
    ]);
    const measurer = new OpentypeTextMeasurer(fonts);
    const pxPerPt = 96 / 72;

    // Test note.
    const width = measurer.measureTextWidth("A漢A", 18, false, "Latin", "EA");
    const expectedLatin = (500 / 1000) * 18 * pxPerPt;
    const expectedEa = (1000 / 1000) * 18 * pxPerPt;
    expect(width).toBeCloseTo(expectedLatin + expectedEa + expectedLatin, 1);
  });

  it("covers opentype-text-measurer behavior 15", () => {
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

describe("font/opentype-text-measurer.test behavior", () => {
  afterEach(() => {
    resetFontMapping();
    vi.restoreAllMocks();
  });

  it("covers opentype-text-measurer behavior 16", async () => {
    // Test note.
    const mod = await import("./cjk-font-fallback.js");
    vi.spyOn(mod, "getCjkFallbackFonts").mockReturnValue([
      "Hiragino Sans",
      "Hiragino Kaku Gothic ProN",
    ]);

    const hiraginoFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 600 },
    });
    const fonts = new Map([["Hiragino Sans", hiraginoFont]]);
    setFontMapping({ Meiryo: "Noto Sans JP" });
    const measurer = new OpentypeTextMeasurer(fonts);
    const width = measurer.measureTextWidth("A", 18, false, "Meiryo");
    const expected = (600 / 1000) * 18 * (96 / 72);
    expect(width).toBeCloseTo(expected, 1);
  });
});

describe("font/opentype-text-measurer.test behavior", () => {
  afterEach(() => {
    resetFontMapping();
  });

  it("covers opentype-text-measurer behavior 17", () => {
    const regularFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 600 },
    });
    const boldFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 750 },
    });
    const fonts = new Map<string, OpentypeFont>([
      ["Carlito", regularFont],
      ["Carlito Bold", boldFont],
    ]);
    setFontMapping({ Calibri: "Carlito" });
    const measurer = new OpentypeTextMeasurer(fonts);
    const pxPerPt = 96 / 72;
    const boldWidth = measurer.measureTextWidth("A", 18, true, "Calibri");
    const expectedBoldWidth = (750 / 1000) * 18 * pxPerPt;
    expect(boldWidth).toBeCloseTo(expectedBoldWidth, 1);
  });
});

describe("font/opentype-text-measurer.test behavior", () => {
  afterEach(() => {
    resetFontMapping();
    initWarningLogger("off");
  });

  it("covers opentype-text-measurer behavior 18", () => {
    initWarningLogger("warn");
    const measurer = new OpentypeTextMeasurer(new Map());
    measurer.measureTextWidth("A", 18, false, "UnknownFont");
    const entries = getWarningEntries();
    expect(entries.some((e) => e.feature === "font.notFound")).toBe(true);
    expect(entries.some((e) => e.message.includes("UnknownFont"))).toBe(true);
  });

  it("covers opentype-text-measurer behavior 19", () => {
    initWarningLogger("warn");
    const measurer = new OpentypeTextMeasurer(new Map());
    measurer.measureTextWidth("A", 18, false, "UnknownFont");
    measurer.measureTextWidth("A", 18, false, "UnknownFont");
    const entries = getWarningEntries().filter(
      (e) => e.feature === "font.notFound" && e.message.includes("UnknownFont"),
    );
    expect(entries).toHaveLength(1);
  });

  it("covers opentype-text-measurer behavior 20", () => {
    initWarningLogger("warn");
    const font = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 600 },
    });
    const fonts = new Map([["TestFont", font]]);
    const measurer = new OpentypeTextMeasurer(fonts);
    measurer.measureTextWidth("A", 18, false, "TestFont");
    const entries = getWarningEntries().filter((e) => e.feature === "font.notFound");
    expect(entries).toHaveLength(0);
  });
});
