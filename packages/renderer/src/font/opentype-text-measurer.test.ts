import { afterEach, describe, expect, it, vi } from "vitest";

import { createWarningLogger, getWarningEntries, initWarningLogger } from "../warning-logger.js";
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
  it("Measure width using registered font", () => {
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

  it("BOLD applies BOLD_FACTOR", () => {
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

  it("Use real glyph width if bold font (${fontFamily} Bold) is present", () => {
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
    // Check that the width is different from the width when BOLD_FACTOR is applied.
    const expectedWithFactor = (600 / 1000) * 18 * pxPerPt * 1.05;
    expect(boldWidth).not.toBeCloseTo(expectedWithFactor, 1);
  });

  it("Use real glyph width if bold font (${fontFamily}-Bold) is present", () => {
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

  it("Does not apply to CJK characters even if bold font exists", () => {
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

  it("Do not apply BOLD_FACTOR to CJK characters even if they are bold", () => {
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

  it("BOLD_FACTOR only applies to Latin characters in bold mixed text", () => {
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

  it("Fallback to default implementation if font not found", () => {
    const measurer = new OpentypeTextMeasurer(new Map());
    const width = measurer.measureTextWidth("A", 18, false, "Unknown");
    expect(width).toBeGreaterThan(0);
  });

  it("Calculate getLineHeightRatio", () => {
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

  it("Returns 1.2 if font is not found", () => {
    const measurer = new OpentypeTextMeasurer(new Map());
    expect(measurer.getLineHeightRatio("Unknown")).toBe(1.2);
  });

  it("Calculate getAscenderRatio", () => {
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

  it("getAscenderRatio returns 1.0 if font is not found", () => {
    const measurer = new OpentypeTextMeasurer(new Map());
    expect(measurer.getAscenderRatio("Unknown")).toBe(1.0);
  });

  it("Use defaultFont as a fallback", () => {
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

  it("Mixed strings use fontFamilyEa for CJK and fontFamily for Latin", () => {
    const latinFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 500, 漢: 300 }, // CJK width for Latin fonts is inaccurate
    });
    const eaFont = createMockFont({
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      glyphWidths: { A: 400, 漢: 1000 }, // CJK width of EA fonts is accurate
    });
    const fonts = new Map<string, OpentypeFont>([
      ["Latin", latinFont],
      ["EA", eaFont],
    ]);
    const measurer = new OpentypeTextMeasurer(fonts);
    const pxPerPt = 96 / 72;

    // "AkanA" -> A is latinFont(500), Kan is eaFont(1000), A is latinFont(500)
    const width = measurer.measureTextWidth("A漢A", 18, false, "Latin", "EA");
    const expectedLatin = (500 / 1000) * 18 * pxPerPt;
    const expectedEa = (1000 / 1000) * 18 * pxPerPt;
    expect(width).toBeCloseTo(expectedLatin + expectedEa + expectedLatin, 1);
  });

  it("If fontFamilyEa can solve the problem, use it", () => {
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

describe("OpentypeTextMeasurer CJK fallback", () => {
  afterEach(() => {
    resetFontMapping();
    vi.restoreAllMocks();
  });

  it("Attempt CJK fallback chain if no mapping is found", async () => {
    // In Linux CI, the fallback chain is empty, so mock the macOS equivalent value.
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

describe("OpentypeTextMeasurer font mapping", () => {
  afterEach(() => {
    resetFontMapping();
  });

  it("Resolving OSS bold fonts via font mapping", () => {
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

describe("OpentypeTextMeasurer font warnings", () => {
  afterEach(() => {
    resetFontMapping();
    initWarningLogger("off");
  });

  it("Issue font.notFound warning if font is not found", () => {
    initWarningLogger("warn");
    const measurer = new OpentypeTextMeasurer(new Map());
    measurer.measureTextWidth("A", 18, false, "UnknownFont");
    const entries = getWarningEntries();
    expect(entries.some((e) => e.feature === "font.notFound")).toBe(true);
    expect(entries.some((e) => e.message.includes("UnknownFont"))).toBe(true);
  });

  it("records font.notFound warnings into the provided measurement context", () => {
    initWarningLogger("warn");
    const logger = createWarningLogger("warn");
    const measurer = new OpentypeTextMeasurer(new Map());

    const fontWarningCache = new Set<string>();
    const context = {
      warningLogger: logger,
      fontWarningCache,
    };

    measurer.measureTextWidth("A", 18, false, "UnknownFont", null, context);
    measurer.measureTextWidth("A", 18, false, "UnknownFont", null, context);

    expect(
      logger
        .getWarningEntries()
        .some(
          (entry) => entry.feature === "font.notFound" && entry.message.includes("UnknownFont"),
        ),
    ).toBe(true);
    expect(
      logger
        .getWarningEntries()
        .filter(
          (entry) => entry.feature === "font.notFound" && entry.message.includes("UnknownFont"),
        ),
    ).toHaveLength(1);
    expect(getWarningEntries()).toHaveLength(0);
  });

  it("Warnings with the same font name are not duplicated", () => {
    initWarningLogger("warn");
    const measurer = new OpentypeTextMeasurer(new Map());
    measurer.measureTextWidth("A", 18, false, "UnknownFont");
    measurer.measureTextWidth("A", 18, false, "UnknownFont");
    const entries = getWarningEntries().filter(
      (e) => e.feature === "font.notFound" && e.message.includes("UnknownFont"),
    );
    expect(entries).toHaveLength(1);
  });

  it("Don't warn if font is found", () => {
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
