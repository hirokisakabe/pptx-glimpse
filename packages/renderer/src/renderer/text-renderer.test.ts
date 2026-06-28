import { afterEach, describe, expect, it } from "vitest";

import type { OpentypeFullFont } from "../font/text-path-context.js";
import {
  DefaultTextPathFontResolver,
  resetTextPathFontResolver,
  setTextPathFontResolver,
} from "../font/text-path-context.js";
import type { Transform } from "../model/shape.js";
import type {
  BodyProperties,
  BulletType,
  Paragraph,
  RunProperties,
  SpacingValue,
  TextBody,
} from "../model/text.js";
import {
  buildFontFamilyValue,
  computeSpAutofitHeight,
  formatAutoNum,
  renderTextBody,
} from "./text-renderer.js";

function makeTransform(widthEmu: number, heightEmu: number): Transform {
  return {
    offsetX: 0,
    offsetY: 0,
    extentWidth: widthEmu,
    extentHeight: heightEmu,
    rotation: 0,
    flipH: false,
    flipV: false,
  };
}

function defaultParagraphProperties(
  overrides?: Partial<Paragraph["properties"]>,
): Paragraph["properties"] {
  return {
    alignment: "l",
    lineSpacing: null,
    spaceBefore: { type: "pts", value: 0 },
    spaceAfter: { type: "pts", value: 0 },
    level: 0,
    bullet: null,
    bulletFont: null,
    bulletColor: null,
    bulletSizePct: null,
    marginLeft: 0,
    indent: 0,
    tabStops: [],
    ...overrides,
  };
}

function makeTextBody(
  texts: string[],
  overrides?: {
    wrap?: "square" | "none";
    anchor?: "t" | "ctr" | "b";
    alignment?: "l" | "ctr" | "r" | "just";
    fontSize?: number;
    fontScale?: number;
    lnSpcReduction?: number;
    autoFit?: "noAutofit" | "normAutofit" | "spAutofit";
    vert?: BodyProperties["vert"];
  },
): TextBody {
  const autoFit =
    overrides?.autoFit ?? (overrides?.fontScale !== undefined ? "normAutofit" : "noAutofit");
  return {
    bodyProperties: {
      anchor: overrides?.anchor ?? "t",
      marginLeft: 91440, // ~9.6px
      marginRight: 91440,
      marginTop: 45720, // ~4.8px
      marginBottom: 45720,
      wrap: overrides?.wrap ?? "square",
      autoFit,
      fontScale: overrides?.fontScale ?? 1,
      lnSpcReduction: overrides?.lnSpcReduction ?? 0,
      numCol: 1,
      vert: overrides?.vert ?? "horz",
    },
    paragraphs: [
      {
        runs: texts.map((text) => ({
          text,
          properties: {
            fontSize: overrides?.fontSize ?? 18,
            fontFamily: null,
            fontFamilyEa: null,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            color: null,
            baseline: 0,
          },
        })),
        properties: defaultParagraphProperties({
          alignment: overrides?.alignment ?? "l",
        }),
      },
    ],
  };
}

function makeBulletTextBody(
  paragraphs: {
    text: string;
    bullet: BulletType | null;
    marginLeft?: number;
    indent?: number;
    level?: number;
  }[],
): TextBody {
  return {
    bodyProperties: {
      anchor: "t",
      marginLeft: 91440,
      marginRight: 91440,
      marginTop: 45720,
      marginBottom: 45720,
      wrap: "square",
      autoFit: "noAutofit",
      fontScale: 1,
      lnSpcReduction: 0,
      numCol: 1,
      vert: "horz" as const,
    },
    paragraphs: paragraphs.map((p) => ({
      runs: [
        {
          text: p.text,
          properties: {
            fontSize: 18,
            fontFamily: null,
            fontFamilyEa: null,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            color: null,
            baseline: 0,
          },
        },
      ],
      properties: defaultParagraphProperties({
        bullet: p.bullet,
        marginLeft: p.marginLeft ?? 342900,
        indent: p.indent ?? -342900,
        level: p.level ?? 0,
      }),
    })),
  };
}

// 16:9 Slide size (EMU)
const SLIDE_WIDTH = 9144000;
const SLIDE_HEIGHT = 5143500;

describe("renderTextBody", () => {
  it("Returns an empty string if there is no text", () => {
    const textBody = makeTextBody([""]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toBe("");
  });

  it("Render short text correctly", () => {
    const textBody = makeTextBody(["Hello"]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("<text");
    expect(result).toContain("Hello");
    expect(result).toContain("<tspan");
  });

  it("If wrap=square, long text will be wrapped into multiple tspans", () => {
    const textBody = makeTextBody(
      ["The quick brown fox jumps over the lazy dog and continues with more words"],
      { wrap: "square" },
    );
    // narrow text box
    const transform = makeTransform(2000000, 2000000); // ~209px width
    const result = renderTextBody(textBody, transform);

    // There is a tspan with multiple x attributes (= multiple lines)
    const xMatches = result.match(/x="/g);
    expect(xMatches).not.toBeNull();
    expect(xMatches!.length).toBeGreaterThan(1);
  });

  it("If wrap=none, the text will not wrap", () => {
    const textBody = makeTextBody(
      ["The quick brown fox jumps over the lazy dog and continues with more words"],
      { wrap: "none" },
    );
    const transform = makeTransform(2000000, 2000000);
    const result = renderTextBody(textBody, transform);

    // Only one tspan has an x attribute (excluding x on the text element)
    const tspanXMatches = result.match(/<tspan[^>]*x="/g);
    expect(tspanXMatches).toHaveLength(1);
  });

  it("For center alignment text-anchor=middle is set", () => {
    const textBody = makeTextBody(["Center"], { alignment: "ctr" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('text-anchor="middle"');
  });

  it("For right alignment, text-anchor=end is set.", () => {
    const textBody = makeTextBody(["Right"], { alignment: "r" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('text-anchor="end"');
  });

  it("For left alignment, text-anchor=start is set.", () => {
    const textBody = makeTextBody(["Left"], { alignment: "l" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('text-anchor="start"');
  });

  it("Render multiple paragraphs correctly", () => {
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "First",
              properties: {
                fontSize: 18,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
        {
          runs: [
            {
              text: "Second",
              properties: {
                fontSize: 18,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
      ],
    };
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("First");
    expect(result).toContain("Second");
  });

  it("font-size attribute is set correctly", () => {
    const textBody = makeTextBody(["Test"], { fontSize: 24 });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('font-size="24pt"');
  });

  it("Font size is reduced when fontScale is applied", () => {
    const textBody = makeTextBody(["Test"], { fontSize: 24, fontScale: 0.625 });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    // 24 * 0.625 = 15
    expect(result).toContain('font-size="15pt"');
  });

  it("If fontScale=1, the font size will not change", () => {
    const textBody = makeTextBody(["Test"], { fontSize: 24, fontScale: 1 });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('font-size="24pt"');
  });

  it("CJK text wrapping works", () => {
    const textBody = makeTextBody(["本日は晴天なり今日もいい天気です素晴らしい一日"], {
      wrap: "square",
    });
    const transform = makeTransform(2000000, 2000000);
    const result = renderTextBody(textBody, transform);
    const xMatches = result.match(/<tspan[^>]*x="/g);
    expect(xMatches).not.toBeNull();
    expect(xMatches!.length).toBeGreaterThan(1);
  });

  it("Superscript is set to baseline-shift=super", () => {
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "H",
              properties: {
                fontSize: 18,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
            {
              text: "2",
              properties: {
                fontSize: 12,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 30,
              },
            },
            {
              text: "O",
              properties: {
                fontSize: 18,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
      ],
    };
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('baseline-shift="super"');
  });

  it("Baseline-shift=sub is set for subscripts", () => {
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "H",
              properties: {
                fontSize: 18,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
            {
              text: "2",
              properties: {
                fontSize: 12,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: -25,
              },
            },
            {
              text: "O",
              properties: {
                fontSize: 18,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
      ],
    };
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('baseline-shift="sub"');
  });

  it("baseline-shift is not set if baseline=0", () => {
    const textBody = makeTextBody(["Normal"]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).not.toContain("baseline-shift");
  });
});

describe("text renderer paragraph spacing", () => {
  function makeRunProps(fontSize: number = 18) {
    return {
      fontSize,
      fontFamily: null,
      fontFamilyEa: null,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      color: null,
      baseline: 0,
    };
  }

  function makeSpacingTextBody(
    paragraphs: { text: string; spaceBefore?: SpacingValue; spaceAfter?: SpacingValue }[],
  ): TextBody {
    return {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: paragraphs.map((p) => ({
        runs: [{ text: p.text, properties: makeRunProps() }],
        properties: defaultParagraphProperties({
          spaceBefore: p.spaceBefore ?? { type: "pts", value: 0 },
          spaceAfter: p.spaceAfter ?? { type: "pts", value: 0 },
        }),
      })),
    };
  }

  function extractDyValues(svg: string): number[] {
    const matches = svg.matchAll(/dy="([^"]+)"/g);
    return [...matches].map((m) => parseFloat(m[1]));
  }

  it("spaceBefore (pts) is reflected in paragraph spacing", () => {
    const withSpacing = makeSpacingTextBody([
      { text: "First" },
      { text: "Second", spaceBefore: { type: "pts", value: 1200 } }, // 12pt
    ]);
    const withoutSpacing = makeSpacingTextBody([{ text: "First" }, { text: "Second" }]);

    const dyWith = extractDyValues(
      renderTextBody(withSpacing, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT)),
    );
    const dyWithout = extractDyValues(
      renderTextBody(withoutSpacing, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT)),
    );

    // If spaceBefore is set, the second paragraph's dy will be larger
    expect(dyWith[1]).toBeGreaterThan(dyWithout[1]);
  });

  it("spaceAfter (pts) is reflected in the spacing between next paragraph", () => {
    const withSpaceAfter = makeSpacingTextBody([
      { text: "First", spaceAfter: { type: "pts", value: 1200 } }, // 12pt
      { text: "Second" },
    ]);
    const withoutSpacing = makeSpacingTextBody([{ text: "First" }, { text: "Second" }]);

    const dyWith = extractDyValues(
      renderTextBody(withSpaceAfter, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT)),
    );
    const dyWithout = extractDyValues(
      renderTextBody(withoutSpacing, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT)),
    );

    // If spaceAfter is set, the second paragraph's dy will be larger
    expect(dyWith[1]).toBeGreaterThan(dyWithout[1]);
  });

  it("The larger of spaceAfter and spaceBefore is applied.", () => {
    const spaceAfterOnly = makeSpacingTextBody([
      { text: "First", spaceAfter: { type: "pts", value: 2000 } }, // 20pt
      { text: "Second", spaceBefore: { type: "pts", value: 500 } }, // 5pt
    ]);
    const spaceBeforeOnly = makeSpacingTextBody([
      { text: "First", spaceAfter: { type: "pts", value: 500 } }, // 5pt
      { text: "Second", spaceBefore: { type: "pts", value: 2000 } }, // 20pt
    ]);

    const dyAfter = extractDyValues(
      renderTextBody(spaceAfterOnly, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT)),
    );
    const dyBefore = extractDyValues(
      renderTextBody(spaceBeforeOnly, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT)),
    );

    // A 20pt interval is applied to both (max is used, so the same dy)
    expect(dyAfter[1]).toBeCloseTo(dyBefore[1], 1);
  });

  it("spaceBefore (pct) is calculated based on font size", () => {
    const withPct = makeSpacingTextBody([
      { text: "First" },
      { text: "Second", spaceBefore: { type: "pct", value: 100000 } }, // 100%
    ]);
    const withoutSpacing = makeSpacingTextBody([{ text: "First" }, { text: "Second" }]);

    const dyWith = extractDyValues(
      renderTextBody(withPct, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT)),
    );
    const dyWithout = extractDyValues(
      renderTextBody(withoutSpacing, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT)),
    );

    // 100% = additional interval for font size (18pt)
    expect(dyWith[1]).toBeGreaterThan(dyWithout[1]);
    // Additional spacing is 18pt * (96/72) = 24px
    const diff = dyWith[1] - dyWithout[1];
    expect(diff).toBeCloseTo(18 * (96 / 72), 1);
  });
});

describe("text renderer line spacing", () => {
  function makeLineSpacingTextBody(
    lineSpacing: SpacingValue | null,
    fontSize: number = 14,
  ): TextBody {
    return {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "The quick brown fox jumps over the lazy dog and continues with more words",
              properties: {
                fontSize,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties({ lineSpacing }),
        },
      ],
    };
  }

  function extractDyValues(svg: string): number[] {
    const matches = svg.matchAll(/dy="([^"]+)"/g);
    return [...matches].map((m) => parseFloat(m[1]));
  }

  // Narrow text box that wraps to 2 or more lines (~209px width)
  const NARROW_TRANSFORM = makeTransform(2000000, 2000000);

  it("With spcPts (fixed leading), the baseline spacing is a fixed value independent of font size.", () => {
    // 21pt fixed = 28px @96dpi
    const textBody = makeLineSpacingTextBody({ type: "pts", value: 2100 });
    const dyValues = extractDyValues(renderTextBody(textBody, NARROW_TRANSFORM));

    expect(dyValues.length).toBeGreaterThan(1);
    for (const dy of dyValues.slice(1)) {
      expect(dy).toBeCloseTo(28, 1);
    }
  });

  it("spcPct (magnification) makes line leading proportional to font size", () => {
    const dy100 = extractDyValues(
      renderTextBody(makeLineSpacingTextBody({ type: "pct", value: 100000 }), NARROW_TRANSFORM),
    );
    const dy200 = extractDyValues(
      renderTextBody(makeLineSpacingTextBody({ type: "pct", value: 200000 }), NARROW_TRANSFORM),
    );

    expect(dy100.length).toBeGreaterThan(1);
    expect(dy200.length).toBeGreaterThan(1);
    expect(dy200[1]).toBeCloseTo(dy100[1] * 2, 1);
  });

  it("spcPts (fixed leading) is applied even to empty paragraphs", () => {
    const makeParagraph = (text: string, lineSpacing: SpacingValue | null): Paragraph => ({
      runs: text
        ? [
            {
              text,
              properties: {
                fontSize: 14,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
          ]
        : [],
      properties: defaultParagraphProperties({ lineSpacing }),
    });
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        makeParagraph("First", null),
        makeParagraph("", { type: "pts", value: 2100 }),
        makeParagraph("Second", null),
      ],
    };
    const dyValues = extractDyValues(
      renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT)),
    );

    // The leading of empty paragraph (2nd) is fixed at 21pt = 28px
    expect(dyValues).toHaveLength(3);
    expect(dyValues[1]).toBeCloseTo(28, 1);
  });

  it("Without lnSpc, default leading (matches spcPct 100%)", () => {
    const dyDefault = extractDyValues(
      renderTextBody(makeLineSpacingTextBody(null), NARROW_TRANSFORM),
    );
    const dy100 = extractDyValues(
      renderTextBody(makeLineSpacingTextBody({ type: "pct", value: 100000 }), NARROW_TRANSFORM),
    );

    expect(dyDefault.length).toBeGreaterThan(1);
    expect(dyDefault[1]).toBeCloseTo(dy100[1], 1);
  });
});

describe("text renderer bullet output", () => {
  it("buChar bullet points included in SVG", () => {
    const textBody = makeBulletTextBody([
      { text: "Item 1", bullet: { type: "char", char: "\u2022" } },
      { text: "Item 2", bullet: { type: "char", char: "\u2022" } },
    ]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("\u2022");
    expect(result).toContain("Item 1");
    expect(result).toContain("Item 2");
  });

  it("buAutoNum (arabicPeriod) draws the correct number", () => {
    const textBody = makeBulletTextBody([
      { text: "First", bullet: { type: "autoNum", scheme: "arabicPeriod", startAt: 1 } },
      { text: "Second", bullet: { type: "autoNum", scheme: "arabicPeriod", startAt: 1 } },
      { text: "Third", bullet: { type: "autoNum", scheme: "arabicPeriod", startAt: 1 } },
    ]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("1.");
    expect(result).toContain("2.");
    expect(result).toContain("3.");
  });

  it("If buNone, no symbol is drawn", () => {
    const textBody = makeBulletTextBody([{ text: "No bullet", bullet: { type: "none" } }]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("No bullet");
    // There is no bullet tspan (text tspan only)
    const tspanCount = (result.match(/<tspan/g) ?? []).length;
    expect(tspanCount).toBe(1);
  });

  it("No symbol is drawn if bullet=null (unspecified)", () => {
    const textBody = makeBulletTextBody([{ text: "Plain text", bullet: null }]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("Plain text");
    const tspanCount = (result.match(/<tspan/g) ?? []).length;
    expect(tspanCount).toBe(1);
  });

  it("Indentation by marginLeft is reflected in the x coordinate", () => {
    const noIndent = makeBulletTextBody([
      { text: "No indent", bullet: null, marginLeft: 0, indent: 0 },
    ]);
    const withIndent = makeBulletTextBody([
      { text: "Indented", bullet: null, marginLeft: 457200, indent: 0 },
    ]);

    const resultNoIndent = renderTextBody(noIndent, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    const resultIndent = renderTextBody(withIndent, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));

    // Get x attribute of tspan (instead of x of <text>)
    const xNoIndent = resultNoIndent.match(/<tspan[^>]*x="([^"]+)"/)?.[1];
    const xIndent = resultIndent.match(/<tspan[^>]*x="([^"]+)"/)?.[1];
    expect(Number(xIndent)).toBeGreaterThan(Number(xNoIndent));
  });

  it("startAt of buAutoNum is reflected", () => {
    const textBody = makeBulletTextBody([
      { text: "Item", bullet: { type: "autoNum", scheme: "arabicPeriod", startAt: 5 } },
    ]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("5.");
  });
});

describe("formatAutoNum", () => {
  it("arabicPeriod: 1. 2. 3.", () => {
    expect(formatAutoNum("arabicPeriod", 1)).toBe("1.");
    expect(formatAutoNum("arabicPeriod", 10)).toBe("10.");
  });

  it("arabicParenR: 1) 2) 3)", () => {
    expect(formatAutoNum("arabicParenR", 1)).toBe("1)");
    expect(formatAutoNum("arabicParenR", 3)).toBe("3)");
  });

  it("arabicPlain: 1 2 3", () => {
    expect(formatAutoNum("arabicPlain", 1)).toBe("1");
    expect(formatAutoNum("arabicPlain", 5)).toBe("5");
  });

  it("romanUcPeriod: I. II. III.", () => {
    expect(formatAutoNum("romanUcPeriod", 1)).toBe("I.");
    expect(formatAutoNum("romanUcPeriod", 4)).toBe("IV.");
    expect(formatAutoNum("romanUcPeriod", 9)).toBe("IX.");
  });

  it("romanLcPeriod: i. ii. iii.", () => {
    expect(formatAutoNum("romanLcPeriod", 1)).toBe("i.");
    expect(formatAutoNum("romanLcPeriod", 3)).toBe("iii.");
  });

  it("alphaUcPeriod: A. B. C.", () => {
    expect(formatAutoNum("alphaUcPeriod", 1)).toBe("A.");
    expect(formatAutoNum("alphaUcPeriod", 26)).toBe("Z.");
    expect(formatAutoNum("alphaUcPeriod", 27)).toBe("AA.");
  });

  it("alphaLcPeriod: a. b. c.", () => {
    expect(formatAutoNum("alphaLcPeriod", 1)).toBe("a.");
    expect(formatAutoNum("alphaLcPeriod", 3)).toBe("c.");
  });

  it("alphaUcParenR: A) B) C)", () => {
    expect(formatAutoNum("alphaUcParenR", 1)).toBe("A)");
  });

  it("alphaLcParenR: a) b) c)", () => {
    expect(formatAutoNum("alphaLcParenR", 1)).toBe("a)");
    expect(formatAutoNum("alphaLcParenR", 2)).toBe("b)");
  });
});

describe("text renderer mixed-script font spans", () => {
  it("tspan splits at script boundaries if fontFamily and fontFamilyEa are different", () => {
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "none",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "Hello世界Test",
              properties: {
                fontSize: 18,
                fontFamily: "Calibri",
                fontFamilyEa: "Meiryo",
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
      ],
    };
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    // Fallback list with Calibri at the top for the Latin segment and Meiryo at the top for the EA segment
    // Also includes metrics-compatible OSS fonts
    expect(result).toContain(
      "font-family=\"Calibri, Carlito, Meiryo, 'Noto Sans JP', sans-serif\"",
    );
    expect(result).toContain(
      "font-family=\"Meiryo, 'Noto Sans JP', Calibri, Carlito, sans-serif\"",
    );
    expect(result).toContain("Hello");
    expect(result).toContain("世界");
    expect(result).toContain("Test");
  });

  it("Not split if fontFamily and fontFamilyEa are the same", () => {
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "none",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "Hello世界",
              properties: {
                fontSize: 18,
                fontFamily: "Calibri",
                fontFamilyEa: "Calibri",
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
      ],
    };
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    // There is only one tspan because it is not divided.
    const tspanCount = (result.match(/<tspan/g) ?? []).length;
    expect(tspanCount).toBe(1);
    expect(result).toContain("Hello世界");
  });

  it("If fontFamilyEa is null, it will not be divided by fontFamily only", () => {
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "none",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "Hello世界",
              properties: {
                fontSize: 18,
                fontFamily: "Calibri",
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
      ],
    };
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    const tspanCount = (result.match(/<tspan/g) ?? []).length;
    expect(tspanCount).toBe(1);
    expect(result).toContain("Hello世界");
  });
});

describe("buildFontFamilyValue", () => {
  it("Return single font + metrics fallback + generic family", () => {
    expect(buildFontFamilyValue(["Calibri"])).toBe("Calibri, Carlito, sans-serif");
  });

  it("Include both latin and ea in fallback list", () => {
    expect(buildFontFamilyValue(["Calibri", "Meiryo"])).toBe(
      "Calibri, Carlito, Meiryo, 'Noto Sans JP', sans-serif",
    );
  });

  it("Remove duplicate font names", () => {
    expect(buildFontFamilyValue(["Calibri", "Calibri"])).toBe("Calibri, Carlito, sans-serif");
  });

  it("Works correctly even with lists containing null", () => {
    expect(buildFontFamilyValue(["Calibri", null])).toBe("Calibri, Carlito, sans-serif");
    expect(buildFontFamilyValue([null, "Meiryo"])).toBe("Meiryo, 'Noto Sans JP', sans-serif");
  });

  it("Returns null if all are null", () => {
    expect(buildFontFamilyValue([null, null])).toBeNull();
    expect(buildFontFamilyValue([])).toBeNull();
  });

  it("Enclose font names with spaces in single quotes", () => {
    expect(buildFontFamilyValue(["Times New Roman"])).toBe(
      "'Times New Roman', Tinos, 'Liberation Serif', serif",
    );
    expect(buildFontFamilyValue(["Calibri", "Noto Sans JP"])).toBe(
      "Calibri, Carlito, 'Noto Sans JP', sans-serif",
    );
  });

  it("The general-purpose family of serif fonts becomes serif.", () => {
    expect(buildFontFamilyValue(["Times New Roman"])).toBe(
      "'Times New Roman', Tinos, 'Liberation Serif', serif",
    );
    expect(buildFontFamilyValue(["Yu Mincho"])).toBe(
      "'Yu Mincho', 'Noto Serif CJK JP', 'Noto Sans JP', serif",
    );
    expect(buildFontFamilyValue(["游明朝"])).toBe(
      "游明朝, 'Noto Serif CJK JP', 'Noto Sans JP', serif",
    );
  });

  it("The general-purpose family of sans-serif fonts is now sans-serif.", () => {
    expect(buildFontFamilyValue(["Arial"])).toBe("Arial, Arimo, 'Liberation Sans', sans-serif");
    expect(buildFontFamilyValue(["Meiryo"])).toBe("Meiryo, 'Noto Sans JP', sans-serif");
  });

  it("If the fallback font is the same as the original font, it will not overlap.", () => {
    expect(buildFontFamilyValue(["Noto Sans JP"])).toBe("'Noto Sans JP', sans-serif");
  });

  it("No fallback added for unknown fonts", () => {
    expect(buildFontFamilyValue(["UnknownFont"])).toBe("UnknownFont, sans-serif");
  });
});

describe("text renderer normal autofit", () => {
  it("normAutofit automatically reduces font size when text overflows", () => {
    // Put large text in small text box
    const textBody = makeTextBody(
      ["This is a long text that should overflow the small text box and be shrunk to fit"],
      { fontSize: 36, autoFit: "normAutofit" },
    );
    // Very small shape (about 100x50px)
    const smallTransform = makeTransform(960000, 480000);
    const result = renderTextBody(textBody, smallTransform);
    // The font size should be smaller than 36pt.
    expect(result).not.toContain('font-size="36pt"');
    expect(result).toContain("font-size=");
    // Extract the value of font-size and make sure it is less than 36
    const fontSizeMatch = result.match(/font-size="([0-9.]+)pt"/);
    expect(fontSizeMatch).not.toBeNull();
    expect(Number(fontSizeMatch![1])).toBeLessThan(36);
  });

  it("font size does not change if text fits with normAutofit", () => {
    const textBody = makeTextBody(["Hi"], { fontSize: 12, autoFit: "normAutofit" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('font-size="12pt"');
  });

  it("If noAutofit, the font size will not change even if the text extends", () => {
    const textBody = makeTextBody(
      ["This is a long text that should overflow the small text box but NOT be shrunk"],
      { fontSize: 36, autoFit: "noAutofit" },
    );
    const smallTransform = makeTransform(960000, 480000);
    const result = renderTextBody(textBody, smallTransform);
    expect(result).toContain('font-size="36pt"');
  });

  it("normAutofit scales correctly even if fontSize is not specified in run", () => {
    // Set fontSize to undefined and verify fallback to defaultFontSize
    // Confirm that there is no double scaling by matching the dy value with the case with fontSize specified.
    const longText =
      "This is a long text without explicit fontSize that should be shrunk to fit inside the box";

    const makeBodyWithFontSize = (fontSize: number | undefined): TextBody => ({
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "normAutofit",
        fontScale: 0.8,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: longText,
              properties: {
                fontSize,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
      ],
    });

    const smallTransform = makeTransform(960000, 480000);
    // Case where fontSize=18 (same as default value) is specified
    const resultExplicit = renderTextBody(makeBodyWithFontSize(18), smallTransform);
    // Case of falling back to default value with fontSize=undefined
    const resultImplicit = renderTextBody(makeBodyWithFontSize(undefined), smallTransform);

    // Both dy values must match (same result without double scaling)
    const dyExplicit = resultExplicit.match(/dy="([0-9.]+)"/g);
    const dyImplicit = resultImplicit.match(/dy="([0-9.]+)"/g);
    expect(dyExplicit).not.toBeNull();
    expect(dyImplicit).not.toBeNull();
    expect(dyImplicit).toEqual(dyExplicit);
  });

  it("normAutofit further reduces even if existing fontScale is set", () => {
    const textBody = makeTextBody(
      ["This is a long text that should overflow even with fontScale 0.8 applied"],
      { fontSize: 36, fontScale: 0.8, autoFit: "normAutofit" },
    );
    const smallTransform = makeTransform(960000, 480000);
    const result = renderTextBody(textBody, smallTransform);
    // It should be even smaller than 36 * 0.8 = 28.8
    const fontSizeMatch = result.match(/font-size="([0-9.]+)pt"/);
    expect(fontSizeMatch).not.toBeNull();
    expect(Number(fontSizeMatch![1])).toBeLessThan(28.8);
  });
});

describe("text renderer shape autofit", () => {
  it("Font size does not change when text extends with spAutofit", () => {
    const textBody = makeTextBody(
      ["This is a long text that should overflow the small text box but font size stays the same"],
      { fontSize: 36, autoFit: "spAutofit" },
    );
    const smallTransform = makeTransform(960000, 480000);
    const result = renderTextBody(textBody, smallTransform);
    expect(result).toContain('font-size="36pt"');
  });

  it("spAutofit returns null if text fits", () => {
    const textBody = makeTextBody(["Hi"], { fontSize: 12, autoFit: "spAutofit" });
    const result = computeSpAutofitHeight(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toBeNull();
  });

  it("Returns the required height when text overflows with spAutofit", () => {
    const textBody = makeTextBody(
      ["This is a long text that should overflow the small text box and require more height"],
      { fontSize: 36, autoFit: "spAutofit" },
    );
    const smallTransform = makeTransform(960000, 480000);
    const result = computeSpAutofitHeight(textBody, smallTransform);
    expect(result).not.toBeNull();
    expect(result!).toBeGreaterThan(480000);
  });

  it("Returns null if there is no text", () => {
    const textBody = makeTextBody([""], { autoFit: "spAutofit" });
    const result = computeSpAutofitHeight(textBody, makeTransform(960000, 480000));
    expect(result).toBeNull();
  });
});

describe("text renderer vertical text", () => {
  it("If vert='vert', it will be wrapped in <g> and a rotate(90) transformation will be applied.", () => {
    const textBody = makeTextBody(["Vertical"], { vert: "vert" });
    const result = renderTextBody(textBody, makeTransform(4000000, 2000000));
    expect(result).toContain("<g transform=");
    expect(result).toContain("rotate(90)");
    expect(result).toContain("Vertical");
  });

  it("If vert='vert270', rotate(-90) transformation is applied", () => {
    const textBody = makeTextBody(["Vertical270"], { vert: "vert270" });
    const result = renderTextBody(textBody, makeTransform(4000000, 2000000));
    expect(result).toContain("<g transform=");
    expect(result).toContain("rotate(-90)");
    expect(result).toContain("Vertical270");
  });

  it("If vert='eaVert', rotate(90) is applied like vert", () => {
    const textBody = makeTextBody(["EAVertical"], { vert: "eaVert" });
    const result = renderTextBody(textBody, makeTransform(4000000, 2000000));
    expect(result).toContain("<g transform=");
    expect(result).toContain("rotate(90)");
  });

  it("If vert='horz' (default), no <g> wrapping", () => {
    const textBody = makeTextBody(["Horizontal"]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).not.toContain("<g transform=");
    expect(result).toMatch(/^<text/);
  });

  it("vert='vert' uses the shape's width for translation", () => {
    const widthEmu = 4000000;
    const textBody = makeTextBody(["Test"], { vert: "vert" });
    const result = renderTextBody(textBody, makeTransform(widthEmu, 2000000));
    const widthPx = (widthEmu / 914400) * 96;
    expect(result).toContain(`translate(${widthPx}, 0)`);
  });

  it("vert='vert270' uses shape height for translate", () => {
    const heightEmu = 2000000;
    const textBody = makeTextBody(["Test"], { vert: "vert270" });
    const result = renderTextBody(textBody, makeTransform(4000000, heightEmu));
    const heightPx = (heightEmu / 914400) * 96;
    expect(result).toContain(`translate(0, ${heightPx})`);
  });
});

// ============================================================
// Text -> path conversion (path mode)
// ============================================================

function createMockFont(name: string): OpentypeFullFont {
  return {
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    getPath: (text: string, x: number, y: number, _fontSize: number) => ({
      toPathData: () => `M${x.toFixed(1)} ${y.toFixed(1)}L${name} ${text.length}`,
    }),
    getAdvanceWidth: (text: string, fontSize: number) => text.length * fontSize * 0.6,
  };
}

function setupPathMode(fontName = "TestFont"): void {
  const font = createMockFont(fontName);
  const fonts = new Map([[fontName, font]]);
  setTextPathFontResolver(new DefaultTextPathFontResolver(fonts, font));
}

describe("renderTextBody (path mode)", () => {
  afterEach(() => {
    resetTextPathFontResolver();
  });

  it("Generate <path> element if font resolver is set", () => {
    setupPathMode();
    const textBody = makeTextBody(["Hello"]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("<path");
    expect(result).toContain('d="M');
    expect(result).not.toContain("<text");
    expect(result).not.toContain("<tspan");
  });

  it("The text color is reflected in the fill attribute of path", () => {
    setupPathMode();
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "Red Text",
              properties: {
                fontSize: 18,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: { hex: "#FF0000", alpha: 1 },
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
      ],
    };
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('fill="#FF0000"');
  });

  it("Transparency is reflected in fill-opacity", () => {
    setupPathMode();
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "Transparent",
              properties: {
                fontSize: 18,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: { hex: "#0000FF", alpha: 0.5 },
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
      ],
    };
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('fill="#0000FF"');
    expect(result).toContain('fill-opacity="0.5"');
  });

  it("Hyperlinks are wrapped in <a> tags", () => {
    setupPathMode();
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "Click me",
              properties: {
                fontSize: 18,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
                hyperlink: { url: "https://example.com" },
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
      ],
    };
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('<a href="https://example.com">');
    expect(result).toContain("</a>");
  });

  it("Underline is drawn as a <line> element", () => {
    setupPathMode();
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "Underlined",
              properties: {
                fontSize: 18,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: true,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
      ],
    };
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("<line");
    expect(result).toContain("stroke=");
  });

  it("Strikethrough is drawn as a <line> element", () => {
    setupPathMode();
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "Strikethrough",
              properties: {
                fontSize: 18,
                fontFamily: null,
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: true,
                color: null,
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties(),
        },
      ],
    };
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("<line");
  });

  it("Returns empty string on empty text", () => {
    setupPathMode();
    const textBody = makeTextBody([""]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toBe("");
  });

  it("rotate(90) is applied in vertical writing (vert)", () => {
    setupPathMode();
    const textBody = makeTextBody(["Vertical"], { vert: "vert" });
    const result = renderTextBody(textBody, makeTransform(4000000, 2000000));
    expect(result).toContain("rotate(90)");
    expect(result).toContain("<path");
    expect(result).not.toContain("<text");
  });

  it("rotate(-90) is applied in vertical writing (vert270)", () => {
    setupPathMode();
    const textBody = makeTextBody(["Vert270"], { vert: "vert270" });
    const result = renderTextBody(textBody, makeTransform(4000000, 2000000));
    expect(result).toContain("rotate(-90)");
    expect(result).toContain("<path");
  });

  it("If bulletFont is null, bullets are rendered in the text run's font.", () => {
    const notoFont = createMockFont("Noto Sans JP");
    const defaultFont = createMockFont("DefaultFont");
    const fonts = new Map([
      ["Noto Sans JP", notoFont],
      ["DefaultFont", defaultFont],
    ]);
    setTextPathFontResolver(new DefaultTextPathFontResolver(fonts, defaultFont));

    const textBody: TextBody = {
      bodyProperties: {
        anchor: "t",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          runs: [
            {
              text: "List item",
              properties: {
                fontSize: 18,
                fontFamily: "Noto Sans JP",
                fontFamilyEa: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
                baseline: 0,
              },
            },
          ],
          properties: defaultParagraphProperties({
            bullet: { type: "char", char: "\u2022" },
            bulletFont: null,
            marginLeft: 342900,
            indent: -342900,
          }),
        },
      ],
    };

    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    // Bullet points are drawn as paths (using Noto Sans JP mock font)
    expect(result).toContain("<path");
    expect(result).toContain("Noto Sans JP");
  });

  it("Fallback to tspan rendering if font resolver is null", () => {
    // Do not call setupPathMode() -> fontResolver is null
    const textBody = makeTextBody(["Fallback"]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("<text");
    expect(result).toContain("<tspan");
    expect(result).not.toContain("<path");
  });

  it("Centering in path mode uses x position based on getAdvanceWidth()", () => {
    // Using a font (1.0x) whose width is different from the default mock (0.6x)
    // Make sure measureLineWidth is switched to getAdvanceWidth()
    const wideFont: OpentypeFullFont = {
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      getPath: (_text: string, x: number, y: number, _fontSize: number) => ({
        toPathData: () => `M${x.toFixed(2)} ${y.toFixed(2)}`,
      }),
      getAdvanceWidth: (text: string, fontSize: number) => text.length * fontSize,
    };
    const fonts = new Map([["WideFont", wideFont]]);
    setTextPathFontResolver(new DefaultTextPathFontResolver(fonts, wideFont));

    const textBody = makeTextBody(["Hi"], { alignment: "ctr", fontSize: 18, wrap: "none" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));

    // lineWidth = getAdvanceWidth("Hi", 24px) = 2 * 24 = 48px
    // marginLeftPx = 9.6px, effectiveTextWidth = 940.8px
    // lineStartX = 9.6 + (940.8 - 48) / 2 = 456.0px
    expect(result).toContain("M456.00");
    expect(result).toContain("<path");
  });

  it("Right alignment in path mode uses x position based on getAdvanceWidth()", () => {
    const wideFont: OpentypeFullFont = {
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      getPath: (_text: string, x: number, y: number, _fontSize: number) => ({
        toPathData: () => `M${x.toFixed(2)} ${y.toFixed(2)}`,
      }),
      getAdvanceWidth: (text: string, fontSize: number) => text.length * fontSize,
    };
    const fonts = new Map([["WideFont", wideFont]]);
    setTextPathFontResolver(new DefaultTextPathFontResolver(fonts, wideFont));

    const textBody = makeTextBody(["Hi"], { alignment: "r", fontSize: 18, wrap: "none" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));

    // lineWidth = 2 * 24 = 48px, marginRightPx = 9.6px
    // lineStartX = 960 - 9.6 - 48 = 902.4px
    expect(result).toContain("M902.40");
    expect(result).toContain("<path");
  });

  it("fontSize in endParaRunProperties is used to calculate the height of empty paragraphs", () => {
    const defaultRunProps: RunProperties = {
      fontSize: null,
      fontFamily: null,
      fontFamilyEa: null,
      fontFamilyCs: null,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      color: null,
      baseline: 0,
      hyperlink: null,
      outline: null,
    };
    const textBody: TextBody = {
      bodyProperties: {
        anchor: "ctr",
        marginLeft: 91440,
        marginRight: 91440,
        marginTop: 45720,
        marginBottom: 45720,
        wrap: "square",
        autoFit: "noAutofit",
        fontScale: 1,
        lnSpcReduction: 0,
        numCol: 1,
        vert: "horz",
      },
      paragraphs: [
        {
          // Empty paragraph (spacer): Specify 3pt with endParaRPr
          runs: [{ text: "", properties: { ...defaultRunProps } }],
          properties: defaultParagraphProperties(),
          endParaRunProperties: { ...defaultRunProps, fontSize: 3 },
        },
        {
          // Actual text paragraph: 12pt
          runs: [{ text: "テスト", properties: { ...defaultRunProps, fontSize: 12 } }],
          properties: defaultParagraphProperties(),
        },
      ],
    };

    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    // dy of empty paragraph is calculated based on fontSize (3pt) of endParaRunProperties
    // The value will be different from when using the default (12pt)
    const dyMatches = result.match(/dy="([^"]+)"/g) ?? [];
    // At least two tspans are printed
    expect(dyMatches.length).toBeGreaterThanOrEqual(2);

    // Compare with without endParaRunProperties
    const textBodyNoEndPara: TextBody = {
      ...textBody,
      paragraphs: [
        {
          runs: [{ text: "", properties: { ...defaultRunProps } }],
          properties: defaultParagraphProperties(),
          // endParaRunProperties None -> default size
        },
        {
          runs: [{ text: "テスト", properties: { ...defaultRunProps, fontSize: 12 } }],
          properties: defaultParagraphProperties(),
        },
      ],
    };
    const resultNoEndPara = renderTextBody(
      textBodyNoEndPara,
      makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT),
    );
    // If endParaRunProperties is present, the spacer paragraph is small, so
    // The text position should be different
    expect(result).not.toEqual(resultNoEndPara);
  });
});
