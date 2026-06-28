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
  it("covers text-renderer behavior 1", () => {
    const textBody = makeTextBody([""]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toBe("");
  });

  it("covers text-renderer behavior 2", () => {
    const textBody = makeTextBody(["Hello"]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("<text");
    expect(result).toContain("Hello");
    expect(result).toContain("<tspan");
  });

  it("covers text-renderer behavior 3", () => {
    const textBody = makeTextBody(
      ["The quick brown fox jumps over the lazy dog and continues with more words"],
      { wrap: "square" },
    );
    // Test note.
    const transform = makeTransform(2000000, 2000000); // ~209px width
    const result = renderTextBody(textBody, transform);

    // Test note.
    const xMatches = result.match(/x="/g);
    expect(xMatches).not.toBeNull();
    expect(xMatches!.length).toBeGreaterThan(1);
  });

  it("covers text-renderer behavior 4", () => {
    const textBody = makeTextBody(
      ["The quick brown fox jumps over the lazy dog and continues with more words"],
      { wrap: "none" },
    );
    const transform = makeTransform(2000000, 2000000);
    const result = renderTextBody(textBody, transform);

    // Test note.
    const tspanXMatches = result.match(/<tspan[^>]*x="/g);
    expect(tspanXMatches).toHaveLength(1);
  });

  it("covers text-renderer behavior 5", () => {
    const textBody = makeTextBody(["Center"], { alignment: "ctr" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('text-anchor="middle"');
  });

  it("covers text-renderer behavior 6", () => {
    const textBody = makeTextBody(["Right"], { alignment: "r" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('text-anchor="end"');
  });

  it("covers text-renderer behavior 7", () => {
    const textBody = makeTextBody(["Left"], { alignment: "l" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('text-anchor="start"');
  });

  it("covers text-renderer behavior 8", () => {
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

  it("covers text-renderer behavior 9", () => {
    const textBody = makeTextBody(["Test"], { fontSize: 24 });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('font-size="24pt"');
  });

  it("covers text-renderer behavior 10", () => {
    const textBody = makeTextBody(["Test"], { fontSize: 24, fontScale: 0.625 });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    // 24 * 0.625 = 15
    expect(result).toContain('font-size="15pt"');
  });

  it("covers text-renderer behavior 11", () => {
    const textBody = makeTextBody(["Test"], { fontSize: 24, fontScale: 1 });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('font-size="24pt"');
  });

  it("covers text-renderer behavior 12", () => {
    const textBody = makeTextBody(["本日は晴天なり今日もいい天気です素晴らしい一日"], {
      wrap: "square",
    });
    const transform = makeTransform(2000000, 2000000);
    const result = renderTextBody(textBody, transform);
    const xMatches = result.match(/<tspan[^>]*x="/g);
    expect(xMatches).not.toBeNull();
    expect(xMatches!.length).toBeGreaterThan(1);
  });

  it("covers text-renderer behavior 13", () => {
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

  it("covers text-renderer behavior 14", () => {
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

  it("covers text-renderer behavior 15", () => {
    const textBody = makeTextBody(["Normal"]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).not.toContain("baseline-shift");
  });
});

describe("renderer/text-renderer.test behavior", () => {
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

  it("covers text-renderer behavior 16", () => {
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

    // Test note.
    expect(dyWith[1]).toBeGreaterThan(dyWithout[1]);
  });

  it("covers text-renderer behavior 17", () => {
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

    // Test note.
    expect(dyWith[1]).toBeGreaterThan(dyWithout[1]);
  });

  it("covers text-renderer behavior 18", () => {
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

    // Test note.
    expect(dyAfter[1]).toBeCloseTo(dyBefore[1], 1);
  });

  it("covers text-renderer behavior 19", () => {
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

    // Test note.
    expect(dyWith[1]).toBeGreaterThan(dyWithout[1]);
    // Test note.
    const diff = dyWith[1] - dyWithout[1];
    expect(diff).toBeCloseTo(18 * (96 / 72), 1);
  });
});

describe("renderer/text-renderer.test behavior", () => {
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

  // Test note.
  const NARROW_TRANSFORM = makeTransform(2000000, 2000000);

  it("covers text-renderer behavior 20", () => {
    // Test note.
    const textBody = makeLineSpacingTextBody({ type: "pts", value: 2100 });
    const dyValues = extractDyValues(renderTextBody(textBody, NARROW_TRANSFORM));

    expect(dyValues.length).toBeGreaterThan(1);
    for (const dy of dyValues.slice(1)) {
      expect(dy).toBeCloseTo(28, 1);
    }
  });

  it("covers text-renderer behavior 21", () => {
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

  it("covers text-renderer behavior 22", () => {
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

    // Test note.
    expect(dyValues).toHaveLength(3);
    expect(dyValues[1]).toBeCloseTo(28, 1);
  });

  it("covers text-renderer behavior 23", () => {
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

describe("renderer/text-renderer.test behavior", () => {
  it("covers text-renderer behavior 24", () => {
    const textBody = makeBulletTextBody([
      { text: "Item 1", bullet: { type: "char", char: "\u2022" } },
      { text: "Item 2", bullet: { type: "char", char: "\u2022" } },
    ]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("\u2022");
    expect(result).toContain("Item 1");
    expect(result).toContain("Item 2");
  });

  it("covers text-renderer behavior 25", () => {
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

  it("covers text-renderer behavior 26", () => {
    const textBody = makeBulletTextBody([{ text: "No bullet", bullet: { type: "none" } }]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("No bullet");
    // Test note.
    const tspanCount = (result.match(/<tspan/g) ?? []).length;
    expect(tspanCount).toBe(1);
  });

  it("covers text-renderer behavior 27", () => {
    const textBody = makeBulletTextBody([{ text: "Plain text", bullet: null }]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("Plain text");
    const tspanCount = (result.match(/<tspan/g) ?? []).length;
    expect(tspanCount).toBe(1);
  });

  it("covers text-renderer behavior 28", () => {
    const noIndent = makeBulletTextBody([
      { text: "No indent", bullet: null, marginLeft: 0, indent: 0 },
    ]);
    const withIndent = makeBulletTextBody([
      { text: "Indented", bullet: null, marginLeft: 457200, indent: 0 },
    ]);

    const resultNoIndent = renderTextBody(noIndent, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    const resultIndent = renderTextBody(withIndent, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));

    // Test note.
    const xNoIndent = resultNoIndent.match(/<tspan[^>]*x="([^"]+)"/)?.[1];
    const xIndent = resultIndent.match(/<tspan[^>]*x="([^"]+)"/)?.[1];
    expect(Number(xIndent)).toBeGreaterThan(Number(xNoIndent));
  });

  it("covers text-renderer behavior 29", () => {
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

describe("renderer/text-renderer.test behavior", () => {
  it("covers text-renderer behavior 30", () => {
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
    // Test note.
    // Test note.
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

  it("covers text-renderer behavior 31", () => {
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
    // Test note.
    const tspanCount = (result.match(/<tspan/g) ?? []).length;
    expect(tspanCount).toBe(1);
    expect(result).toContain("Hello世界");
  });

  it("covers text-renderer behavior 32", () => {
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
  it("covers text-renderer behavior 33", () => {
    expect(buildFontFamilyValue(["Calibri"])).toBe("Calibri, Carlito, sans-serif");
  });

  it("covers text-renderer behavior 34", () => {
    expect(buildFontFamilyValue(["Calibri", "Meiryo"])).toBe(
      "Calibri, Carlito, Meiryo, 'Noto Sans JP', sans-serif",
    );
  });

  it("covers text-renderer behavior 35", () => {
    expect(buildFontFamilyValue(["Calibri", "Calibri"])).toBe("Calibri, Carlito, sans-serif");
  });

  it("covers text-renderer behavior 36", () => {
    expect(buildFontFamilyValue(["Calibri", null])).toBe("Calibri, Carlito, sans-serif");
    expect(buildFontFamilyValue([null, "Meiryo"])).toBe("Meiryo, 'Noto Sans JP', sans-serif");
  });

  it("covers text-renderer behavior 37", () => {
    expect(buildFontFamilyValue([null, null])).toBeNull();
    expect(buildFontFamilyValue([])).toBeNull();
  });

  it("covers text-renderer behavior 38", () => {
    expect(buildFontFamilyValue(["Times New Roman"])).toBe(
      "'Times New Roman', Tinos, 'Liberation Serif', serif",
    );
    expect(buildFontFamilyValue(["Calibri", "Noto Sans JP"])).toBe(
      "Calibri, Carlito, 'Noto Sans JP', sans-serif",
    );
  });

  it("covers text-renderer behavior 39", () => {
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

  it("covers text-renderer behavior 40", () => {
    expect(buildFontFamilyValue(["Arial"])).toBe("Arial, Arimo, 'Liberation Sans', sans-serif");
    expect(buildFontFamilyValue(["Meiryo"])).toBe("Meiryo, 'Noto Sans JP', sans-serif");
  });

  it("covers text-renderer behavior 41", () => {
    expect(buildFontFamilyValue(["Noto Sans JP"])).toBe("'Noto Sans JP', sans-serif");
  });

  it("covers text-renderer behavior 42", () => {
    expect(buildFontFamilyValue(["UnknownFont"])).toBe("UnknownFont, sans-serif");
  });
});

describe("renderer/text-renderer.test behavior", () => {
  it("covers text-renderer behavior 43", () => {
    // Test note.
    const textBody = makeTextBody(
      ["This is a long text that should overflow the small text box and be shrunk to fit"],
      { fontSize: 36, autoFit: "normAutofit" },
    );
    // Test note.
    const smallTransform = makeTransform(960000, 480000);
    const result = renderTextBody(textBody, smallTransform);
    // Test note.
    expect(result).not.toContain('font-size="36pt"');
    expect(result).toContain("font-size=");
    // Test note.
    const fontSizeMatch = result.match(/font-size="([0-9.]+)pt"/);
    expect(fontSizeMatch).not.toBeNull();
    expect(Number(fontSizeMatch![1])).toBeLessThan(36);
  });

  it("covers text-renderer behavior 44", () => {
    const textBody = makeTextBody(["Hi"], { fontSize: 12, autoFit: "normAutofit" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('font-size="12pt"');
  });

  it("covers text-renderer behavior 45", () => {
    const textBody = makeTextBody(
      ["This is a long text that should overflow the small text box but NOT be shrunk"],
      { fontSize: 36, autoFit: "noAutofit" },
    );
    const smallTransform = makeTransform(960000, 480000);
    const result = renderTextBody(textBody, smallTransform);
    expect(result).toContain('font-size="36pt"');
  });

  it("covers text-renderer behavior 46", () => {
    // Test note.
    // Test note.
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
    // Test note.
    const resultExplicit = renderTextBody(makeBodyWithFontSize(18), smallTransform);
    // Test note.
    const resultImplicit = renderTextBody(makeBodyWithFontSize(undefined), smallTransform);

    // Test note.
    const dyExplicit = resultExplicit.match(/dy="([0-9.]+)"/g);
    const dyImplicit = resultImplicit.match(/dy="([0-9.]+)"/g);
    expect(dyExplicit).not.toBeNull();
    expect(dyImplicit).not.toBeNull();
    expect(dyImplicit).toEqual(dyExplicit);
  });

  it("covers text-renderer behavior 47", () => {
    const textBody = makeTextBody(
      ["This is a long text that should overflow even with fontScale 0.8 applied"],
      { fontSize: 36, fontScale: 0.8, autoFit: "normAutofit" },
    );
    const smallTransform = makeTransform(960000, 480000);
    const result = renderTextBody(textBody, smallTransform);
    // Test note.
    const fontSizeMatch = result.match(/font-size="([0-9.]+)pt"/);
    expect(fontSizeMatch).not.toBeNull();
    expect(Number(fontSizeMatch![1])).toBeLessThan(28.8);
  });
});

describe("renderer/text-renderer.test behavior", () => {
  it("covers text-renderer behavior 48", () => {
    const textBody = makeTextBody(
      ["This is a long text that should overflow the small text box but font size stays the same"],
      { fontSize: 36, autoFit: "spAutofit" },
    );
    const smallTransform = makeTransform(960000, 480000);
    const result = renderTextBody(textBody, smallTransform);
    expect(result).toContain('font-size="36pt"');
  });

  it("covers text-renderer behavior 49", () => {
    const textBody = makeTextBody(["Hi"], { fontSize: 12, autoFit: "spAutofit" });
    const result = computeSpAutofitHeight(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toBeNull();
  });

  it("covers text-renderer behavior 50", () => {
    const textBody = makeTextBody(
      ["This is a long text that should overflow the small text box and require more height"],
      { fontSize: 36, autoFit: "spAutofit" },
    );
    const smallTransform = makeTransform(960000, 480000);
    const result = computeSpAutofitHeight(textBody, smallTransform);
    expect(result).not.toBeNull();
    expect(result!).toBeGreaterThan(480000);
  });

  it("covers text-renderer behavior 51", () => {
    const textBody = makeTextBody([""], { autoFit: "spAutofit" });
    const result = computeSpAutofitHeight(textBody, makeTransform(960000, 480000));
    expect(result).toBeNull();
  });
});

describe("renderer/text-renderer.test behavior", () => {
  it("covers text-renderer behavior 52", () => {
    const textBody = makeTextBody(["Vertical"], { vert: "vert" });
    const result = renderTextBody(textBody, makeTransform(4000000, 2000000));
    expect(result).toContain("<g transform=");
    expect(result).toContain("rotate(90)");
    expect(result).toContain("Vertical");
  });

  it("covers text-renderer behavior 53", () => {
    const textBody = makeTextBody(["Vertical270"], { vert: "vert270" });
    const result = renderTextBody(textBody, makeTransform(4000000, 2000000));
    expect(result).toContain("<g transform=");
    expect(result).toContain("rotate(-90)");
    expect(result).toContain("Vertical270");
  });

  it("covers text-renderer behavior 54", () => {
    const textBody = makeTextBody(["EAVertical"], { vert: "eaVert" });
    const result = renderTextBody(textBody, makeTransform(4000000, 2000000));
    expect(result).toContain("<g transform=");
    expect(result).toContain("rotate(90)");
  });

  it("covers text-renderer behavior 55", () => {
    const textBody = makeTextBody(["Horizontal"]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).not.toContain("<g transform=");
    expect(result).toMatch(/^<text/);
  });

  it("covers text-renderer behavior 56", () => {
    const widthEmu = 4000000;
    const textBody = makeTextBody(["Test"], { vert: "vert" });
    const result = renderTextBody(textBody, makeTransform(widthEmu, 2000000));
    const widthPx = (widthEmu / 914400) * 96;
    expect(result).toContain(`translate(${widthPx}, 0)`);
  });

  it("covers text-renderer behavior 57", () => {
    const heightEmu = 2000000;
    const textBody = makeTextBody(["Test"], { vert: "vert270" });
    const result = renderTextBody(textBody, makeTransform(4000000, heightEmu));
    const heightPx = (heightEmu / 914400) * 96;
    expect(result).toContain(`translate(0, ${heightPx})`);
  });
});

// ============================================================
// Test note.
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

  it("covers text-renderer behavior 58", () => {
    setupPathMode();
    const textBody = makeTextBody(["Hello"]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("<path");
    expect(result).toContain('d="M');
    expect(result).not.toContain("<text");
    expect(result).not.toContain("<tspan");
  });

  it("covers text-renderer behavior 59", () => {
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

  it("covers text-renderer behavior 60", () => {
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

  it("covers text-renderer behavior 61", () => {
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

  it("covers text-renderer behavior 62", () => {
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

  it("covers text-renderer behavior 63", () => {
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

  it("covers text-renderer behavior 64", () => {
    setupPathMode();
    const textBody = makeTextBody([""]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toBe("");
  });

  it("covers text-renderer behavior 65", () => {
    setupPathMode();
    const textBody = makeTextBody(["Vertical"], { vert: "vert" });
    const result = renderTextBody(textBody, makeTransform(4000000, 2000000));
    expect(result).toContain("rotate(90)");
    expect(result).toContain("<path");
    expect(result).not.toContain("<text");
  });

  it("covers text-renderer behavior 66", () => {
    setupPathMode();
    const textBody = makeTextBody(["Vert270"], { vert: "vert270" });
    const result = renderTextBody(textBody, makeTransform(4000000, 2000000));
    expect(result).toContain("rotate(-90)");
    expect(result).toContain("<path");
  });

  it("covers text-renderer behavior 67", () => {
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
    // Test note.
    expect(result).toContain("<path");
    expect(result).toContain("Noto Sans JP");
  });

  it("covers text-renderer behavior 68", () => {
    // Test note.
    const textBody = makeTextBody(["Fallback"]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("<text");
    expect(result).toContain("<tspan");
    expect(result).not.toContain("<path");
  });

  it("covers text-renderer behavior 69", () => {
    // Test note.
    // Test note.
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

  it("covers text-renderer behavior 70", () => {
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

  it("covers text-renderer behavior 71", () => {
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
          // Test note.
          runs: [{ text: "", properties: { ...defaultRunProps } }],
          properties: defaultParagraphProperties(),
          endParaRunProperties: { ...defaultRunProps, fontSize: 3 },
        },
        {
          // Test note.
          runs: [{ text: "テスト", properties: { ...defaultRunProps, fontSize: 12 } }],
          properties: defaultParagraphProperties(),
        },
      ],
    };

    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    // Test note.
    // Test note.
    const dyMatches = result.match(/dy="([^"]+)"/g) ?? [];
    // Test note.
    expect(dyMatches.length).toBeGreaterThanOrEqual(2);

    // Test note.
    const textBodyNoEndPara: TextBody = {
      ...textBody,
      paragraphs: [
        {
          runs: [{ text: "", properties: { ...defaultRunProps } }],
          properties: defaultParagraphProperties(),
          // Test note.
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
    // Test note.
    // Test note.
    expect(result).not.toEqual(resultNoEndPara);
  });
});
