import { describe, it, expect } from "vitest";
import { renderTextBody, formatAutoNum } from "../../src/renderer/text-renderer.js";
import type { TextBody, Paragraph, BulletType } from "../../src/model/text.js";
import type { Transform } from "../../src/model/shape.js";

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
    spaceBefore: 0,
    spaceAfter: 0,
    level: 0,
    bullet: null,
    bulletFont: null,
    bulletColor: null,
    bulletSizePct: null,
    marginLeft: 0,
    indent: 0,
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
  },
): TextBody {
  return {
    bodyProperties: {
      anchor: overrides?.anchor ?? "t",
      marginLeft: 91440, // ~9.6px
      marginRight: 91440,
      marginTop: 45720, // ~4.8px
      marginBottom: 45720,
      wrap: overrides?.wrap ?? "square",
      autoFit: overrides?.fontScale !== undefined ? "normAutofit" : "noAutofit",
      fontScale: overrides?.fontScale ?? 1,
      lnSpcReduction: overrides?.lnSpcReduction ?? 0,
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

// 16:9 スライドサイズ (EMU)
const SLIDE_WIDTH = 9144000;
const SLIDE_HEIGHT = 5143500;

describe("renderTextBody", () => {
  it("テキストがない場合は空文字列を返す", () => {
    const textBody = makeTextBody([""]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toBe("");
  });

  it("短いテキストを正しくレンダリングする", () => {
    const textBody = makeTextBody(["Hello"]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("<text");
    expect(result).toContain("Hello");
    expect(result).toContain("<tspan");
  });

  it("wrap=square の場合、長いテキストが複数の tspan に折り返される", () => {
    const textBody = makeTextBody(
      ["The quick brown fox jumps over the lazy dog and continues with more words"],
      { wrap: "square" },
    );
    // 狭い幅のテキストボックス
    const transform = makeTransform(2000000, 2000000); // ~209px width
    const result = renderTextBody(textBody, transform);

    // 複数の x 属性を持つ tspan が存在する (= 複数行)
    const xMatches = result.match(/x="/g);
    expect(xMatches).not.toBeNull();
    expect(xMatches!.length).toBeGreaterThan(1);
  });

  it("wrap=none の場合、テキストは折り返されない", () => {
    const textBody = makeTextBody(
      ["The quick brown fox jumps over the lazy dog and continues with more words"],
      { wrap: "none" },
    );
    const transform = makeTransform(2000000, 2000000);
    const result = renderTextBody(textBody, transform);

    // x 属性を持つ tspan は1つだけ (text 要素の x を除く)
    const tspanXMatches = result.match(/<tspan[^>]*x="/g);
    expect(tspanXMatches).toHaveLength(1);
  });

  it("中央揃えの場合 text-anchor=middle が設定される", () => {
    const textBody = makeTextBody(["Center"], { alignment: "ctr" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('text-anchor="middle"');
  });

  it("右揃えの場合 text-anchor=end が設定される", () => {
    const textBody = makeTextBody(["Right"], { alignment: "r" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('text-anchor="end"');
  });

  it("左揃えの場合 text-anchor=start が設定される", () => {
    const textBody = makeTextBody(["Left"], { alignment: "l" });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('text-anchor="start"');
  });

  it("複数段落を正しくレンダリングする", () => {
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

  it("font-size 属性が正しく設定される", () => {
    const textBody = makeTextBody(["Test"], { fontSize: 24 });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('font-size="24pt"');
  });

  it("fontScale が適用されるとフォントサイズが縮小される", () => {
    const textBody = makeTextBody(["Test"], { fontSize: 24, fontScale: 0.625 });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    // 24 * 0.625 = 15
    expect(result).toContain('font-size="15pt"');
  });

  it("fontScale=1 の場合はフォントサイズが変わらない", () => {
    const textBody = makeTextBody(["Test"], { fontSize: 24, fontScale: 1 });
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain('font-size="24pt"');
  });

  it("CJK テキストの折り返しが動作する", () => {
    const textBody = makeTextBody(["本日は晴天なり今日もいい天気です素晴らしい一日"], {
      wrap: "square",
    });
    const transform = makeTransform(2000000, 2000000);
    const result = renderTextBody(textBody, transform);
    const xMatches = result.match(/<tspan[^>]*x="/g);
    expect(xMatches).not.toBeNull();
    expect(xMatches!.length).toBeGreaterThan(1);
  });

  it("上付き文字に baseline-shift=super が設定される", () => {
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

  it("下付き文字に baseline-shift=sub が設定される", () => {
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

  it("baseline=0 の場合は baseline-shift が設定されない", () => {
    const textBody = makeTextBody(["Normal"]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).not.toContain("baseline-shift");
  });
});

describe("箇条書き記号レンダリング", () => {
  it("buChar の箇条書き記号が SVG に含まれる", () => {
    const textBody = makeBulletTextBody([
      { text: "Item 1", bullet: { type: "char", char: "\u2022" } },
      { text: "Item 2", bullet: { type: "char", char: "\u2022" } },
    ]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("\u2022");
    expect(result).toContain("Item 1");
    expect(result).toContain("Item 2");
  });

  it("buAutoNum (arabicPeriod) で正しい番号が描画される", () => {
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

  it("buNone の場合は記号が描画されない", () => {
    const textBody = makeBulletTextBody([{ text: "No bullet", bullet: { type: "none" } }]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("No bullet");
    // bullet tspan が1つも入らない（テキスト用 tspan のみ）
    const tspanCount = (result.match(/<tspan/g) ?? []).length;
    expect(tspanCount).toBe(1);
  });

  it("bullet=null (未指定) の場合は記号が描画されない", () => {
    const textBody = makeBulletTextBody([{ text: "Plain text", bullet: null }]);
    const result = renderTextBody(textBody, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    expect(result).toContain("Plain text");
    const tspanCount = (result.match(/<tspan/g) ?? []).length;
    expect(tspanCount).toBe(1);
  });

  it("marginLeft によるインデントが x 座標に反映される", () => {
    const noIndent = makeBulletTextBody([
      { text: "No indent", bullet: null, marginLeft: 0, indent: 0 },
    ]);
    const withIndent = makeBulletTextBody([
      { text: "Indented", bullet: null, marginLeft: 457200, indent: 0 },
    ]);

    const resultNoIndent = renderTextBody(noIndent, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));
    const resultIndent = renderTextBody(withIndent, makeTransform(SLIDE_WIDTH, SLIDE_HEIGHT));

    // tspan の x 属性を取得（<text> の x ではなく）
    const xNoIndent = resultNoIndent.match(/<tspan[^>]*x="([^"]+)"/)?.[1];
    const xIndent = resultIndent.match(/<tspan[^>]*x="([^"]+)"/)?.[1];
    expect(Number(xIndent)).toBeGreaterThan(Number(xNoIndent));
  });

  it("buAutoNum の startAt が反映される", () => {
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

describe("latin/ea フォント切り替え", () => {
  it("fontFamily と fontFamilyEa が異なる場合、スクリプト境界で tspan が分割される", () => {
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
    expect(result).toContain('font-family="Calibri"');
    expect(result).toContain('font-family="Meiryo"');
    expect(result).toContain("Hello");
    expect(result).toContain("世界");
    expect(result).toContain("Test");
  });

  it("fontFamily と fontFamilyEa が同じ場合は分割されない", () => {
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
    // 分割されないので tspan は1つだけ
    const tspanCount = (result.match(/<tspan/g) ?? []).length;
    expect(tspanCount).toBe(1);
    expect(result).toContain("Hello世界");
  });

  it("fontFamilyEa が null の場合は fontFamily のみで分割されない", () => {
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
