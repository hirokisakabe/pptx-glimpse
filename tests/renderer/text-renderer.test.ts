import { describe, it, expect } from "vitest";
import { renderTextBody } from "../../src/renderer/text-renderer.js";
import type { TextBody } from "../../src/model/text.js";
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

function makeTextBody(
  texts: string[],
  overrides?: {
    wrap?: "square" | "none";
    anchor?: "t" | "ctr" | "b";
    alignment?: "l" | "ctr" | "r" | "just";
    fontSize?: number;
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
    },
    paragraphs: [
      {
        runs: texts.map((text) => ({
          text,
          properties: {
            fontSize: overrides?.fontSize ?? 18,
            fontFamily: null,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            color: null,
          },
        })),
        properties: {
          alignment: overrides?.alignment ?? "l",
          lineSpacing: null,
          spaceBefore: 0,
          spaceAfter: 0,
          level: 0,
        },
      },
    ],
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
      },
      paragraphs: [
        {
          runs: [
            {
              text: "First",
              properties: {
                fontSize: 18,
                fontFamily: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
              },
            },
          ],
          properties: {
            alignment: "l",
            lineSpacing: null,
            spaceBefore: 0,
            spaceAfter: 0,
            level: 0,
          },
        },
        {
          runs: [
            {
              text: "Second",
              properties: {
                fontSize: 18,
                fontFamily: null,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                color: null,
              },
            },
          ],
          properties: {
            alignment: "l",
            lineSpacing: null,
            spaceBefore: 0,
            spaceAfter: 0,
            level: 0,
          },
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
});
