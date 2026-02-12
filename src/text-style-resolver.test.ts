import { describe, it, expect } from "vitest";
import { applyTextStyleInheritance } from "./text-style-resolver.js";
import type { TextStyleContext } from "./text-style-resolver.js";
import type { ShapeElement, GroupElement } from "./model/shape.js";
import type { DefaultTextStyle, RunProperties } from "./model/text.js";

function makeRunProperties(overrides: Partial<RunProperties> = {}): RunProperties {
  return {
    fontSize: null,
    fontFamily: null,
    fontFamilyEa: null,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    color: null,
    baseline: 0,
    hyperlink: null,
    ...overrides,
  };
}

function makeShape(overrides: Partial<ShapeElement> = {}): ShapeElement {
  return {
    type: "shape",
    transform: {
      offsetX: 0,
      offsetY: 0,
      extentWidth: 9144000,
      extentHeight: 5143500,
      rotation: 0,
      flipH: false,
      flipV: false,
    },
    geometry: { type: "preset", preset: "rect", adjustValues: {} },
    fill: null,
    outline: null,
    textBody: {
      paragraphs: [
        {
          runs: [{ text: "test", properties: makeRunProperties() }],
          properties: {
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
          },
        },
      ],
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
    },
    effects: null,
    ...overrides,
  };
}

function makeContext(overrides: Partial<TextStyleContext> = {}): TextStyleContext {
  return {
    layoutPlaceholderStyles: [],
    masterPlaceholderStyles: [],
    ...overrides,
  };
}

function makeDefaultTextStyle(
  levelDefRPr: Partial<{
    fontSize: number;
    fontFamily: string;
    fontFamilyEa: string;
  }>,
  level = 0,
): DefaultTextStyle {
  const levels: (undefined | { defaultRunProperties: typeof levelDefRPr })[] =
    Array(9).fill(undefined);
  levels[level] = { defaultRunProperties: levelDefRPr };
  return { levels };
}

describe("applyTextStyleInheritance", () => {
  describe("基本的な継承チェーン", () => {
    it("fontSize が null の場合、defaultTextStyle から解決される", () => {
      const shape = makeShape();
      const context = makeContext({
        defaultTextStyle: makeDefaultTextStyle({ fontSize: 24 }),
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(24);
    });

    it("fontFamily が null の場合、defaultTextStyle から解決される", () => {
      const shape = makeShape();
      const context = makeContext({
        defaultTextStyle: makeDefaultTextStyle({ fontFamily: "Arial" }),
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontFamily).toBe("Arial");
    });

    it("fontFamilyEa が null の場合、defaultTextStyle から解決される", () => {
      const shape = makeShape();
      const context = makeContext({
        defaultTextStyle: makeDefaultTextStyle({ fontFamilyEa: "MS Gothic" }),
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontFamilyEa).toBe("MS Gothic");
    });

    it("既に値が設定されている場合は上書きしない", () => {
      const shape = makeShape();
      shape.textBody!.paragraphs[0].runs[0].properties.fontSize = 16;
      shape.textBody!.paragraphs[0].runs[0].properties.fontFamily = "Calibri";
      const context = makeContext({
        defaultTextStyle: makeDefaultTextStyle({ fontSize: 24, fontFamily: "Arial" }),
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(16);
      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontFamily).toBe("Calibri");
    });
  });

  describe("フォールバック順序", () => {
    it("layout → master → txStyles → defaultTextStyle の順で解決される", () => {
      const shape = makeShape({ placeholderType: "body" });
      const context = makeContext({
        layoutPlaceholderStyles: [
          { placeholderType: "body", lstStyle: makeDefaultTextStyle({ fontSize: 20 }) },
        ],
        masterPlaceholderStyles: [
          { placeholderType: "body", lstStyle: makeDefaultTextStyle({ fontSize: 24 }) },
        ],
        txStyles: { bodyStyle: makeDefaultTextStyle({ fontSize: 28 }) },
        defaultTextStyle: makeDefaultTextStyle({ fontSize: 32 }),
      });

      applyTextStyleInheritance([shape], context);

      // layout の値が最優先
      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(20);
    });

    it("layout にない場合は master にフォールバック", () => {
      const shape = makeShape({ placeholderType: "body" });
      const context = makeContext({
        layoutPlaceholderStyles: [
          { placeholderType: "body", lstStyle: makeDefaultTextStyle({ fontFamily: "Arial" }) },
        ],
        masterPlaceholderStyles: [
          { placeholderType: "body", lstStyle: makeDefaultTextStyle({ fontSize: 24 }) },
        ],
        txStyles: { bodyStyle: makeDefaultTextStyle({ fontSize: 28 }) },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(24);
      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontFamily).toBe("Arial");
    });

    it("layout/master にない場合は txStyles にフォールバック", () => {
      const shape = makeShape({ placeholderType: "body" });
      const context = makeContext({
        layoutPlaceholderStyles: [],
        masterPlaceholderStyles: [],
        txStyles: { bodyStyle: makeDefaultTextStyle({ fontSize: 28 }) },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(28);
    });

    it("layout/master/txStyles にない場合は defaultTextStyle にフォールバック", () => {
      const shape = makeShape({ placeholderType: "body" });
      const context = makeContext({
        defaultTextStyle: makeDefaultTextStyle({ fontSize: 14 }),
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(14);
    });
  });

  describe("プレースホルダーマッチング", () => {
    it("placeholderIdx で優先的にマッチする", () => {
      const shape = makeShape({ placeholderType: "body", placeholderIdx: 2 });
      const context = makeContext({
        layoutPlaceholderStyles: [
          {
            placeholderType: "body",
            placeholderIdx: 1,
            lstStyle: makeDefaultTextStyle({ fontSize: 20 }),
          },
          {
            placeholderType: "body",
            placeholderIdx: 2,
            lstStyle: makeDefaultTextStyle({ fontSize: 24 }),
          },
        ],
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(24);
    });

    it("idx がマッチしない場合は type でフォールバック", () => {
      const shape = makeShape({ placeholderType: "body", placeholderIdx: 99 });
      const context = makeContext({
        layoutPlaceholderStyles: [
          {
            placeholderType: "body",
            placeholderIdx: 1,
            lstStyle: makeDefaultTextStyle({ fontSize: 20 }),
          },
        ],
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(20);
    });
  });

  describe("txStyles マッピング", () => {
    it("title プレースホルダーは titleStyle を使う", () => {
      const shape = makeShape({ placeholderType: "title" });
      const context = makeContext({
        txStyles: {
          titleStyle: makeDefaultTextStyle({ fontSize: 44 }),
          bodyStyle: makeDefaultTextStyle({ fontSize: 32 }),
          otherStyle: makeDefaultTextStyle({ fontSize: 18 }),
        },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(44);
    });

    it("ctrTitle プレースホルダーは titleStyle を使う", () => {
      const shape = makeShape({ placeholderType: "ctrTitle" });
      const context = makeContext({
        txStyles: {
          titleStyle: makeDefaultTextStyle({ fontSize: 44 }),
          bodyStyle: makeDefaultTextStyle({ fontSize: 32 }),
        },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(44);
    });

    it("body プレースホルダーは bodyStyle を使う", () => {
      const shape = makeShape({ placeholderType: "body" });
      const context = makeContext({
        txStyles: {
          titleStyle: makeDefaultTextStyle({ fontSize: 44 }),
          bodyStyle: makeDefaultTextStyle({ fontSize: 32 }),
        },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(32);
    });

    it("subTitle プレースホルダーは bodyStyle を使う", () => {
      const shape = makeShape({ placeholderType: "subTitle" });
      const context = makeContext({
        txStyles: { bodyStyle: makeDefaultTextStyle({ fontSize: 32 }) },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(32);
    });

    it("obj プレースホルダーは bodyStyle を使う", () => {
      const shape = makeShape({ placeholderType: "obj" });
      const context = makeContext({
        txStyles: { bodyStyle: makeDefaultTextStyle({ fontSize: 32 }) },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(32);
    });

    it("その他のプレースホルダーは otherStyle を使う", () => {
      const shape = makeShape({ placeholderType: "sldNum" });
      const context = makeContext({
        txStyles: {
          titleStyle: makeDefaultTextStyle({ fontSize: 44 }),
          otherStyle: makeDefaultTextStyle({ fontSize: 10 }),
        },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(10);
    });

    it("非プレースホルダーシェイプは otherStyle を使う", () => {
      const shape = makeShape();
      const context = makeContext({
        txStyles: {
          titleStyle: makeDefaultTextStyle({ fontSize: 44 }),
          otherStyle: makeDefaultTextStyle({ fontSize: 18 }),
        },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(18);
    });
  });

  describe("テーマフォント解決", () => {
    it("+mj-lt がテーマの majorFont に解決される", () => {
      const shape = makeShape({ placeholderType: "title" });
      const context = makeContext({
        txStyles: {
          titleStyle: makeDefaultTextStyle({ fontFamily: "+mj-lt" }),
        },
        fontScheme: {
          majorFont: "Calibri Light",
          minorFont: "Calibri",
          majorFontEa: null,
          minorFontEa: null,
        },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontFamily).toBe("Calibri Light");
    });

    it("+mn-ea がテーマの minorFontEa に解決される", () => {
      const shape = makeShape();
      const context = makeContext({
        defaultTextStyle: makeDefaultTextStyle({ fontFamilyEa: "+mn-ea" }),
        fontScheme: {
          majorFont: "Calibri Light",
          minorFont: "Calibri",
          majorFontEa: "MS Gothic",
          minorFontEa: "MS PGothic",
        },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontFamilyEa).toBe("MS PGothic");
    });
  });

  describe("レベル対応", () => {
    it("段落レベルに応じた正しいスタイルが適用される", () => {
      const shape = makeShape({ placeholderType: "body" });
      const levels = Array(9).fill(undefined) as (
        | undefined
        | { defaultRunProperties: { fontSize: number } }
      )[];
      levels[0] = { defaultRunProperties: { fontSize: 32 } };
      levels[1] = { defaultRunProperties: { fontSize: 28 } };
      levels[2] = { defaultRunProperties: { fontSize: 24 } };

      // レベル1の段落を追加
      shape.textBody!.paragraphs.push({
        runs: [{ text: "level 1", properties: makeRunProperties() }],
        properties: {
          alignment: "l",
          lineSpacing: null,
          spaceBefore: 0,
          spaceAfter: 0,
          level: 1,
          bullet: null,
          bulletFont: null,
          bulletColor: null,
          bulletSizePct: null,
          marginLeft: 0,
          indent: 0,
        },
      });

      const context = makeContext({
        txStyles: { bodyStyle: { levels } },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(32);
      expect(shape.textBody!.paragraphs[1].runs[0].properties.fontSize).toBe(28);
    });

    it("指定レベルがない場合は defaultParagraph にフォールバック", () => {
      const shape = makeShape({ placeholderType: "body" });
      const context = makeContext({
        txStyles: {
          bodyStyle: {
            defaultParagraph: { defaultRunProperties: { fontSize: 20 } },
            levels: Array(9).fill(undefined),
          },
        },
      });

      applyTextStyleInheritance([shape], context);

      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(20);
    });
  });

  describe("グループ内のシェイプ", () => {
    it("グループ内のシェイプにも再帰的に適用される", () => {
      const innerShape = makeShape({ placeholderType: "title" });
      const group: GroupElement = {
        type: "group",
        transform: {
          offsetX: 0,
          offsetY: 0,
          extentWidth: 9144000,
          extentHeight: 5143500,
          rotation: 0,
          flipH: false,
          flipV: false,
        },
        childTransform: {
          offsetX: 0,
          offsetY: 0,
          extentWidth: 9144000,
          extentHeight: 5143500,
          rotation: 0,
          flipH: false,
          flipV: false,
        },
        children: [innerShape],
        effects: null,
      };

      const context = makeContext({
        txStyles: { titleStyle: makeDefaultTextStyle({ fontSize: 44 }) },
      });

      applyTextStyleInheritance([group], context);

      expect(innerShape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(44);
    });
  });

  describe("テキストボディなし", () => {
    it("textBody が null のシェイプはスキップされる", () => {
      const shape = makeShape({ textBody: null } as Partial<ShapeElement>);

      expect(() => {
        applyTextStyleInheritance([shape], makeContext());
      }).not.toThrow();
    });
  });

  describe("非プレースホルダーシェイプの layout/master マッチング", () => {
    it("非プレースホルダーシェイプは layout/master の lstStyle を参照しない", () => {
      const shape = makeShape(); // placeholderType なし
      const context = makeContext({
        layoutPlaceholderStyles: [
          { placeholderType: "body", lstStyle: makeDefaultTextStyle({ fontSize: 20 }) },
        ],
        masterPlaceholderStyles: [
          { placeholderType: "body", lstStyle: makeDefaultTextStyle({ fontSize: 24 }) },
        ],
        defaultTextStyle: makeDefaultTextStyle({ fontSize: 14 }),
      });

      applyTextStyleInheritance([shape], context);

      // layout/master の lstStyle は使われず、defaultTextStyle から取得
      expect(shape.textBody!.paragraphs[0].runs[0].properties.fontSize).toBe(14);
    });
  });
});
