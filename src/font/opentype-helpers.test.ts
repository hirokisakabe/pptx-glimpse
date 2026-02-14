import { describe, it, expect } from "vitest";
import { createOpentypeTextMeasurerFromBuffers } from "./opentype-helpers.js";

/**
 * opentype.js を使って最小限の有効な TTF バッファを作成する。
 */
async function createTestFontBuffer(familyName: string): Promise<ArrayBuffer> {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
  const opentype: {
    Glyph: new (opts: Record<string, unknown>) => unknown;
    Path: new () => {
      moveTo(x: number, y: number): void;
      lineTo(x: number, y: number): void;
      close(): void;
    };
    Font: new (opts: Record<string, unknown>) => { toArrayBuffer(): ArrayBuffer };
  } = await import("opentype.js");

  const notdefGlyph = new opentype.Glyph({
    name: ".notdef",
    unicode: 0,
    advanceWidth: 650,
    path: new opentype.Path(),
  });

  const pathA = new opentype.Path();
  pathA.moveTo(0, 0);
  pathA.lineTo(300, 800);
  pathA.lineTo(600, 0);
  pathA.close();

  const glyphA = new opentype.Glyph({
    name: "A",
    unicode: 65,
    advanceWidth: 600,
    path: pathA,
  });

  const spaceGlyph = new opentype.Glyph({
    name: "space",
    unicode: 32,
    advanceWidth: 250,
    path: new opentype.Path(),
  });

  const font = new opentype.Font({
    familyName,
    styleName: "Regular",
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    glyphs: [notdefGlyph, spaceGlyph, glyphA],
  });

  return font.toArrayBuffer();
}

describe("createOpentypeTextMeasurerFromBuffers", () => {
  it("空配列で null を返す", async () => {
    const result = await createOpentypeTextMeasurerFromBuffers([]);
    expect(result).toBeNull();
  });

  it("不正バッファで null を返す", async () => {
    const result = await createOpentypeTextMeasurerFromBuffers([
      { name: "Invalid", data: new Uint8Array([0, 0, 0, 0]) },
    ]);
    expect(result).toBeNull();
  });

  it("フォントバッファから OpentypeTextMeasurer を構築する", async () => {
    const buffer = await createTestFontBuffer("TestFont");
    const measurer = await createOpentypeTextMeasurerFromBuffers([
      { name: "TestFont", data: buffer },
    ]);
    expect(measurer).not.toBeNull();

    // "A" の幅を計測（advanceWidth=600, unitsPerEm=1000）
    const width = measurer!.measureTextWidth("A", 18, false, "TestFont");
    const expected = (600 / 1000) * 18 * (96 / 72);
    expect(width).toBeCloseTo(expected, 1);
  });

  it("行高さを正しく計算する", async () => {
    const buffer = await createTestFontBuffer("TestFont");
    const measurer = await createOpentypeTextMeasurerFromBuffers([
      { name: "TestFont", data: buffer },
    ]);
    expect(measurer).not.toBeNull();

    // ascender=800, descender=-200 → (800+200)/1000 = 1.0
    const ratio = measurer!.getLineHeightRatio("TestFont");
    expect(ratio).toBeCloseTo(1.0, 5);
  });

  it("フォントマッピング逆引きで PPTX 名でも解決できる", async () => {
    const buffer = await createTestFontBuffer("Carlito");
    const measurer = await createOpentypeTextMeasurerFromBuffers([
      { name: "Carlito", data: buffer },
    ]);
    expect(measurer).not.toBeNull();

    // デフォルトマッピング: Calibri → Carlito
    // 逆引きで "Calibri" でも Carlito のフォントが使える
    const widthByOss = measurer!.measureTextWidth("A", 18, false, "Carlito");
    const widthByPptx = measurer!.measureTextWidth("A", 18, false, "Calibri");
    expect(widthByPptx).toBe(widthByOss);
  });

  it("カスタムフォントマッピングで逆引きが機能する", async () => {
    const buffer = await createTestFontBuffer("MyFont");
    const measurer = await createOpentypeTextMeasurerFromBuffers(
      [{ name: "MyFont", data: buffer }],
      { "Custom Name": "MyFont" },
    );
    expect(measurer).not.toBeNull();

    const widthByOss = measurer!.measureTextWidth("A", 18, false, "MyFont");
    const widthByCustom = measurer!.measureTextWidth("A", 18, false, "Custom Name");
    expect(widthByCustom).toBe(widthByOss);
  });

  it("name なしのバッファは defaultFont として使われる", async () => {
    const buffer = await createTestFontBuffer("Unnamed");
    const measurer = await createOpentypeTextMeasurerFromBuffers([{ data: buffer }]);
    expect(measurer).not.toBeNull();

    // 未知のフォント名でも defaultFont が使われるため opentype ベースの計算になる
    const width = measurer!.measureTextWidth("A", 18, false, "UnknownFont");
    const expected = (600 / 1000) * 18 * (96 / 72);
    expect(width).toBeCloseTo(expected, 1);
  });

  it("Uint8Array バッファでも動作する", async () => {
    const arrayBuffer = await createTestFontBuffer("TestFont");
    const uint8 = new Uint8Array(arrayBuffer);
    const measurer = await createOpentypeTextMeasurerFromBuffers([
      { name: "TestFont", data: uint8 },
    ]);
    expect(measurer).not.toBeNull();

    const width = measurer!.measureTextWidth("A", 18, false, "TestFont");
    expect(width).toBeGreaterThan(0);
  });

  it("複数フォントを登録できる", async () => {
    const buffer1 = await createTestFontBuffer("Font1");
    const buffer2 = await createTestFontBuffer("Font2");
    const measurer = await createOpentypeTextMeasurerFromBuffers([
      { name: "Font1", data: buffer1 },
      { name: "Font2", data: buffer2 },
    ]);
    expect(measurer).not.toBeNull();

    const width1 = measurer!.measureTextWidth("A", 18, false, "Font1");
    const width2 = measurer!.measureTextWidth("A", 18, false, "Font2");
    expect(width1).toBeGreaterThan(0);
    expect(width2).toBeGreaterThan(0);
  });
});
