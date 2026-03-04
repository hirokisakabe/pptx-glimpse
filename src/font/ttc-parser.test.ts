import { describe, it, expect } from "vitest";
import { isTtcBuffer, extractTtcFonts } from "./ttc-parser.js";
import { buildTtcFromTtfs } from "./ttc-test-helper.js";

/**
 * opentype.js で最小限の有効な TTF バッファを作成する。
 */
async function createTestTtfBuffer(familyName: string): Promise<ArrayBuffer> {
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

describe("isTtcBuffer", () => {
  it("TTC バッファを正しく判定する", () => {
    const buf = new ArrayBuffer(12);
    const view = new DataView(buf);
    view.setUint32(0, 0x74746366); // "ttcf"
    expect(isTtcBuffer(buf)).toBe(true);
  });

  it("TTF バッファは false を返す", async () => {
    const ttf = await createTestTtfBuffer("TestFont");
    expect(isTtcBuffer(ttf)).toBe(false);
  });

  it("空バッファは false を返す", () => {
    expect(isTtcBuffer(new ArrayBuffer(0))).toBe(false);
  });

  it("4 バイト未満のバッファは false を返す", () => {
    expect(isTtcBuffer(new ArrayBuffer(3))).toBe(false);
  });

  it("Uint8Array 入力で動作する", () => {
    const buf = new ArrayBuffer(12);
    const view = new DataView(buf);
    view.setUint32(0, 0x74746366);
    expect(isTtcBuffer(new Uint8Array(buf))).toBe(true);
  });
});

describe("extractTtcFonts", () => {
  it("TTC から個別 TTF を抽出できる", async () => {
    const ttf1 = await createTestTtfBuffer("FontAlpha");
    const ttf2 = await createTestTtfBuffer("FontBeta");
    const ttc = buildTtcFromTtfs([ttf1, ttf2]);

    const extracted = extractTtcFonts(ttc);
    expect(extracted).toHaveLength(2);
  });

  it("抽出した TTF が opentype.js でパースできる", async () => {
    const ttf1 = await createTestTtfBuffer("FontAlpha");
    const ttf2 = await createTestTtfBuffer("FontBeta");
    const ttc = buildTtcFromTtfs([ttf1, ttf2]);

    const extracted = extractTtcFonts(ttc);
    // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
    const opentype: {
      parse: (buf: ArrayBuffer) => { names: { fontFamily: Record<string, string> } };
    } = await import("opentype.js");

    const font1 = opentype.parse(extracted[0]);
    const font2 = opentype.parse(extracted[1]);

    const names1 = Object.values(font1.names.fontFamily);
    const names2 = Object.values(font2.names.fontFamily);
    expect(names1).toContain("FontAlpha");
    expect(names2).toContain("FontBeta");
  });

  it("TTC でないバッファは空配列を返す", async () => {
    const ttf = await createTestTtfBuffer("TestFont");
    expect(extractTtcFonts(ttf)).toEqual([]);
  });

  it("空バッファは空配列を返す", () => {
    expect(extractTtcFonts(new ArrayBuffer(0))).toEqual([]);
  });

  it("numFonts が 0 の TTC は空配列を返す", () => {
    const buf = new ArrayBuffer(12);
    const view = new DataView(buf);
    view.setUint32(0, 0x74746366);
    view.setUint16(4, 1);
    view.setUint16(6, 0);
    view.setUint32(8, 0);
    expect(extractTtcFonts(buf)).toEqual([]);
  });

  it("Uint8Array 入力で動作する", async () => {
    const ttf1 = await createTestTtfBuffer("FontAlpha");
    const ttc = buildTtcFromTtfs([ttf1]);
    const uint8 = new Uint8Array(ttc);

    const extracted = extractTtcFonts(uint8);
    expect(extracted).toHaveLength(1);
  });

  it("単一フォントの TTC でも動作する", async () => {
    const ttf = await createTestTtfBuffer("SingleFont");
    const ttc = buildTtcFromTtfs([ttf]);

    const extracted = extractTtcFonts(ttc);
    expect(extracted).toHaveLength(1);

    // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
    const opentype: {
      parse: (buf: ArrayBuffer) => { names: { fontFamily: Record<string, string> } };
    } = await import("opentype.js");
    const font = opentype.parse(extracted[0]);
    expect(Object.values(font.names.fontFamily)).toContain("SingleFont");
  });

  it("テーブルオフセットが範囲外のフォントはスキップされる", async () => {
    const ttf1 = await createTestTtfBuffer("GoodFont");
    const ttc = buildTtcFromTtfs([ttf1]);

    // TTC 内の最初のフォントのテーブルレコードのオフセットを範囲外に書き換え
    const view = new DataView(ttc);
    const fontOffset = view.getUint32(12); // 最初のフォントのオフセット
    const firstTableRecordOffset = fontOffset + 12; // 最初のテーブルレコード
    // テーブルオフセットを巨大な値に書き換え
    view.setUint32(firstTableRecordOffset + 8, 0xffffffff);

    const extracted = extractTtcFonts(ttc);
    // 不正なフォントはスキップされるので空配列
    expect(extracted).toEqual([]);
  });

  it("2フォント中1フォントが不正でも正常なフォントは抽出される", async () => {
    const ttf1 = await createTestTtfBuffer("GoodFont");
    const ttf2 = await createTestTtfBuffer("BadFont");
    const ttc = buildTtcFromTtfs([ttf1, ttf2]);

    // 2番目のフォントのテーブルオフセットを範囲外に書き換え
    const view = new DataView(ttc);
    const font2Offset = view.getUint32(16); // 2番目のフォントのオフセット
    const firstTableRecordOffset = font2Offset + 12;
    view.setUint32(firstTableRecordOffset + 8, 0xffffffff);

    const extracted = extractTtcFonts(ttc);
    // 1番目の正常なフォントのみ抽出
    expect(extracted).toHaveLength(1);
  });
});
