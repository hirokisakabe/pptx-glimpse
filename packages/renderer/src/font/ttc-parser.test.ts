import { describe, expect, it } from "vitest";

import { extractTtcFonts, isTtcBuffer } from "./ttc-parser.js";
import { buildTtcFromTtfs } from "./ttc-test-helper.js";

/**
 * Test note.
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
  it("covers ttc-parser behavior 1", () => {
    const buf = new ArrayBuffer(12);
    const view = new DataView(buf);
    view.setUint32(0, 0x74746366); // "ttcf"
    expect(isTtcBuffer(buf)).toBe(true);
  });

  it("covers ttc-parser behavior 2", async () => {
    const ttf = await createTestTtfBuffer("TestFont");
    expect(isTtcBuffer(ttf)).toBe(false);
  });

  it("covers ttc-parser behavior 3", () => {
    expect(isTtcBuffer(new ArrayBuffer(0))).toBe(false);
  });

  it("covers ttc-parser behavior 4", () => {
    expect(isTtcBuffer(new ArrayBuffer(3))).toBe(false);
  });

  it("covers ttc-parser behavior 5", () => {
    const buf = new ArrayBuffer(12);
    const view = new DataView(buf);
    view.setUint32(0, 0x74746366);
    expect(isTtcBuffer(new Uint8Array(buf))).toBe(true);
  });
});

describe("extractTtcFonts", () => {
  it("covers ttc-parser behavior 6", async () => {
    const ttf1 = await createTestTtfBuffer("FontAlpha");
    const ttf2 = await createTestTtfBuffer("FontBeta");
    const ttc = buildTtcFromTtfs([ttf1, ttf2]);

    const extracted = extractTtcFonts(ttc);
    expect(extracted).toHaveLength(2);
  });

  it("covers ttc-parser behavior 7", async () => {
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

  it("covers ttc-parser behavior 8", async () => {
    const ttf = await createTestTtfBuffer("TestFont");
    expect(extractTtcFonts(ttf)).toEqual([]);
  });

  it("covers ttc-parser behavior 9", () => {
    expect(extractTtcFonts(new ArrayBuffer(0))).toEqual([]);
  });

  it("covers ttc-parser behavior 10", () => {
    const buf = new ArrayBuffer(12);
    const view = new DataView(buf);
    view.setUint32(0, 0x74746366);
    view.setUint16(4, 1);
    view.setUint16(6, 0);
    view.setUint32(8, 0);
    expect(extractTtcFonts(buf)).toEqual([]);
  });

  it("covers ttc-parser behavior 11", async () => {
    const ttf1 = await createTestTtfBuffer("FontAlpha");
    const ttc = buildTtcFromTtfs([ttf1]);
    const uint8 = new Uint8Array(ttc);

    const extracted = extractTtcFonts(uint8);
    expect(extracted).toHaveLength(1);
  });

  it("covers ttc-parser behavior 12", async () => {
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

  it("covers ttc-parser behavior 13", async () => {
    const ttf1 = await createTestTtfBuffer("GoodFont");
    const ttc = buildTtcFromTtfs([ttf1]);

    // Test note.
    const view = new DataView(ttc);
    const fontOffset = view.getUint32(12); // Test note.
    const firstTableRecordOffset = fontOffset + 12; // Test note.
    // Test note.
    view.setUint32(firstTableRecordOffset + 8, 0xffffffff);

    const extracted = extractTtcFonts(ttc);
    // Test note.
    expect(extracted).toEqual([]);
  });

  it("covers ttc-parser behavior 14", async () => {
    const ttf1 = await createTestTtfBuffer("GoodFont");
    const ttf2 = await createTestTtfBuffer("BadFont");
    const ttc = buildTtcFromTtfs([ttf1, ttf2]);

    // Test note.
    const view = new DataView(ttc);
    const font2Offset = view.getUint32(16); // Test note.
    const firstTableRecordOffset = font2Offset + 12;
    view.setUint32(firstTableRecordOffset + 8, 0xffffffff);

    const extracted = extractTtcFonts(ttc);
    // Test note.
    expect(extracted).toHaveLength(1);
  });
});
