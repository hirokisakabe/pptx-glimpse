import { describe, expect, it } from "vitest";

import { extractTtcFonts, isTtcBuffer } from "./ttc-parser.js";
import { buildTtcFromTtfs } from "./ttc-test-helper.js";

/**
 * Create a minimal valid TTF buffer with opentype.js.
 */
async function createTestTtfBuffer(familyName: string): Promise<ArrayBuffer> {
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
  it("Correctly determine TTC buffer", () => {
    const buf = new ArrayBuffer(12);
    const view = new DataView(buf);
    view.setUint32(0, 0x74746366); // "ttcf"
    expect(isTtcBuffer(buf)).toBe(true);
  });

  it("TTF buffer returns false", async () => {
    const ttf = await createTestTtfBuffer("TestFont");
    expect(isTtcBuffer(ttf)).toBe(false);
  });

  it("Empty buffer returns false", () => {
    expect(isTtcBuffer(new ArrayBuffer(0))).toBe(false);
  });

  it("Buffers less than 4 bytes return false", () => {
    expect(isTtcBuffer(new ArrayBuffer(3))).toBe(false);
  });

  it("Works with Uint8Array input", () => {
    const buf = new ArrayBuffer(12);
    const view = new DataView(buf);
    view.setUint32(0, 0x74746366);
    expect(isTtcBuffer(new Uint8Array(buf))).toBe(true);
  });
});

describe("extractTtcFonts", () => {
  it("Individual TTF can be extracted from TTC", async () => {
    const ttf1 = await createTestTtfBuffer("FontAlpha");
    const ttf2 = await createTestTtfBuffer("FontBeta");
    const ttc = buildTtcFromTtfs([ttf1, ttf2]);

    const extracted = extractTtcFonts(ttc);
    expect(extracted).toHaveLength(2);
  });

  it("The extracted TTF can be parsed with opentype.js", async () => {
    const ttf1 = await createTestTtfBuffer("FontAlpha");
    const ttf2 = await createTestTtfBuffer("FontBeta");
    const ttc = buildTtcFromTtfs([ttf1, ttf2]);

    const extracted = extractTtcFonts(ttc);
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

  it("Non-TTC buffers return an empty array", async () => {
    const ttf = await createTestTtfBuffer("TestFont");
    expect(extractTtcFonts(ttf)).toEqual([]);
  });

  it("Empty buffer returns empty array", () => {
    expect(extractTtcFonts(new ArrayBuffer(0))).toEqual([]);
  });

  it("TTC with numFonts 0 returns empty array", () => {
    const buf = new ArrayBuffer(12);
    const view = new DataView(buf);
    view.setUint32(0, 0x74746366);
    view.setUint16(4, 1);
    view.setUint16(6, 0);
    view.setUint32(8, 0);
    expect(extractTtcFonts(buf)).toEqual([]);
  });

  it("Works with Uint8Array input", async () => {
    const ttf1 = await createTestTtfBuffer("FontAlpha");
    const ttc = buildTtcFromTtfs([ttf1]);
    const uint8 = new Uint8Array(ttc);

    const extracted = extractTtcFonts(uint8);
    expect(extracted).toHaveLength(1);
  });

  it("Works even with single font TTC", async () => {
    const ttf = await createTestTtfBuffer("SingleFont");
    const ttc = buildTtcFromTtfs([ttf]);

    const extracted = extractTtcFonts(ttc);
    expect(extracted).toHaveLength(1);

    const opentype: {
      parse: (buf: ArrayBuffer) => { names: { fontFamily: Record<string, string> } };
    } = await import("opentype.js");
    const font = opentype.parse(extracted[0]);
    expect(Object.values(font.names.fontFamily)).toContain("SingleFont");
  });

  it("Fonts with table offsets out of range are skipped", async () => {
    const ttf1 = await createTestTtfBuffer("GoodFont");
    const ttc = buildTtcFromTtfs([ttf1]);

    // Rewrite offset of table record of first font in TTC to be out of range
    const view = new DataView(ttc);
    const fontOffset = view.getUint32(12); // first font offset
    const firstTableRecordOffset = fontOffset + 12; // first table record
    // Rewrite table offset to huge value
    view.setUint32(firstTableRecordOffset + 8, 0xffffffff);

    const extracted = extractTtcFonts(ttc);
    // Empty array because invalid fonts will be skipped
    expect(extracted).toEqual([]);
  });

  it("Even if one out of two fonts is invalid, the normal font will be extracted.", async () => {
    const ttf1 = await createTestTtfBuffer("GoodFont");
    const ttf2 = await createTestTtfBuffer("BadFont");
    const ttc = buildTtcFromTtfs([ttf1, ttf2]);

    // Rewrite table offset of second font to be out of range
    const view = new DataView(ttc);
    const font2Offset = view.getUint32(16); // second font offset
    const firstTableRecordOffset = font2Offset + 12;
    view.setUint32(firstTableRecordOffset + 8, 0xffffffff);

    const extracted = extractTtcFonts(ttc);
    // Extract only the first normal font
    expect(extracted).toHaveLength(1);
  });
});
