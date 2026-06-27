import { describe, expect, it } from "vitest";

import { subsetFont } from "./font-subsetter.js";
import type { OpentypeFullFont } from "./text-path-context.js";

interface ParsedFontForTest {
  charToGlyph(char: string): { index: number };
  glyphs: { length: number };
}

interface OpentypeTestModule {
  Glyph: new (opts: Record<string, unknown>) => unknown;
  Path: new () => {
    moveTo(x: number, y: number): void;
    lineTo(x: number, y: number): void;
    close(): void;
  };
  Font: new (opts: Record<string, unknown>) => { toArrayBuffer(): ArrayBuffer };
  parse: (buffer: ArrayBuffer) => ParsedFontForTest;
}

async function loadOpentype(): Promise<OpentypeTestModule> {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
  const mod: OpentypeTestModule = await import("opentype.js");
  return mod;
}

/**
 * opentype.js を使ってテスト用の TTF フォントを作成しパースして返す。
 * グリフ: .notdef, space, A, B (A と B は同形の三角形)
 */
async function createParsedTestFont(): Promise<OpentypeFullFont> {
  const opentype = await loadOpentype();

  const notdefGlyph = new opentype.Glyph({
    name: ".notdef",
    advanceWidth: 650,
    path: new opentype.Path(),
  });

  const makeTrianglePath = () => {
    const path = new opentype.Path();
    path.moveTo(0, 0);
    path.lineTo(300, 800);
    path.lineTo(600, 0);
    path.close();
    return path;
  };

  const glyphA = new opentype.Glyph({
    name: "A",
    unicode: 65,
    advanceWidth: 600,
    path: makeTrianglePath(),
  });

  const glyphB = new opentype.Glyph({
    name: "B",
    unicode: 66,
    advanceWidth: 700,
    path: makeTrianglePath(),
  });

  const spaceGlyph = new opentype.Glyph({
    name: "space",
    unicode: 32,
    advanceWidth: 250,
    path: new opentype.Path(),
  });

  const font = new opentype.Font({
    familyName: "SubsetTestFont",
    styleName: "Regular",
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    glyphs: [notdefGlyph, spaceGlyph, glyphA, glyphB],
  });

  return opentype.parse(font.toArrayBuffer()) as unknown as OpentypeFullFont;
}

async function parseSubsetBuffer(buffer: Uint8Array): Promise<ParsedFontForTest> {
  const opentype = await loadOpentype();
  return opentype.parse(
    buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer,
  );
}

describe("subsetFont", () => {
  it("使用文字のみを含むパース可能な OTF を生成する", async () => {
    const font = await createParsedTestFont();
    const buffer = await subsetFont(font, new Set(["A", " "]), "SubsetTestFont");

    expect(buffer).not.toBeNull();

    const parsed = await parseSubsetBuffer(buffer!);

    // A と space は収録され、B は収録されない (.notdef + space + A = 3 グリフ)
    expect(parsed.charToGlyph("A").index).toBeGreaterThan(0);
    expect(parsed.charToGlyph(" ").index).toBeGreaterThan(0);
    expect(parsed.charToGlyph("B").index).toBe(0);
    expect(parsed.glyphs.length).toBe(3);
  });

  it("フォント未収録の文字はサブセットに含めない", async () => {
    const font = await createParsedTestFont();
    const buffer = await subsetFont(font, new Set(["A", "Z"]), "SubsetTestFont");

    expect(buffer).not.toBeNull();

    const parsed = await parseSubsetBuffer(buffer!);

    // Z は元フォントに無いので .notdef にもならず除外される
    expect(parsed.charToGlyph("Z").index).toBe(0);
    expect(parsed.glyphs.length).toBe(2); // .notdef + A
  });

  it("収録文字が 1 つも無い場合は null を返す", async () => {
    const font = await createParsedTestFont();
    const buffer = await subsetFont(font, new Set(["Z", "あ"]), "SubsetTestFont");

    expect(buffer).toBeNull();
  });

  it("空の文字集合で null を返す", async () => {
    const font = await createParsedTestFont();
    const buffer = await subsetFont(font, new Set(), "SubsetTestFont");

    expect(buffer).toBeNull();
  });

  it("charToGlyph を持たないオブジェクトで null を返す", async () => {
    const fakeFont = { unitsPerEm: 1000 } as unknown as OpentypeFullFont;
    const buffer = await subsetFont(fakeFont, new Set(["A"]), "Fake");

    expect(buffer).toBeNull();
  });
});
