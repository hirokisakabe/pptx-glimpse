import { describe, expect, it } from "vitest";

import { unsafeFixtureAssertion } from "../unsafe-type-assertion.js";
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
 * Create, parse, and return a test TTF font using opentype.js.
 * Glyph:.notdef, space, A, B (A and B are isomorphic triangles)
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

  return unsafeFixtureAssertion<OpentypeFullFont>(opentype.parse(font.toArrayBuffer()));
}

async function parseSubsetBuffer(buffer: Uint8Array): Promise<ParsedFontForTest> {
  const opentype = await loadOpentype();
  return opentype.parse(
    unsafeFixtureAssertion<ArrayBuffer>(
      buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength),
    ),
  );
}

describe("subsetFont", () => {
  it("Generate a parsable OTF containing only used characters", async () => {
    const font = await createParsedTestFont();
    const buffer = await subsetFont(font, new Set(["A", " "]), "SubsetTestFont");

    expect(buffer).not.toBeNull();

    const parsed = await parseSubsetBuffer(buffer!);

    // A and space are included, B is not included (.notdef + space + A = 3 glyphs)
    expect(parsed.charToGlyph("A").index).toBeGreaterThan(0);
    expect(parsed.charToGlyph(" ").index).toBeGreaterThan(0);
    expect(parsed.charToGlyph("B").index).toBe(0);
    expect(parsed.glyphs.length).toBe(3);
  });

  it("Characters not included in the font are not included in the subset.", async () => {
    const font = await createParsedTestFont();
    const buffer = await subsetFont(font, new Set(["A", "Z"]), "SubsetTestFont");

    expect(buffer).not.toBeNull();

    const parsed = await parseSubsetBuffer(buffer!);

    // Since Z is not in the original font, it will not become.notdef and will be excluded.
    expect(parsed.charToGlyph("Z").index).toBe(0);
    expect(parsed.glyphs.length).toBe(2); // .notdef + A
  });

  it("Returns null if there are no characters included.", async () => {
    const font = await createParsedTestFont();
    const buffer = await subsetFont(font, new Set(["Z", "あ"]), "SubsetTestFont");

    expect(buffer).toBeNull();
  });

  it("return null on empty charset", async () => {
    const font = await createParsedTestFont();
    const buffer = await subsetFont(font, new Set(), "SubsetTestFont");

    expect(buffer).toBeNull();
  });

  it("Return null on objects that don't have charToGlyph", async () => {
    const fakeFont = unsafeFixtureAssertion<OpentypeFullFont>({ unitsPerEm: 1000 });
    const buffer = await subsetFont(fakeFont, new Set(["A"]), "Fake");

    expect(buffer).toBeNull();
  });
});
