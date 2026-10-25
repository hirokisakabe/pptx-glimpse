import { afterEach, describe, expect, it, vi } from "vitest";

import { createOpentypeSetupFromBuffers, tryLoadOpentype } from "./opentype-buffer-helpers.js";

interface OpentypeTestModule {
  parse: (buffer: ArrayBuffer, options?: { lowMemory?: boolean }) => EagerFont;
  Glyph: new (opts: Record<string, unknown>) => unknown;
  Path: new () => {
    moveTo(x: number, y: number): void;
    lineTo(x: number, y: number): void;
    close(): void;
  };
  Font: new (opts: Record<string, unknown>) => { toArrayBuffer(): ArrayBuffer };
}

interface EagerFont {
  getAdvanceWidth(text: string, fontSize: number): number;
  getPath(text: string, x: number, y: number, fontSize: number): { toPathData(): string };
}

const parseCalls = vi.hoisted(() => {
  const calls: unknown[] = [];
  return calls;
});

vi.mock("opentype.js", async (importOriginal) => {
  const actual = await importOriginal<OpentypeTestModule>();
  return {
    ...actual,
    parse: (buffer: ArrayBuffer, options?: { lowMemory?: boolean }) => {
      parseCalls.push(options);
      return actual.parse(buffer, options);
    },
  };
});

/**
 * Build a CFF-flavored OpenType font with distinct glyph shapes and advance widths so that
 * lazy and eager parsing can be compared glyph by glyph.
 */
async function createMultiGlyphFont(familyName: string): Promise<ArrayBuffer> {
  const opentype = await vi.importActual<OpentypeTestModule>("opentype.js");
  const triangle = (width: number, height: number) => {
    const path = new opentype.Path();
    path.moveTo(0, 0);
    path.lineTo(width / 2, height);
    path.lineTo(width, 0);
    path.close();
    return path;
  };
  const glyphs = [
    new opentype.Glyph({
      name: ".notdef",
      unicode: 0,
      advanceWidth: 500,
      path: new opentype.Path(),
    }),
    new opentype.Glyph({
      name: "space",
      unicode: 32,
      advanceWidth: 250,
      path: new opentype.Path(),
    }),
    new opentype.Glyph({ name: "A", unicode: 65, advanceWidth: 600, path: triangle(600, 700) }),
    new opentype.Glyph({ name: "B", unicode: 66, advanceWidth: 640, path: triangle(640, 720) }),
    new opentype.Glyph({
      name: "uni4E2D",
      unicode: 0x4e2d,
      advanceWidth: 1000,
      path: triangle(1000, 880),
    }),
  ];
  const font = new opentype.Font({
    familyName,
    styleName: "Regular",
    unitsPerEm: 1000,
    ascender: 880,
    descender: -120,
    glyphs,
  });
  return font.toArrayBuffer();
}

afterEach(() => {
  parseCalls.length = 0;
});

describe("tryLoadOpentype", () => {
  it("parses font buffers in opentype.js low-memory mode", async () => {
    const opentype = await tryLoadOpentype();
    expect(opentype).not.toBeNull();

    opentype?.parse(await createMultiGlyphFont("LowMemoryFont"));

    expect(parseCalls).toEqual([{ lowMemory: true }]);
  });
});

describe("createOpentypeSetupFromBuffers (low-memory parsing)", () => {
  it("measures and outlines text exactly like an eagerly parsed font", async () => {
    const buffer = await createMultiGlyphFont("ParityFont");
    const actual = await vi.importActual<OpentypeTestModule>("opentype.js");
    const eager = actual.parse(buffer.slice(0));

    const setup = await createOpentypeSetupFromBuffers([{ data: buffer }]);
    expect(setup).not.toBeNull();
    expect(parseCalls).toEqual([{ lowMemory: true }]);

    const text = "AB A中B";
    const fontSizePt = 18;
    const fontSizePx = (fontSizePt * 96) / 72;
    expect(setup?.measurer.measureTextWidth(text, fontSizePt, false, "ParityFont")).toBeCloseTo(
      eager.getAdvanceWidth(text, fontSizePx),
      6,
    );

    const resolved = setup?.fontResolver.resolveFont("ParityFont", null);
    expect(resolved).toBeDefined();
    expect(resolved?.getPath(text, 0, 0, fontSizePx).toPathData()).toBe(
      eager.getPath(text, 0, 0, fontSizePx).toPathData(),
    );
  });
});
