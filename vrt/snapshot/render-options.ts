import { existsSync } from "node:fs";
import { mkdir, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import type { ConvertOptions } from "../../packages/core/src/converter.js";

type VrtRenderOptions = Pick<ConvertOptions, "fontDirs" | "skipSystemFonts">;

// Document-path parity VRT compares the parser and document renderers at zero
// tolerance. Keep it fontless so the gate remains focused on adapter parity
// while still avoiding local system-font scans; standard snapshot VRT uses the
// generated fonts below for text-path visual coverage.
export const DOCUMENT_PATH_VRT_RENDER_OPTIONS = {
  skipSystemFonts: true,
} as const satisfies Pick<ConvertOptions, "skipSystemFonts">;

const VRT_FONT_DIR = join(tmpdir(), "pptx-glimpse-vrt-fonts-v2");
const VRT_FONT_FAMILIES = [
  "Carlito",
  "Arimo",
  "Tinos",
  "Cousine",
  "Caladea",
  "Noto Sans JP",
  "Noto Serif CJK JP",
] as const;
const VRT_CJK_SAMPLE_TEXT =
  "、。いえかがきくこさしじすそたっつてでとなにのはべまみめもよらりるれを" +
  "アイウエオカキクグケコサシステトドネフプベホマメャュョラリンー" +
  "一下中丸事互交位作使値側像優先入全内円出前力加動半単印参収右合同向図在型基場大央字完定小常座弧形必成承折拡持指挙挟接揃描数文方日普曲書最未本楕標横次正残段混済準漢照生用画白的直矢矩確示空章素紫継緑線縦縮置義背色落行表装複要角設認語赤返通配長間青非面飾黄（）";

let renderOptionsPromise: Promise<VrtRenderOptions> | null = null;

// Local snapshot VRT must not depend on parsing every OS system font. Generate a
// small deterministic font set and use it for both Docker snapshots and local VRT.
export function getVrtRenderOptions(): Promise<VrtRenderOptions> {
  renderOptionsPromise ??= ensureVrtFontDir().then((fontDir) => ({
    fontDirs: [fontDir],
    skipSystemFonts: true,
  }));
  return renderOptionsPromise;
}

async function ensureVrtFontDir(): Promise<string> {
  await mkdir(VRT_FONT_DIR, { recursive: true });
  await Promise.all(
    VRT_FONT_FAMILIES.map(async (familyName) => {
      const fontPath = join(VRT_FONT_DIR, `${familyName.replaceAll(" ", "")}.ttf`);
      if (existsSync(fontPath)) return;
      const buffer = await createVrtFontBuffer(familyName);
      await writeFile(fontPath, Buffer.from(buffer));
    }),
  );
  return VRT_FONT_DIR;
}

async function createVrtFontBuffer(familyName: string): Promise<ArrayBuffer> {
  const opentype: {
    Glyph: new (opts: Record<string, unknown>) => unknown;
    Path: new () => {
      moveTo(x: number, y: number): void;
      lineTo(x: number, y: number): void;
      close(): void;
    };
    Font: new (opts: Record<string, unknown>) => { toArrayBuffer(): ArrayBuffer };
  } = await import("opentype.js");

  const boxPath = new opentype.Path();
  boxPath.moveTo(80, 0);
  boxPath.lineTo(80, 700);
  boxPath.lineTo(540, 700);
  boxPath.lineTo(540, 0);
  boxPath.close();

  const notdefGlyph = new opentype.Glyph({
    name: ".notdef",
    unicode: 0,
    advanceWidth: 620,
    path: boxPath,
  });

  const spaceGlyph = new opentype.Glyph({
    name: "space",
    unicode: 32,
    advanceWidth: 260,
    path: new opentype.Path(),
  });

  const glyphs = [notdefGlyph, spaceGlyph];
  for (let codePoint = 33; codePoint <= 126; codePoint++) {
    glyphs.push(
      new opentype.Glyph({
        name: `uni${codePoint.toString(16).toUpperCase().padStart(4, "0")}`,
        unicode: codePoint,
        advanceWidth: 620,
        path: createLatinGlyphPath(),
      }),
    );
  }

  for (const char of VRT_CJK_SAMPLE_TEXT) {
    const codePoint = char.codePointAt(0);
    if (codePoint === undefined) continue;
    glyphs.push(
      new opentype.Glyph({
        name: `uni${codePoint.toString(16).toUpperCase().padStart(4, "0")}`,
        unicode: codePoint,
        advanceWidth: 1000,
        path: createCjkGlyphPath(),
      }),
    );
  }

  const font = new opentype.Font({
    familyName,
    styleName: "Regular",
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    glyphs,
  });

  return font.toArrayBuffer();

  function createLatinGlyphPath(): InstanceType<typeof opentype.Path> {
    const path = new opentype.Path();
    path.moveTo(70, 0);
    path.lineTo(310, 700);
    path.lineTo(550, 0);
    path.close();
    return path;
  }

  function createCjkGlyphPath(): InstanceType<typeof opentype.Path> {
    const path = new opentype.Path();
    path.moveTo(90, 0);
    path.lineTo(90, 760);
    path.lineTo(910, 760);
    path.lineTo(910, 0);
    path.close();
    return path;
  }
}
