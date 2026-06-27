import { existsSync } from "node:fs";
import { mkdir, readdir, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";

import type { ConvertOptions } from "../../packages/core/src/converter.js";

type VrtRenderOptions = Pick<ConvertOptions, "fontDirs" | "skipSystemFonts">;

const __dirname = dirname(fileURLToPath(import.meta.url));

// Document-path parity VRT compares the parser and document renderers at zero
// tolerance. Keep it fontless so the gate remains focused on adapter parity
// while still avoiding local system-font scans; standard snapshot VRT uses the
// generated fonts below for text-path visual coverage.
export const DOCUMENT_PATH_VRT_RENDER_OPTIONS = {
  skipSystemFonts: true,
} as const satisfies Pick<ConvertOptions, "skipSystemFonts">;

const VRT_FONT_DIR = join(tmpdir(), "pptx-glimpse-vrt-fonts-v4");
const VRT_FONT_MANIFEST = join(VRT_FONT_DIR, "manifest.txt");
const VRT_FONT_FAMILIES = [
  "Carlito",
  "Arimo",
  "Tinos",
  "Cousine",
  "Caladea",
  "Noto Sans JP",
  "Noto Serif CJK JP",
] as const;
const GENERATED_FIXTURE_DIR = join(__dirname, "fixtures");
const SHARED_FIXTURE_DIR = join(__dirname, "..", "..", "shared-fixtures");

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
  const codePoints = await collectVrtTextCodePoints();
  const manifest = [...codePoints].sort((a, b) => a - b).join(",");
  const fontPaths = VRT_FONT_FAMILIES.map((familyName) =>
    join(VRT_FONT_DIR, `${familyName.replaceAll(" ", "")}.ttf`),
  );
  if (
    fontPaths.every((fontPath) => existsSync(fontPath)) &&
    (await readTextIfExists(VRT_FONT_MANIFEST)) === manifest
  ) {
    return VRT_FONT_DIR;
  }

  await Promise.all(
    VRT_FONT_FAMILIES.map(async (familyName, index) => {
      const fontPath = fontPaths[index];
      const buffer = await createVrtFontBuffer(familyName, codePoints);
      await writeFile(fontPath, Buffer.from(buffer));
    }),
  );
  await writeFile(VRT_FONT_MANIFEST, manifest);
  return VRT_FONT_DIR;
}

async function readTextIfExists(path: string): Promise<string | null> {
  try {
    return await readFile(path, "utf8");
  } catch {
    return null;
  }
}

async function collectVrtTextCodePoints(): Promise<Set<number>> {
  const codePoints = new Set<number>();
  codePoints.add(0);
  codePoints.add(0x0a);
  for (let codePoint = 32; codePoint <= 126; codePoint++) {
    codePoints.add(codePoint);
  }

  for (const fixtureDir of [GENERATED_FIXTURE_DIR, SHARED_FIXTURE_DIR]) {
    if (!existsSync(fixtureDir)) continue;
    const entries = await readdir(fixtureDir);
    await Promise.all(
      entries
        .filter((entry) => entry.endsWith(".pptx"))
        .map(async (entry) => {
          const input = await readFile(join(fixtureDir, entry));
          const zip = await JSZip.loadAsync(input);
          await Promise.all(
            Object.entries(zip.files)
              .filter(([name]) => name.endsWith(".xml"))
              .map(async ([, file]) => {
                const xml = await file.async("string");
                for (const match of xml.matchAll(/<a:t[^>]*>(.*?)<\/a:t>/gs)) {
                  for (const char of match[1]) {
                    const codePoint = char.codePointAt(0);
                    if (codePoint !== undefined) codePoints.add(codePoint);
                  }
                }
              }),
          );
        }),
    );
  }

  return codePoints;
}

async function createVrtFontBuffer(
  familyName: string,
  codePoints: ReadonlySet<number>,
): Promise<ArrayBuffer> {
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
  for (const codePoint of [...codePoints].sort((a, b) => a - b)) {
    if (codePoint === 0 || codePoint === 32) continue;
    const isControl = codePoint < 32;
    const isWide = codePoint >= 0x3000 || codePoint > 0xffff;
    glyphs.push(
      new opentype.Glyph({
        name: `uni${codePoint.toString(16).toUpperCase().padStart(4, "0")}`,
        unicode: codePoint,
        advanceWidth: isControl ? 0 : isWide ? 1000 : 620,
        path: isControl
          ? new opentype.Path()
          : isWide
            ? createCjkGlyphPath()
            : createLatinGlyphPath(),
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
