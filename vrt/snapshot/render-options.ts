import { existsSync } from "node:fs";
import { mkdir, readdir, readFile, unlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";

import type { ConvertOptions } from "../../packages/core/src/converter.js";
import { createFontMapping, getMappedFont } from "../../packages/renderer/src/font/font-mapping.js";

type VrtRenderOptions = Pick<ConvertOptions, "fontDirs" | "skipSystemFonts">;

const __dirname = dirname(fileURLToPath(import.meta.url));

// Document-path parity VRT compares the parser and document renderers at zero
// tolerance. Keep it fontless so the gate remains focused on adapter parity
// while still avoiding local system-font scans; standard snapshot VRT uses the
// generated fonts below for text-path visual coverage.
export const DOCUMENT_PATH_VRT_RENDER_OPTIONS = {
  skipSystemFonts: true,
} as const satisfies Pick<ConvertOptions, "skipSystemFonts">;

const VRT_FONT_GENERATOR_VERSION = 6;
const VRT_FONT_DIR = join(tmpdir(), `pptx-glimpse-vrt-fonts-v${VRT_FONT_GENERATOR_VERSION}`);
const VRT_FONT_MANIFEST = join(VRT_FONT_DIR, "manifest.txt");
const BASE_VRT_FONT_FAMILIES = [
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
const VRT_FONT_MAPPING = createFontMapping();

interface VrtFontInventory {
  codePoints: Set<number>;
  families: string[];
}

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
  const { codePoints, families } = await collectVrtFontInventory();
  const manifest = createVrtFontManifest(codePoints, families);
  const fontFileNames = families.map((familyName, index) => createFontFileName(familyName, index));
  const fontPaths = fontFileNames.map((fileName) => join(VRT_FONT_DIR, fileName));
  if (
    fontPaths.every((fontPath) => existsSync(fontPath)) &&
    (await hasExpectedFontFiles(fontFileNames)) &&
    (await readTextIfExists(VRT_FONT_MANIFEST)) === manifest
  ) {
    return VRT_FONT_DIR;
  }

  await removeStaleFontFiles(fontFileNames);
  await Promise.all(
    families.map(async (familyName, index) => {
      const fontPath = fontPaths[index];
      const buffer = await createVrtFontBuffer(familyName, codePoints);
      await writeFile(fontPath, Buffer.from(buffer));
    }),
  );
  await writeFile(VRT_FONT_MANIFEST, manifest);
  return VRT_FONT_DIR;
}

function createVrtFontManifest(
  codePoints: ReadonlySet<number>,
  families: readonly string[],
): string {
  return JSON.stringify({
    generatorVersion: VRT_FONT_GENERATOR_VERSION,
    families,
    codePoints: [...codePoints].sort((a, b) => a - b),
  });
}

function createFontFileName(familyName: string, index: number): string {
  const safeFamilyName = familyName.replace(/[^A-Za-z0-9_-]+/g, "_") || "font";
  return `${index.toString().padStart(2, "0")}-${safeFamilyName}.ttf`;
}

async function hasExpectedFontFiles(expectedFileNames: readonly string[]): Promise<boolean> {
  const expected = new Set(expectedFileNames);
  const entries = await readdir(VRT_FONT_DIR);
  return entries.filter((entry) => entry.endsWith(".ttf")).every((entry) => expected.has(entry));
}

async function removeStaleFontFiles(expectedFileNames: readonly string[]): Promise<void> {
  const expected = new Set(expectedFileNames);
  const entries = await readdir(VRT_FONT_DIR);
  await Promise.all(
    entries
      .filter((entry) => entry.endsWith(".ttf") && !expected.has(entry))
      .map((entry) => unlink(join(VRT_FONT_DIR, entry))),
  );
}

async function readTextIfExists(path: string): Promise<string | null> {
  try {
    return await readFile(path, "utf8");
  } catch {
    return null;
  }
}

async function collectVrtFontInventory(): Promise<VrtFontInventory> {
  const codePoints = new Set<number>();
  const families = new Set<string>(BASE_VRT_FONT_FAMILIES);
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
                  addTextCodePoints(codePoints, decodeXmlText(match[1]));
                }
                for (const match of xml.matchAll(/<c:v[^>]*>(.*?)<\/c:v>/gs)) {
                  addTextCodePoints(codePoints, decodeXmlText(match[1]));
                }
                for (const match of xml.matchAll(/<a:buChar\b[^>]*\bchar="([^"]*)"/g)) {
                  addTextCodePoints(codePoints, decodeXmlText(match[1]));
                }
                for (const match of xml.matchAll(/\btypeface="([^"]*)"/g)) {
                  addVrtFontFamily(families, decodeXmlText(match[1]));
                }
              }),
          );
        }),
    );
  }

  const baseFamilies = new Set<string>(BASE_VRT_FONT_FAMILIES);
  const extraFamilies = [...families]
    .filter((familyName) => !baseFamilies.has(familyName))
    .sort((a, b) => a.localeCompare(b));
  return { codePoints, families: [...BASE_VRT_FONT_FAMILIES, ...extraFamilies] };
}

function addVrtFontFamily(families: Set<string>, familyName: string): void {
  const trimmed = familyName.trim();
  if (!trimmed || trimmed.startsWith("+")) return;
  if (getMappedFont(trimmed, VRT_FONT_MAPPING) !== null) return;
  families.add(trimmed);
}

function addTextCodePoints(codePoints: Set<number>, text: string): void {
  for (const char of text) {
    const codePoint = char.codePointAt(0);
    if (codePoint !== undefined) codePoints.add(codePoint);
  }
}

function decodeXmlText(text: string): string {
  return text.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (entity, body) => {
    if (typeof body !== "string") return entity;
    if (body.startsWith("#x")) {
      const codePoint = Number.parseInt(body.slice(2), 16);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : entity;
    }
    if (body.startsWith("#")) {
      const codePoint = Number.parseInt(body.slice(1), 10);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : entity;
    }
    switch (body) {
      case "amp":
        return "&";
      case "lt":
        return "<";
      case "gt":
        return ">";
      case "quot":
        return '"';
      case "apos":
        return "'";
      default:
        return entity;
    }
  });
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
    const advanceWidth = getGlyphAdvanceWidth(codePoint, isWide);
    glyphs.push(
      new opentype.Glyph({
        name: `uni${codePoint.toString(16).toUpperCase().padStart(4, "0")}`,
        unicode: codePoint,
        advanceWidth: isControl ? 0 : advanceWidth,
        path: isControl ? new opentype.Path() : createGlyphPath(codePoint, isWide),
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

  function getGlyphAdvanceWidth(codePoint: number, isWide: boolean): number {
    const hash = hashCodePoint(codePoint);
    return isWide ? 940 + (hash % 121) : 560 + (hash % 101);
  }

  function createGlyphPath(codePoint: number, isWide: boolean): InstanceType<typeof opentype.Path> {
    return isWide ? createWideGlyphPath(codePoint) : createNarrowGlyphPath(codePoint);
  }

  function createNarrowGlyphPath(codePoint: number): InstanceType<typeof opentype.Path> {
    const hash = hashCodePoint(codePoint);
    const left = 50 + (hash % 35);
    const right = 530 - ((hash >> 3) % 45);
    const peakX = 250 + ((hash >> 6) % 130);
    const peakY = 620 + ((hash >> 9) % 95);
    const notchY = 30 + ((hash >> 12) % 90);
    const path = new opentype.Path();
    path.moveTo(left, 0);
    path.lineTo(peakX, peakY);
    path.lineTo(right, 0);
    path.lineTo(350 + ((hash >> 15) % 80), notchY);
    path.lineTo(180 + ((hash >> 18) % 80), notchY);
    path.close();
    return path;
  }

  function createWideGlyphPath(codePoint: number): InstanceType<typeof opentype.Path> {
    const hash = hashCodePoint(codePoint);
    const left = 70 + (hash % 45);
    const right = 930 - ((hash >> 3) % 60);
    const top = 690 + ((hash >> 6) % 80);
    const shoulder = 170 + ((hash >> 9) % 180);
    const notchX = 440 + ((hash >> 12) % 160);
    const notchY = 40 + ((hash >> 15) % 120);
    const path = new opentype.Path();
    path.moveTo(left, 0);
    path.lineTo(left, top - shoulder);
    path.lineTo(notchX, top);
    path.lineTo(right, top - ((hash >> 18) % 90));
    path.lineTo(right, 0);
    path.lineTo(notchX, notchY);
    path.close();
    return path;
  }

  function hashCodePoint(codePoint: number): number {
    let hash = codePoint >>> 0;
    for (const char of familyName) {
      hash ^= char.codePointAt(0) ?? 0;
      hash = Math.imul(hash, 0x45d9f3b);
    }
    hash ^= hash >>> 16;
    hash = Math.imul(hash, 0x7feb352d);
    hash ^= hash >>> 15;
    hash = Math.imul(hash, 0x846ca68b);
    hash ^= hash >>> 16;
    return hash >>> 0;
  }
}
