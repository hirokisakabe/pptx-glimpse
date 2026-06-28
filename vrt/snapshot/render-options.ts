import { createHash } from "node:crypto";
import { existsSync } from "node:fs";
import { mkdir, readdir, readFile, unlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";

import type { ConvertOptions } from "../../packages/core/src/converter.js";
import { createFontMapping, getMappedFont } from "../../packages/renderer/src/font/font-mapping.js";
import { subsetFont } from "../../packages/renderer/src/font/font-subsetter.js";
import type { OpentypeFullFont } from "../../packages/renderer/src/font/text-path-context.js";
import { unsafeVrtInteropAssertion } from "../unsafe-type-assertion.js";

type VrtRenderOptions = Pick<ConvertOptions, "fontDirs" | "skipSystemFonts">;

const __dirname = dirname(fileURLToPath(import.meta.url));

const VRT_FONT_GENERATOR_VERSION = 7;
const VRT_FONT_DIR = join(tmpdir(), `pptx-glimpse-vrt-fonts-v${VRT_FONT_GENERATOR_VERSION}`);
const VRT_FONT_SOURCE_DIR = join(
  tmpdir(),
  `pptx-glimpse-vrt-font-sources-v${VRT_FONT_GENERATOR_VERSION}`,
);
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

type VrtFontSourceId =
  | "carlito"
  | "arimo"
  | "tinos"
  | "cousine"
  | "caladea"
  | "notoSansJp"
  | "notoSansCjkJp"
  | "notoSerifCjkJp";

interface VrtFontSource {
  id: VrtFontSourceId;
  fileName: string;
  url: string;
  sha256: string;
}

interface VrtFontFamilySource {
  familyName: string;
  sourceId: VrtFontSourceId;
}

const GOOGLE_FONTS_REVISION = "295d98a7a0c17c68f1341eaeea354e7960ea70d3";
const CARLITO_REVISION = "3a810cab78ebd6e2e4eed42af9e8453c4f9b850a";
const NOTO_CJK_REVISION = "f8d157532fbfaeda587e826d4cd5b21a49186f7c";

const VRT_FONT_SOURCES = [
  {
    id: "carlito",
    fileName: "Carlito-Regular.ttf",
    url: `https://raw.githubusercontent.com/googlefonts/carlito/${CARLITO_REVISION}/fonts/ttf/Carlito-Regular.ttf`,
    sha256: "f6418f708baede9789daef5d458c0f53d2a888af9820e8062934e504fedc6595",
  },
  {
    id: "arimo",
    fileName: "Arimo.ttf",
    url: `https://raw.githubusercontent.com/google/fonts/${GOOGLE_FONTS_REVISION}/apache/arimo/Arimo%5Bwght%5D.ttf`,
    sha256: "c75270dfa8b5747c666d9e1915b8c9a6cb6e2de74d103cd0f6ee0104675a3604",
  },
  {
    id: "tinos",
    fileName: "Tinos-Regular.ttf",
    url: `https://raw.githubusercontent.com/google/fonts/${GOOGLE_FONTS_REVISION}/apache/tinos/Tinos-Regular.ttf`,
    sha256: "1061395ac6775f3cea27dc9ef3d7a3b9cc34c2b4a2d97aa649411294d5165990",
  },
  {
    id: "cousine",
    fileName: "Cousine-Regular.ttf",
    url: `https://raw.githubusercontent.com/google/fonts/${GOOGLE_FONTS_REVISION}/apache/cousine/Cousine-Regular.ttf`,
    sha256: "69e1ea59eb770014204e5174f805750f9a793db4a2531e6516b30b7460d470b3",
  },
  {
    id: "caladea",
    fileName: "Caladea-Regular.ttf",
    url: `https://raw.githubusercontent.com/google/fonts/${GOOGLE_FONTS_REVISION}/ofl/caladea/Caladea-Regular.ttf`,
    sha256: "f1e899278b7b4491aba5b6a8253c4b04c050cc59b21865be5c37559a775153cd",
  },
  {
    id: "notoSansJp",
    fileName: "NotoSansJP.ttf",
    url: `https://raw.githubusercontent.com/google/fonts/${GOOGLE_FONTS_REVISION}/ofl/notosansjp/NotoSansJP%5Bwght%5D.ttf`,
    sha256: "c2f3b4d463500a2ddcd3849cded1fceeb9fd6d1c32e6cbecd568453ba50fc68f",
  },
  {
    id: "notoSansCjkJp",
    fileName: "NotoSansCJKjp-Regular.otf",
    url: `https://raw.githubusercontent.com/notofonts/noto-cjk/${NOTO_CJK_REVISION}/Sans/OTF/Japanese/NotoSansCJKjp-Regular.otf`,
    sha256: "68a3fc98800b2a27b371f2fb79991daf3633bd89309d4ffaa6946fd587f375b5",
  },
  {
    id: "notoSerifCjkJp",
    fileName: "NotoSerifCJKjp-Regular.otf",
    url: `https://raw.githubusercontent.com/notofonts/noto-cjk/${NOTO_CJK_REVISION}/Serif/OTF/Japanese/NotoSerifCJKjp-Regular.otf`,
    sha256: "d9854c7a8ef170b5a7932558856fd64eb8de0b007cd823fed6f9f514ad2803d3",
  },
] as const satisfies readonly VrtFontSource[];

const VRT_FONT_FAMILY_SOURCES = [
  {
    familyName: "Carlito",
    sourceId: "carlito",
  },
  {
    familyName: "Arimo",
    sourceId: "arimo",
  },
  {
    familyName: "Tinos",
    sourceId: "tinos",
  },
  {
    familyName: "Cousine",
    sourceId: "cousine",
  },
  {
    familyName: "Caladea",
    sourceId: "caladea",
  },
  {
    familyName: "Noto Sans JP",
    sourceId: "notoSansJp",
  },
  {
    familyName: "Noto Sans CJK JP",
    sourceId: "notoSansCjkJp",
  },
  {
    familyName: "Noto Serif CJK JP",
    sourceId: "notoSerifCjkJp",
  },
] as const satisfies readonly VrtFontFamilySource[];

interface VrtFontInventory {
  codePoints: Set<number>;
  families: string[];
}

let renderOptionsPromise: Promise<VrtRenderOptions> | null = null;

// Local snapshot VRT must not depend on parsing every OS system font. Build a
// small deterministic directory of real-font subsets and use it for both Docker
// snapshots and local VRT, preserving readable Latin/CJK visual coverage.
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
  const sourceFonts = await loadSourceFonts();
  await Promise.all(
    families.map(async (familyName, index) => {
      const fontPath = fontPaths[index];
      const sourceFont = selectSourceFont(familyName, sourceFonts);
      const buffer = await subsetFont(sourceFont, codePointsToChars(codePoints), familyName);
      if (buffer === null) {
        throw new Error(`Failed to create VRT font subset for ${familyName}.`);
      }
      await writeFile(fontPath, buffer);
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
    sources: VRT_FONT_SOURCES.map(({ fileName, sha256, url }) => ({ fileName, sha256, url })),
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

async function loadSourceFonts(): Promise<Map<VrtFontSourceId, OpentypeFullFont>> {
  await mkdir(VRT_FONT_SOURCE_DIR, { recursive: true });
  const opentype: { parse: (buffer: ArrayBuffer) => OpentypeFullFont } =
    await import("opentype.js");
  const fonts = new Map<VrtFontSourceId, OpentypeFullFont>();
  await Promise.all(
    VRT_FONT_SOURCES.map(async (source) => {
      const sourcePath = await ensureSourceFont(source);
      const sourceBuffer = await readFile(sourcePath);
      fonts.set(source.id, opentype.parse(toArrayBuffer(sourceBuffer)));
    }),
  );
  return fonts;
}

async function ensureSourceFont(source: VrtFontSource): Promise<string> {
  const path = join(VRT_FONT_SOURCE_DIR, source.fileName);
  const existing = await readBufferIfExists(path);
  if (existing !== null && sha256(existing) === source.sha256) return path;

  const response = await fetch(source.url);
  if (!response.ok) {
    throw new Error(`Failed to download VRT source font ${source.fileName}: ${response.status}`);
  }
  const buffer = Buffer.from(await response.arrayBuffer());
  const actualSha256 = sha256(buffer);
  if (actualSha256 !== source.sha256) {
    throw new Error(
      `Downloaded VRT source font ${source.fileName} has sha256 ${actualSha256}, expected ${source.sha256}.`,
    );
  }
  await writeFile(path, buffer);
  return path;
}

async function readBufferIfExists(path: string): Promise<Buffer | null> {
  try {
    return await readFile(path);
  } catch {
    return null;
  }
}

function toArrayBuffer(buffer: Uint8Array): ArrayBuffer {
  return unsafeVrtInteropAssertion<ArrayBuffer>(
    buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength),
  );
}

function sha256(buffer: Uint8Array): string {
  return createHash("sha256").update(buffer).digest("hex");
}

function selectSourceFont(
  familyName: string,
  sourceFonts: ReadonlyMap<VrtFontSourceId, OpentypeFullFont>,
): OpentypeFullFont {
  const explicit = VRT_FONT_FAMILY_SOURCES.find((source) => source.familyName === familyName);
  const sourceId = explicit?.sourceId ?? inferSourceId(familyName);
  const font = sourceFonts.get(sourceId);
  if (font === undefined) {
    throw new Error(`VRT source font not loaded for ${sourceId}.`);
  }
  return font;
}

function inferSourceId(familyName: string): VrtFontSourceId {
  const lower = familyName.toLowerCase();
  if (lower.includes("mono") || lower.includes("courier") || lower.includes("cousine")) {
    return "cousine";
  }
  if (lower.includes("serif") || lower.includes("times") || lower.includes("mincho")) {
    return lower.includes("noto") ? "notoSerifCjkJp" : "tinos";
  }
  if (lower.includes("noto") || lower.includes("gothic") || lower.includes("meiryo")) {
    return "notoSansJp";
  }
  return "arimo";
}

function codePointsToChars(codePoints: ReadonlySet<number>): Set<string> {
  const chars = new Set<string>();
  for (const codePoint of codePoints) {
    if (codePoint === 0 || codePoint === 0x0a) continue;
    chars.add(String.fromCodePoint(codePoint));
  }
  return chars;
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
