/**
 * Helper that reads fonts and builds an OpentypeTextMeasurer using opentype.js.
 */
import { readFile } from "node:fs/promises";

import { unsafeExternalInteropAssertion } from "../unsafe-type-assertion.js";
import type { FontMapping } from "./font-mapping.js";
import { createFontMapping } from "./font-mapping.js";
import type { OpentypeFont } from "./opentype-text-measurer.js";
import { OpentypeTextMeasurer } from "./opentype-text-measurer.js";
import { collectFontFilePaths } from "./system-font-loader.js";
import type { OpentypeFullFont, TextPathFontResolver } from "./text-path-context.js";
import { DefaultTextPathFontResolver } from "./text-path-context.js";
import { extractTtcFonts, isTtcBuffer } from "./ttc-parser.js";

/** Font buffer input format */
export interface FontBuffer {
  name?: string;
  data: ArrayBuffer | Uint8Array;
}

interface OpentypeFontWithNames extends OpentypeFont {
  names: {
    fontFamily?: Record<string, string>;
    preferredFamily?: Record<string, string>;
  };
}

/**
 * Load opentype.js with dynamic import.
 * Returns null if opentype.js is not installed.
 */
async function tryLoadOpentype(): Promise<{
  parse: (buffer: ArrayBuffer) => OpentypeFontWithNames;
} | null> {
  try {
    // Use a variable to prevent bundlers from statically resolving this import
    const specifier = "opentype.js";
    // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
    const mod: { parse: (buffer: ArrayBuffer) => OpentypeFontWithNames } = await import(
      /* @vite-ignore */ specifier
    );
    return { parse: mod.parse };
  } catch {
    return null;
  }
}

/**
 * Build a reverse lookup table for font mapping.
 * Mapping of OSS font name -> PPTX font name[].
 * Example: "Carlito" -> ["Calibri"]
 */
function buildReverseMapping(mapping: FontMapping): Map<string, string[]> {
  const reverse = new Map<string, string[]>();
  for (const [pptxName, ossName] of Object.entries(mapping)) {
    const existing = reverse.get(ossName) ?? [];
    existing.push(pptxName);
    reverse.set(ossName, existing);
  }
  return reverse;
}

/**
 * ArrayBuffer | Convert from Uint8Array to ArrayBuffer.
 * For Uint8Array, use slice to obtain an independent ArrayBuffer.
 */
function toArrayBuffer(data: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (data instanceof ArrayBuffer) return data;
  return unsafeExternalInteropAssertion<ArrayBuffer>(
    data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength),
  );
}

/**
 * Returns a parsed font array from a buffer (TTF/OTF or TTC).
 * In the case of TTC, only the first font is extracted and parsed to reduce memory consumption.
 */
function parseFontBuffer(
  arrayBuffer: ArrayBuffer,
  opentype: { parse: (buffer: ArrayBuffer) => OpentypeFontWithNames },
): OpentypeFontWithNames[] {
  if (isTtcBuffer(arrayBuffer)) {
    // Only the first font is extracted from TTC.
    // CJK TTC (such as NotoSansCJK) consumes hundreds of MB of memory when all fonts are expanded.
    const fonts = extractTtcFonts(arrayBuffer);
    if (fonts.length > 0) {
      try {
        return [opentype.parse(fonts[0])];
      } catch {
        // Skip parse failure
      }
    }
    return [];
  }
  return [opentype.parse(arrayBuffer)];
}

/**
 * Construct an OpentypeTextMeasurer from a font buffer array.
 *
 * Dynamically import opentype.js internally to parse the font.
 * Returns null if opentype.js is not available.
 */
export async function createOpentypeTextMeasurerFromBuffers(
  fontBuffers: FontBuffer[],
  fontMapping?: FontMapping,
): Promise<OpentypeTextMeasurer | null> {
  const setup = await createOpentypeSetupFromBuffers(fontBuffers, fontMapping);
  return setup?.measurer ?? null;
}

export interface OpentypeSetup {
  measurer: OpentypeTextMeasurer;
  fontResolver: TextPathFontResolver;
}

/**
 * Construct OpentypeTextMeasurer and TextPathFontResolver simultaneously from the font buffer array.
 *
 * The object returned by opentype.parse() satisfies both OpentypeFont and OpentypeFullFont, so
 * Pass the same Font object to both measurer and fontResolver.
 */
export async function createOpentypeSetupFromBuffers(
  fontBuffers: FontBuffer[],
  fontMapping?: FontMapping,
): Promise<OpentypeSetup | null> {
  if (fontBuffers.length === 0) return null;

  const opentype = await tryLoadOpentype();
  if (!opentype) return null;

  const mapping = createFontMapping(fontMapping);
  const reverseMap = buildReverseMapping(mapping);
  const measurerFonts = new Map<string, OpentypeFont>();
  const resolverFonts = new Map<string, OpentypeFullFont>();
  let firstMeasurerFont: OpentypeFont | null = null;
  let firstResolverFont: OpentypeFullFont | null = null;

  for (const buffer of fontBuffers) {
    try {
      const arrayBuffer = toArrayBuffer(buffer.data);
      const isTtc = isTtcBuffer(arrayBuffer);
      const fonts = parseFontBuffer(arrayBuffer, opentype);

      for (const font of fonts) {
        if (!firstMeasurerFont) firstMeasurerFont = font;
        if (!firstResolverFont)
          firstResolverFont = unsafeExternalInteropAssertion<OpentypeFullFont>(font);

        if (isTtc) {
          // TTC: Get and register font name from names table
          for (const name of collectFontNames(font)) {
            registerFont(name, font, reverseMap, measurerFonts, resolverFonts);
          }
        } else if (buffer.name) {
          registerFont(buffer.name, font, reverseMap, measurerFonts, resolverFonts);
        }
      }
    } catch {
      // Skip fonts that fail parsing
    }
  }

  if (measurerFonts.size === 0 && !firstMeasurerFont) return null;

  const measurer = new OpentypeTextMeasurer(measurerFonts, firstMeasurerFont ?? undefined);
  const fontResolver = new DefaultTextPathFontResolver(
    resolverFonts,
    firstResolverFont ?? undefined,
  );

  return { measurer, fontResolver };
}

function registerFont(
  name: string,
  font: OpentypeFontWithNames,
  reverseMap: Map<string, string[]>,
  measurerFonts: Map<string, OpentypeFont>,
  resolverFonts: Map<string, OpentypeFullFont>,
): void {
  const fullFont = unsafeExternalInteropAssertion<OpentypeFullFont>(font);
  if (!measurerFonts.has(name)) {
    measurerFonts.set(name, font);
    resolverFonts.set(name, fullFont);
  }

  // Also register PPTX font name by reverse lookup
  const pptxNames = reverseMap.get(name);
  if (pptxNames) {
    for (const pptxName of pptxNames) {
      if (!measurerFonts.has(pptxName)) {
        measurerFonts.set(pptxName, font);
        resolverFonts.set(pptxName, fullFont);
      }
    }
  }
}

/**
 * Gather a set of font names from the font's names table.
 * Include both fontFamily and preferredFamily.
 * For Variable Font, fontFamily is like "Noto Sans JP Thin"
 * Since this is the instance name, also register preferredFamily ("Noto Sans JP").
 */
function collectFontNames(font: OpentypeFontWithNames): Set<string> {
  const names = new Set<string>();
  if (font.names.fontFamily) {
    for (const name of Object.values(font.names.fontFamily)) {
      names.add(name);
    }
  }
  if (font.names.preferredFamily) {
    for (const name of Object.values(font.names.preferredFamily)) {
      names.add(name);
    }
  }
  return names;
}

/**
 * Generate a cache key. Uniquely identified by the combination of fontDirs and fontMapping.
 */
function buildCacheKey(
  additionalFontDirs?: string[],
  fontMapping?: FontMapping,
  skipSystemFonts = false,
): string {
  const dirsKey = additionalFontDirs ? [...additionalFontDirs].sort().join("\0") : "";
  const mappingKey = fontMapping
    ? JSON.stringify(fontMapping, Object.keys(fontMapping).sort())
    : "";
  return `${dirsKey}\n${mappingKey}\n${skipSystemFonts}`;
}

/** Caching parsed Font objects */
let cachedSetup: OpentypeSetup | null = null;
let cachedSetupKey: string | null = null;

/**
 * Clear the font object cache.
 * Normally there is no need to call it, but after installing/uninstalling a font
 * Use this when you want to force reload.
 */
export function clearFontCache(): void {
  cachedSetup = null;
  cachedSetupKey = null;
}

/**
 * Build OpentypeTextMeasurer and TextPathFontResolver from system fonts + additional directories.
 *
 * 1. Collect font file paths with collectFontFilePaths()
 * 2. Parse each file with readFile + opentype.parse
 * 3. Register the font name in the map as a key (including reverse mapping)
 *
 * Parsed Font objects are cached at the module level and
 * Subsequent calls with the same fontDirs / fontMapping return the cache.
 */
export async function createOpentypeSetupFromSystem(
  additionalFontDirs?: string[],
  fontMapping?: FontMapping,
  skipSystemFonts = false,
): Promise<OpentypeSetup | null> {
  const key = buildCacheKey(additionalFontDirs, fontMapping, skipSystemFonts);
  if (cachedSetup && cachedSetupKey === key) {
    return cachedSetup;
  }

  const opentype = await tryLoadOpentype();
  if (!opentype) return null;

  const fontFilePaths = collectFontFilePaths(additionalFontDirs, skipSystemFonts);
  if (fontFilePaths.length === 0) return null;

  const mapping = createFontMapping(fontMapping);
  const reverseMap = buildReverseMapping(mapping);
  const measurerFonts = new Map<string, OpentypeFont>();
  const resolverFonts = new Map<string, OpentypeFullFont>();
  let firstMeasurerFont: OpentypeFont | null = null;
  let firstResolverFont: OpentypeFullFont | null = null;

  for (const filePath of fontFilePaths) {
    try {
      const data = await readFile(filePath);
      const arrayBuffer = toArrayBuffer(data);
      const fonts = parseFontBuffer(arrayBuffer, opentype);

      for (const font of fonts) {
        if (!firstMeasurerFont) firstMeasurerFont = font;
        if (!firstResolverFont)
          firstResolverFont = unsafeExternalInteropAssertion<OpentypeFullFont>(font);

        // Get the font name from the names table and register it
        for (const name of collectFontNames(font)) {
          registerFont(name, font, reverseMap, measurerFonts, resolverFonts);
        }
      }
    } catch {
      // Skip fonts that fail parsing
    }
  }

  if (measurerFonts.size === 0 && !firstMeasurerFont) return null;

  const measurer = new OpentypeTextMeasurer(measurerFonts, firstMeasurerFont ?? undefined);
  const fontResolver = new DefaultTextPathFontResolver(
    resolverFonts,
    firstResolverFont ?? undefined,
  );

  const setup = { measurer, fontResolver };
  cachedSetup = setup;
  cachedSetupKey = key;

  return setup;
}
