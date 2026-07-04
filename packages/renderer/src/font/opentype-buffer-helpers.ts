/**
 * Helpers that build OpenType-backed font services from caller-provided bytes.
 *
 * This module is intentionally free of Node.js fs/path imports so it can be
 * used from browser and Edge Runtime bundles.
 */
import { unsafeExternalInteropAssertion } from "../unsafe-type-assertion.js";
import type { FontMapping } from "./font-mapping.js";
import { createFontMapping } from "./font-mapping.js";
import type { OpentypeFont } from "./opentype-text-measurer.js";
import { OpentypeTextMeasurer } from "./opentype-text-measurer.js";
import type { OpentypeFullFont, TextPathFontResolver } from "./text-path-context.js";
import { DefaultTextPathFontResolver } from "./text-path-context.js";
import { extractTtcFonts, isTtcBuffer } from "./ttc-parser.js";

/**
 * Font file data supplied directly by the caller.
 *
 * Use this when system font scanning is unavailable or undesirable, such as in
 * browsers, Edge Runtime, or serverless environments. `name` can be used to
 * register TTF/OTF buffers under a specific family name; family names from the
 * font names table are also registered when available.
 */
export interface FontBuffer {
  /**
   * Optional font family name to register for this buffer.
   */
  name?: string;
  /**
   * Raw TTF, OTF, or TTC font file bytes.
   */
  data: ArrayBuffer | Uint8Array;
}

interface OpentypeFontWithNames extends OpentypeFont {
  names: {
    fontFamily?: Record<string, string>;
    preferredFamily?: Record<string, string>;
  };
}

export interface OpentypeParser {
  parse: (buffer: ArrayBuffer) => OpentypeFontWithNames;
}

/**
 * Load opentype.js with dynamic import.
 * Returns null if opentype.js is not installed.
 */
export async function tryLoadOpentype(): Promise<OpentypeParser | null> {
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
export function buildReverseMapping(mapping: FontMapping): Map<string, string[]> {
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
export function toArrayBuffer(data: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (data instanceof ArrayBuffer) return data;
  return unsafeExternalInteropAssertion<ArrayBuffer>(
    data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength),
  );
}

/**
 * Returns a parsed font array from a buffer (TTF/OTF or TTC).
 * In the case of TTC, only the first font is extracted and parsed to reduce memory consumption.
 */
export function parseFontBuffer(
  arrayBuffer: ArrayBuffer,
  opentype: OpentypeParser,
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

export interface OpentypeSetupState {
  measurerFonts: Map<string, OpentypeFont>;
  resolverFonts: Map<string, OpentypeFullFont>;
  firstMeasurerFont: OpentypeFont | null;
  firstResolverFont: OpentypeFullFont | null;
}

export function createOpentypeSetupState(): OpentypeSetupState {
  return {
    measurerFonts: new Map<string, OpentypeFont>(),
    resolverFonts: new Map<string, OpentypeFullFont>(),
    firstMeasurerFont: null,
    firstResolverFont: null,
  };
}

export function registerParsedFont(
  font: OpentypeFontWithNames,
  reverseMap: Map<string, string[]>,
  state: OpentypeSetupState,
  explicitName?: string,
): void {
  if (!state.firstMeasurerFont) state.firstMeasurerFont = font;
  if (!state.firstResolverFont)
    state.firstResolverFont = unsafeExternalInteropAssertion<OpentypeFullFont>(font);

  if (explicitName) {
    registerFont(explicitName, font, reverseMap, state.measurerFonts, state.resolverFonts);
  }

  for (const name of collectFontNames(font)) {
    registerFont(name, font, reverseMap, state.measurerFonts, state.resolverFonts);
  }
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
 * Font services created from OpenType font data.
 */
export interface OpentypeSetup {
  /**
   * Measures text width for layout and wrapping.
   */
  measurer: OpentypeTextMeasurer;
  /**
   * Resolves fonts for converting text to SVG path outlines.
   */
  fontResolver: TextPathFontResolver;
}

export function buildOpentypeSetupFromState(state: OpentypeSetupState): OpentypeSetup | null {
  if (state.measurerFonts.size === 0 && !state.firstMeasurerFont) return null;

  const measurer = new OpentypeTextMeasurer(
    state.measurerFonts,
    state.firstMeasurerFont ?? undefined,
  );
  const fontResolver = new DefaultTextPathFontResolver(
    state.resolverFonts,
    state.firstResolverFont ?? undefined,
  );

  return { measurer, fontResolver };
}

/**
 * Create a text measurer from caller-provided font buffers.
 *
 * This low-level helper is useful when rendering in an environment where fonts
 * are bundled as bytes instead of discovered from the OS. It dynamically imports
 * `opentype.js` and returns `null` if no fonts can be parsed or the optional
 * dependency is unavailable.
 *
 * @param fontBuffers Font file bytes to parse.
 * @param fontMapping Optional PPTX font name to replacement font mapping.
 * @returns A text measurer, or `null` when setup is not possible.
 */
export async function createOpentypeTextMeasurerFromBuffers(
  fontBuffers: FontBuffer[],
  fontMapping?: FontMapping,
): Promise<OpentypeTextMeasurer | null> {
  const setup = await createOpentypeSetupFromBuffers(fontBuffers, fontMapping);
  return setup?.measurer ?? null;
}

/**
 * Create font measurement and SVG path resolution services from font buffers.
 *
 * Use this for integrations that need to provide fonts explicitly instead of
 * relying on Node.js system font directories.
 *
 * @param fontBuffers TTF, OTF, or TTC font file bytes.
 * @param fontMapping Optional PPTX font name to replacement font mapping.
 * @returns Font services, or `null` when no usable font setup can be created.
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
  const state = createOpentypeSetupState();

  for (const buffer of fontBuffers) {
    try {
      const arrayBuffer = toArrayBuffer(buffer.data);
      const fonts = parseFontBuffer(arrayBuffer, opentype);

      for (const font of fonts) {
        registerParsedFont(font, reverseMap, state, buffer.name);
      }
    } catch {
      // Skip fonts that fail parsing
    }
  }

  return buildOpentypeSetupFromState(state);
}

interface SystemFontCacheStore {
  setup: OpentypeSetup | null;
  key: string | null;
}

const SYSTEM_FONT_CACHE_KEY = "__pptxGlimpseSystemFontCache__";

function getSystemFontCacheStore(): SystemFontCacheStore {
  const globalObject = globalThis as typeof globalThis & {
    [SYSTEM_FONT_CACHE_KEY]?: SystemFontCacheStore;
  };
  globalObject[SYSTEM_FONT_CACHE_KEY] ??= { setup: null, key: null };
  return globalObject[SYSTEM_FONT_CACHE_KEY];
}

export function getCachedSystemOpentypeSetup(key: string): OpentypeSetup | null | undefined {
  const cache = getSystemFontCacheStore();
  return cache.key === key ? cache.setup : undefined;
}

export function setCachedSystemOpentypeSetup(key: string, setup: OpentypeSetup): void {
  const cache = getSystemFontCacheStore();
  cache.setup = setup;
  cache.key = key;
}

/**
 * Clear the module-level cache of parsed system fonts.
 *
 * Conversion APIs cache parsed font objects for repeated calls with the same
 * font options. Call this after installing, removing, or replacing fonts in a
 * long-running process when subsequent conversions must reload font files.
 */
export function clearFontCache(): void {
  const cache = getSystemFontCacheStore();
  cache.setup = null;
  cache.key = null;
}
