/**
 * Internal note.
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

/** Internal note. */
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
 * Internal note.
 * Internal note.
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
 * Internal note.
 * Internal note.
 * Example: "Carlito" → ["Calibri"]
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
 * Internal note.
 * Internal note.
 */
function toArrayBuffer(data: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (data instanceof ArrayBuffer) return data;
  return unsafeExternalInteropAssertion<ArrayBuffer>(
    data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength),
  );
}

/**
 * Internal note.
 * Internal note.
 */
function parseFontBuffer(
  arrayBuffer: ArrayBuffer,
  opentype: { parse: (buffer: ArrayBuffer) => OpentypeFontWithNames },
): OpentypeFontWithNames[] {
  if (isTtcBuffer(arrayBuffer)) {
    // Internal note.
    // Internal note.
    const fonts = extractTtcFonts(arrayBuffer);
    if (fonts.length > 0) {
      try {
        return [opentype.parse(fonts[0])];
      } catch {
        // Internal note.
      }
    }
    return [];
  }
  return [opentype.parse(arrayBuffer)];
}

/**
 * Internal note.
 *
 * Internal note.
 * Internal note.
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
 * Internal note.
 *
 * Internal note.
 * Internal note.
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
          // Internal note.
          for (const name of collectFontNames(font)) {
            registerFont(name, font, reverseMap, measurerFonts, resolverFonts);
          }
        } else if (buffer.name) {
          registerFont(buffer.name, font, reverseMap, measurerFonts, resolverFonts);
        }
      }
    } catch {
      // Internal note.
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

  // Internal note.
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
 * Internal note.
 * Internal note.
 * Internal note.
 * Internal note.
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
 * Internal note.
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

/** Internal note. */
let cachedSetup: OpentypeSetup | null = null;
let cachedSetupKey: string | null = null;

/**
 * Internal note.
 * Internal note.
 * Internal note.
 */
export function clearFontCache(): void {
  cachedSetup = null;
  cachedSetupKey = null;
}

/**
 * Internal note.
 *
 * Internal note.
 * Internal note.
 * Internal note.
 *
 * Internal note.
 * Internal note.
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

        // Internal note.
        for (const name of collectFontNames(font)) {
          registerFont(name, font, reverseMap, measurerFonts, resolverFonts);
        }
      }
    } catch {
      // Internal note.
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
