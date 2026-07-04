/**
 * Node.js-only helpers for building OpenType-backed font services from files.
 */
import { readFile } from "node:fs/promises";

import type { FontMapping } from "./font-mapping.js";
import { createFontMapping } from "./font-mapping.js";
import type { OpentypeSetup } from "./opentype-buffer-helpers.js";
import {
  buildOpentypeSetupFromState,
  buildReverseMapping,
  createOpentypeSetupState,
  getCachedSystemOpentypeSetup,
  parseFontBuffer,
  registerParsedFont,
  setCachedSystemOpentypeSetup,
  toArrayBuffer,
  tryLoadOpentype,
} from "./opentype-buffer-helpers.js";
import { collectFontFilePaths } from "./system-font-loader.js";

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
  const cached = getCachedSystemOpentypeSetup(key);
  if (cached !== undefined) {
    return cached;
  }

  const opentype = await tryLoadOpentype();
  if (!opentype) return null;

  const fontFilePaths = collectFontFilePaths(additionalFontDirs, skipSystemFonts);
  if (fontFilePaths.length === 0) return null;

  const reverseMap = buildReverseMapping(createFontMapping(fontMapping));
  const state = createOpentypeSetupState();

  for (const filePath of fontFilePaths) {
    try {
      const data = await readFile(filePath);
      const arrayBuffer = toArrayBuffer(data);
      const fonts = parseFontBuffer(arrayBuffer, opentype);

      for (const font of fonts) {
        registerParsedFont(font, reverseMap, state);
      }
    } catch {
      // Skip fonts that fail parsing
    }
  }

  const setup = buildOpentypeSetupFromState(state);
  if (setup === null) return null;

  setCachedSystemOpentypeSetup(key, setup);
  return setup;
}
