/**
 * Keeps the font mapping table used during rendering at module level.
 * Same as setTextMeasurer/resetTextMeasurer pattern in text-measurer.ts.
 */
import type { FontMapping } from "./font-mapping.js";
import { DEFAULT_FONT_MAPPING, getMappedFont } from "./font-mapping.js";

let currentMapping: FontMapping = { ...DEFAULT_FONT_MAPPING };

export function setFontMapping(mapping: FontMapping): void {
  currentMapping = mapping;
}

export function resetFontMapping(): void {
  currentMapping = { ...DEFAULT_FONT_MAPPING };
}

/**
 * Shortcut for getting an OSS replacement font name from the current mapping table.
 */
export function getCurrentMappedFont(fontFamily: string | null | undefined): string | null {
  return getMappedFont(fontFamily, currentMapping);
}
