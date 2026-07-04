import { readFileSync, statSync } from "node:fs";

import { collectFontFilePaths } from "@pptx-glimpse/renderer/node";

let cachedFontBuffers: Uint8Array[] | null = null;
let cachedFontBuffersKey: string | null = null;

const MAX_TOTAL_FONT_BUFFER_BYTES = 100 * 1024 * 1024;

export function loadFontBuffersFromSystem(
  fontDirs?: string[],
  skipSystemFonts?: boolean,
): Uint8Array[] {
  const key = `${(fontDirs ?? []).join("\0")}\n${skipSystemFonts ?? false}`;
  if (cachedFontBuffers !== null && cachedFontBuffersKey === key) {
    return cachedFontBuffers;
  }

  const fontPaths = collectFontFilePaths(fontDirs, skipSystemFonts).filter((path) => {
    const lower = path.toLowerCase();
    return lower.endsWith(".ttf") || lower.endsWith(".otf");
  });
  const readableFontPaths: { path: string; size: number }[] = [];
  for (const path of fontPaths) {
    try {
      readableFontPaths.push({ path, size: statSync(path).size });
    } catch {
      // Ignore unreadable font files.
    }
  }
  readableFontPaths.sort((a, b) => a.size - b.size || a.path.localeCompare(b.path));

  const buffers: Uint8Array[] = [];
  let totalSize = 0;
  for (const { path, size } of readableFontPaths) {
    if (totalSize + size > MAX_TOTAL_FONT_BUFFER_BYTES) break;
    try {
      buffers.push(new Uint8Array(readFileSync(path)));
      totalSize += size;
    } catch {
      // Ignore fonts that disappear between stat and read.
    }
  }

  cachedFontBuffers = buffers;
  cachedFontBuffersKey = key;
  return buffers;
}
