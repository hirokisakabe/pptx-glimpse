/**
 * Internal note.
 * Internal note.
 */

import { Buffer } from "node:buffer";

import { subsetFont } from "./font-subsetter.js";
import type { FontUsage } from "./font-usage-collector.js";
import { getJpanFallbackFont } from "./script-font-context.js";
import type { TextPathFontResolver } from "./text-path-context.js";

/**
 * Escapes a font name for CSS string literals.
 * Internal note.
 */
function escapeCssFamilyName(name: string): string {
  return name
    .replace(/[\\"]/g, (c) => `\\${c}`)
    .replace(/[<>&]/g, (c) => `\\${c.codePointAt(0)!.toString(16)} `);
}

/**
 * Internal note.
 * Internal note.
 */
export async function buildFontFaceStyle(
  usages: Map<string, FontUsage>,
  fontResolver: TextPathFontResolver,
): Promise<string> {
  const faces: string[] = [];
  const jpanFallback = getJpanFallbackFont();

  for (const [familyName, usage] of usages) {
    const font = fontResolver.resolveFont(
      usage.fonts[0],
      usage.fonts[1] ?? null,
      usage.fonts[2] ?? jpanFallback,
    );
    if (!font) continue;

    const buffer = await subsetFont(font, usage.chars, familyName);
    if (!buffer) continue;

    const base64 = Buffer.from(buffer).toString("base64");
    faces.push(
      `@font-face{font-family:"${escapeCssFamilyName(familyName)}";src:url(data:font/otf;base64,${base64}) format("opentype");}`,
    );
  }

  if (faces.length === 0) return "";
  return `<style type="text/css">${faces.join("")}</style>`;
}
