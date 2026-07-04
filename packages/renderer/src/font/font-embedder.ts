/**
 * Construct a <style> element with @font-face definition from font usage.
 * Used to embed subsetted fonts in SVG in native <text> output mode.
 */

import { uint8ArrayToBase64 } from "../utils/base64.js";
import { subsetFont } from "./font-subsetter.js";
import type { FontUsage } from "./font-usage-collector.js";
import { getJpanFallbackFont } from "./script-font-context.js";
import type { TextPathFontResolver } from "./text-path-context.js";

/**
 * Escapes a font name for CSS string literals.
 * Since the content of <style> is also an XML text node, XML special characters are represented using CSS hexadecimal escapes.
 */
function escapeCssFamilyName(name: string): string {
  return name
    .replace(/[\\"]/g, (c) => `\\${c}`)
    .replace(/[<>&]/g, (c) => `\\${c.codePointAt(0)!.toString(16)} `);
}

/**
 * Subset the collected font usage and return <style> elements with @font-face definitions.
 * If there are no embeddable fonts, an empty string is returned.
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

    const base64 = uint8ArrayToBase64(buffer);
    faces.push(
      `@font-face{font-family:"${escapeCssFamilyName(familyName)}";src:url(data:font/otf;base64,${base64}) format("opentype");}`,
    );
  }

  if (faces.length === 0) return "";
  return `<style type="text/css">${faces.join("")}</style>`;
}
