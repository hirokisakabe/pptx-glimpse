/**
 * Use opentype.js to subset the font to only used characters.
 * For @font-face embedding in native <text> output mode.
 */

import { unsafeExternalInteropAssertion } from "../unsafe-type-assertion.js";
import { warn } from "../warning-logger.js";
import type { OpentypeFullFont } from "./text-path-context.js";

interface OpentypeGlyph {
  index: number;
  name?: string | null;
  unicode?: number;
  unicodes?: number[];
  advanceWidth?: number;
  path: unknown;
}

interface SubsettableFont {
  unitsPerEm: number;
  ascender: number;
  descender: number;
  charToGlyph(char: string): OpentypeGlyph | null;
  glyphs: { get(index: number): OpentypeGlyph };
}

interface OpentypeCtors {
  Font: new (options: Record<string, unknown>) => { toArrayBuffer(): ArrayBuffer };
  Glyph: new (options: Record<string, unknown>) => unknown;
}

/**
 * Load the opentype.js constructor with dynamic import.
 * Returns null if opentype.js is not installed.
 */
async function tryLoadOpentypeCtors(): Promise<OpentypeCtors | null> {
  try {
    // Use a variable to prevent bundlers from statically resolving this import
    const specifier = "opentype.js";
    // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
    const mod: OpentypeCtors = await import(/* @vite-ignore */ specifier);
    return { Font: mod.Font, Glyph: mod.Glyph };
  } catch {
    return null;
  }
}

function glyphName(glyph: OpentypeGlyph, firstUnicode: number): string {
  if (glyph.name) return glyph.name;
  return `uni${firstUnicode.toString(16).toUpperCase().padStart(4, "0")}`;
}

/**
 * Subsets the font to only the specified characters and returns it as an OTF (CFF) binary.
 *
 * - Characters for which glyphs do not exist in the font (characters that become.notdef) are not included in the subset.
 * To defer to subsequent fallbacks of font-family on the browser side.
 * - Returns null if there is no target character or if subsetting fails.
 */
export async function subsetFont(
  font: OpentypeFullFont,
  chars: Set<string>,
  familyName: string,
): Promise<Uint8Array | null> {
  const opentype = await tryLoadOpentypeCtors();
  if (!opentype) return null;

  const source = unsafeExternalInteropAssertion<SubsettableFont>(font);
  if (typeof source.charToGlyph !== "function" || !source.glyphs) return null;

  // glyph index -> { glyph, responsible unicode set }.
  // Summarize cases where multiple characters are mapped to the same glyph (e.g. ligatureless merging).
  const glyphMap = new Map<number, { glyph: OpentypeGlyph; unicodes: Set<number> }>();
  for (const char of chars) {
    const codePoint = char.codePointAt(0);
    if (codePoint === undefined) continue;
    let glyph: OpentypeGlyph | null;
    try {
      glyph = source.charToGlyph(char);
    } catch {
      continue;
    }
    // index 0 (.notdef) is not included in the font -> excluded from the subset
    if (!glyph || !glyph.index) continue;
    const entry = glyphMap.get(glyph.index);
    if (entry) {
      entry.unicodes.add(codePoint);
    } else {
      glyphMap.set(glyph.index, { glyph, unicodes: new Set([codePoint]) });
    }
  }

  if (glyphMap.size === 0) return null;

  try {
    const notdefSource = source.glyphs.get(0);
    const glyphs: unknown[] = [
      new opentype.Glyph({
        name: ".notdef",
        advanceWidth: notdefSource?.advanceWidth ?? source.unitsPerEm / 2,
        path: notdefSource?.path,
      }),
    ];

    for (const { glyph, unicodes } of glyphMap.values()) {
      const unicodeList = [...unicodes].sort((a, b) => a - b);
      glyphs.push(
        new opentype.Glyph({
          name: glyphName(glyph, unicodeList[0]),
          unicode: unicodeList[0],
          unicodes: unicodeList,
          advanceWidth: glyph.advanceWidth ?? 0,
          path: glyph.path,
        }),
      );
    }

    const subset = new opentype.Font({
      familyName: familyName || "EmbeddedFont",
      styleName: "Regular",
      unitsPerEm: source.unitsPerEm,
      ascender: source.ascender,
      // opentype.js requires negative values for descender
      descender: source.descender < 0 ? source.descender : -1,
      glyphs,
    });

    return new Uint8Array(subset.toArrayBuffer());
  } catch (e) {
    warn(
      "font.subsetFailed",
      `Failed to subset font "${familyName}": ${e instanceof Error ? e.message : String(e)}`,
    );
    return null;
  }
}
