/**
 * Keep the theme's script-based font (Japan) at the module level.
 * Used as a fallback when rendering CJK text.
 */

let jpanMajorFont: string | null = null;
let jpanMinorFont: string | null = null;

export function setScriptFonts(majorJpan: string | null, minorJpan: string | null): void {
  jpanMajorFont = majorJpan;
  jpanMinorFont = minorJpan;
}

export function resetScriptFonts(): void {
  jpanMajorFont = null;
  jpanMinorFont = null;
}

/**
 * Returns the Japan font for fallback for CJK text.
 * If major/minor distinction is not necessary, major takes precedence.
 */
export function getJpanFallbackFont(): string | null {
  return jpanMajorFont ?? jpanMinorFont;
}
