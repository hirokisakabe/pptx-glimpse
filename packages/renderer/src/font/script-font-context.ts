/**
 * Internal note.
 * Internal note.
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
 * Internal note.
 * Internal note.
 */
export function getJpanFallbackFont(): string | null {
  return jpanMajorFont ?? jpanMinorFont;
}
