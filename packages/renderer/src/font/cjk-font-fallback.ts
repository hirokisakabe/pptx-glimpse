/**
 * Secondary fallback chain for CJK fonts.
 * When the mapped font (e.g., Noto Sans JP) does not exist on the system.
 * falls back to OS-preinstalled CJK fonts.
 */

type CjkFallbackMap = Readonly<Record<string, readonly string[]>>;

const MACOS_FALLBACKS: CjkFallbackMap = {
  // Gothic family
  "Noto Sans JP": ["Hiragino Sans", "Hiragino Kaku Gothic ProN"],
  "Noto Sans CJK JP": ["Hiragino Sans", "Hiragino Kaku Gothic ProN"],
  // Mincho family
  "Noto Serif CJK JP": ["Hiragino Mincho ProN"],
};

const WINDOWS_FALLBACKS: CjkFallbackMap = {
  // Gothic family
  "Noto Sans JP": ["Yu Gothic", "Meiryo", "MS Gothic"],
  "Noto Sans CJK JP": ["Yu Gothic", "Meiryo", "MS Gothic"],
  // Mincho family
  "Noto Serif CJK JP": ["Yu Mincho", "MS Mincho"],
};

const EMPTY: readonly string[] = [];

let cachedFallbacks: CjkFallbackMap | null = null;
let platformOverrideForTest: string | null = null;

function getRuntimePlatform(): string {
  if (platformOverrideForTest !== null) return platformOverrideForTest;
  return typeof process !== "undefined" ? process.platform : "";
}

function getFallbackMap(): CjkFallbackMap {
  if (cachedFallbacks) return cachedFallbacks;

  const os = getRuntimePlatform();
  if (os === "darwin") {
    cachedFallbacks = MACOS_FALLBACKS;
  } else if (os === "win32") {
    cachedFallbacks = WINDOWS_FALLBACKS;
  } else {
    cachedFallbacks = {};
  }
  return cachedFallbacks;
}

/**
 * Returns the OS-specific CJK fallback chain for the mapped font name.
 * Returns an empty array if no fallback is defined.
 */
export function getCjkFallbackFonts(mappedFontName: string): readonly string[] {
  return getFallbackMap()[mappedFontName] ?? EMPTY;
}

/** Test-only: clear cache */
export function _resetCjkFallbackCache(): void {
  cachedFallbacks = null;
  platformOverrideForTest = null;
}

/** Test-only: override platform without importing node:os. */
export function _setCjkFallbackPlatformForTest(platform: string): void {
  cachedFallbacks = null;
  platformOverrideForTest = platform;
}
