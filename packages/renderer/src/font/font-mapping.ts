/**
 * PPTX font name -> OSS alternative font (Google Fonts) mapping.
 * Users can extend or override it.
 */

/** Font mapping table type */
export type FontMapping = Record<string, string>;

/** Default font mapping table */
export const DEFAULT_FONT_MAPPING: Readonly<FontMapping> = {
  // Latin fonts
  Calibri: "Carlito",
  "Calibri Light": "Carlito",
  Arial: "Arimo",
  "Times New Roman": "Tinos",
  "Courier New": "Cousine",
  Cambria: "Caladea",

  // Japanese Gothic fonts -> Noto Sans JP
  // Use "Noto Sans JP" instead of "Noto Sans CJK JP".
  // NotoSansCJK TTC extracts only the first font, so it may not always be possible to obtain the JP variant.
  // The font name matches the standalone NotoSansJP.ttf downloaded in the Docker environment.
  メイリオ: "Noto Sans JP",
  Meiryo: "Noto Sans JP",
  游ゴシック: "Noto Sans JP",
  "Yu Gothic": "Noto Sans JP",
  "MS ゴシック": "Noto Sans JP",
  "MS Gothic": "Noto Sans JP",
  "MS Pゴシック": "Noto Sans JP",
  "MS PGothic": "Noto Sans JP",

  // Japanese Mincho fonts -> Noto Serif CJK JP
  "MS 明朝": "Noto Serif CJK JP",
  "MS Mincho": "Noto Serif CJK JP",
  "MS P明朝": "Noto Serif CJK JP",
  "MS PMincho": "Noto Serif CJK JP",
  游明朝: "Noto Serif CJK JP",
  "Yu Mincho": "Noto Serif CJK JP",
};

/**
 * Generate a table that merges default mapping and user mapping.
 * User-specified entries take precedence.
 */
export function createFontMapping(userMapping?: FontMapping): FontMapping {
  if (!userMapping) return { ...DEFAULT_FONT_MAPPING };
  return { ...DEFAULT_FONT_MAPPING, ...userMapping };
}

/**
 * Normalizes full-width alphanumerics and symbols to half-width.
 * PPTX themes may use full-width spellings such as "\uFF2D\uFF33 \uFF30\u30B4\u30B7\u30C3\u30AF".
 */
function normalizeFullWidth(s: string): string {
  return s
    .replace(/[\uff01-\uff5e]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0))
    .replace(/\u3000/g, " ");
}

/**
 * Get the OSS alternative font name from the mapping table.
 * Looks up without case sensitivity.
 */
export function getMappedFont(
  fontFamily: string | null | undefined,
  mapping: FontMapping,
): string | null {
  if (!fontFamily) return null;

  const direct = mapping[fontFamily];
  if (direct !== undefined) return direct;

  const normalized = normalizeFullWidth(fontFamily);

  // Exact match after normalization
  if (normalized !== fontFamily) {
    const directNormalized = mapping[normalized];
    if (directNormalized !== undefined) return directNormalized;
  }

  // Case-insensitive fallback
  const lower = normalized.toLowerCase();
  for (const key of Object.keys(mapping)) {
    if (normalizeFullWidth(key).toLowerCase() === lower) {
      return mapping[key];
    }
  }

  return null;
}
