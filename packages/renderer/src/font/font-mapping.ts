/**
 * Mapping from PPTX font family names to replacement font family names.
 *
 * Keys are font names found in PPTX files. Values are font names that should be
 * present in the rendering environment, commonly open-source alternatives to
 * proprietary Microsoft Office fonts.
 */
export type FontMapping = Record<string, string>;

/**
 * Default replacement mapping for common Microsoft Office fonts.
 *
 * The table maps fonts such as Calibri, Arial, Meiryo, and MS Gothic to
 * open-source alternatives used during text measurement, SVG path generation,
 * and font lookup.
 */
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
 * Create a font mapping table by merging defaults with user overrides.
 *
 * @param userMapping Custom PPTX font name to replacement font name entries.
 * User-specified entries take precedence over `DEFAULT_FONT_MAPPING`.
 * @returns A new mutable mapping object.
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
 * Look up the replacement font for a PPTX font family.
 *
 * Matching is case-insensitive and normalizes full-width alphanumeric
 * characters used by some Japanese Office font names.
 *
 * @param fontFamily PPTX font family name to resolve.
 * @param mapping Mapping table, usually from `createFontMapping`.
 * @returns The replacement font name, or `null` when no mapping exists.
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
