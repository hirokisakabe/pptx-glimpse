import { type FontMetrics, getFontMetrics } from "../data/font-metrics.js";

type CharCategory = "narrow" | "normal" | "wide";

// Heuristic width ratio (vs. fontSize) in the absence of font metrics.
// Measurement-based approximations.
const WIDTH_RATIO: Record<CharCategory, number> = {
  narrow: 0.3, // i, l, 1, punctuation, etc.
  normal: 0.6, // average width of latin letters
  wide: 1.0, // CJK characters (full-width)
};

const BOLD_FACTOR = 1.05;
const PX_PER_PT = 96 / 72;
// Line-height fallback for fonts without OpenType metrics (CSS default equivalent)
const DEFAULT_LINE_HEIGHT_RATIO = 1.2;
// ascender ratio fallback for fonts without OpenType metrics
const DEFAULT_ASCENDER_RATIO = 1.0;

/**
 * Natural font line-height ratio.
 * Calculated as (ascender + |descender|) / unitsPerEm.
 * Returns fallback value (1.2) if metric is not found.
 */
export function getLineHeightRatio(
  fontFamily?: string | null,
  fontFamilyEa?: string | null,
): number {
  const metrics = getFontMetrics(fontFamily) ?? getFontMetrics(fontFamilyEa);
  if (!metrics) return DEFAULT_LINE_HEIGHT_RATIO;
  return (metrics.ascender + Math.abs(metrics.descender)) / metrics.unitsPerEm;
}

/**
 * Font ascender ratio.
 * Calculated as ascender / unitsPerEm.
 * SVG's `<text y="...">` specifies the baseline position, so
 * the first-line baseline offset uses this value rather than the line-height ratio.
 * Returns a fallback value (1.0) if the metric is not found.
 */
export function getAscenderRatio(fontFamily?: string | null, fontFamilyEa?: string | null): number {
  const metrics = getFontMetrics(fontFamily) ?? getFontMetrics(fontFamilyEa);
  if (!metrics) return DEFAULT_ASCENDER_RATIO;
  return metrics.ascender / metrics.unitsPerEm;
}

// CJK character determination (code point range based on Unicode Standard)
export function isCjkCodePoint(codePoint: number): boolean {
  return (
    (codePoint >= 0x3000 && codePoint <= 0x9fff) || // CJK symbols, hiragana, katakana, integrated kanji
    (codePoint >= 0xf900 && codePoint <= 0xfaff) || // CJK compatible kanji
    (codePoint >= 0xff01 && codePoint <= 0xff60) || // Full-width alphanumeric characters/symbols
    (codePoint >= 0x20000 && codePoint <= 0x2a6df) // CJK Integrated Kanji Expansion B
  );
}

function categorizeChar(codePoint: number): CharCategory {
  // CJK unified ideographs, hiragana, katakana, CJK symbols, and full-width alphanumerics
  if (isCjkCodePoint(codePoint)) {
    return "wide";
  }

  // Narrow characters
  if (
    codePoint === 0x20 || // space
    codePoint === 0x21 || // !
    codePoint === 0x2c || // ,
    codePoint === 0x2e || // .
    codePoint === 0x3a || // :
    codePoint === 0x3b || // ;
    codePoint === 0x69 || // i
    codePoint === 0x6a || // j
    codePoint === 0x6c || // l
    codePoint === 0x31 || // 1
    codePoint === 0x7c || // |
    codePoint === 0x27 || // '
    codePoint === 0x28 || // (
    codePoint === 0x29 || // )
    codePoint === 0x5b || // [
    codePoint === 0x5d || // ]
    codePoint === 0x7b || // {
    codePoint === 0x7d // }
  ) {
    return "narrow";
  }

  return "normal";
}

function measureCharHeuristic(codePoint: number, baseSizePx: number): number {
  return baseSizePx * WIDTH_RATIO[categorizeChar(codePoint)];
}

function measureCharMetrics(
  char: string,
  codePoint: number,
  baseSizePx: number,
  metrics: FontMetrics,
): number {
  const charWidth = metrics.widths[char];
  if (charWidth !== undefined) {
    return (charWidth / metrics.unitsPerEm) * baseSizePx;
  }
  if (isCjkCodePoint(codePoint)) {
    return (metrics.cjkWidth / metrics.unitsPerEm) * baseSizePx;
  }
  return (metrics.defaultWidth / metrics.unitsPerEm) * baseSizePx;
}

/**
 * Calculates estimated text width in pixels.
 * If fontFamily is specified and the corresponding metrics exist, calculate based on metrics.
 * Otherwise, it falls back to a heuristic (fixed proportions by character category).
 */
export function measureTextWidth(
  text: string,
  fontSizePt: number,
  bold: boolean,
  fontFamily?: string | null,
  fontFamilyEa?: string | null,
): number {
  const baseSizePx = fontSizePt * PX_PER_PT;
  const latinMetrics = getFontMetrics(fontFamily);
  const eaMetrics = getFontMetrics(fontFamilyEa);
  let totalWidth = 0;

  for (const char of text) {
    const codePoint = char.codePointAt(0)!;
    const isEa = isCjkCodePoint(codePoint);
    const metrics = isEa && eaMetrics ? eaMetrics : latinMetrics;
    let charWidth: number;
    if (metrics) {
      charWidth = measureCharMetrics(char, codePoint, baseSizePx, metrics);
    } else {
      charWidth = measureCharHeuristic(codePoint, baseSizePx);
    }
    if (bold && !isEa) {
      charWidth *= BOLD_FACTOR;
    }
    totalWidth += charWidth;
  }

  return totalWidth;
}
