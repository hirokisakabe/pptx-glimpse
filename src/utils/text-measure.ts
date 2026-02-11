import { type FontMetrics, getFontMetrics } from "../data/font-metrics.js";

type CharCategory = "narrow" | "normal" | "wide";

const WIDTH_RATIO: Record<CharCategory, number> = {
  narrow: 0.3,
  normal: 0.6,
  wide: 1.0,
};

const BOLD_FACTOR = 1.05;
const PX_PER_PT = 96 / 72;

function isCjkCodePoint(codePoint: number): boolean {
  return (
    (codePoint >= 0x3000 && codePoint <= 0x9fff) ||
    (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
    (codePoint >= 0xff01 && codePoint <= 0xff60) ||
    (codePoint >= 0x20000 && codePoint <= 0x2a6df)
  );
}

function categorizeChar(codePoint: number): CharCategory {
  // CJK 統合漢字、ひらがな、カタカナ、CJK 記号、全角英数
  if (isCjkCodePoint(codePoint)) {
    return "wide";
  }

  // 狭い文字
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
 * テキストの推定幅を計算する (ピクセル単位)。
 * fontFamily が指定され、対応するメトリクスが存在する場合はメトリクスベースで計算する。
 * それ以外はヒューリスティック (文字カテゴリ別の固定比率) にフォールバックする。
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
    if (metrics) {
      totalWidth += measureCharMetrics(char, codePoint, baseSizePx, metrics);
    } else {
      totalWidth += measureCharHeuristic(codePoint, baseSizePx);
    }
  }

  if (bold) {
    totalWidth *= BOLD_FACTOR;
  }

  return totalWidth;
}
