type CharCategory = "narrow" | "normal" | "wide";

const WIDTH_RATIO: Record<CharCategory, number> = {
  narrow: 0.3,
  normal: 0.6,
  wide: 1.0,
};

const BOLD_FACTOR = 1.05;
const PX_PER_PT = 96 / 72;

function categorizeChar(codePoint: number): CharCategory {
  // CJK 統合漢字、ひらがな、カタカナ、CJK 記号、全角英数
  if (
    (codePoint >= 0x3000 && codePoint <= 0x9fff) ||
    (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
    (codePoint >= 0xff01 && codePoint <= 0xff60) ||
    (codePoint >= 0x20000 && codePoint <= 0x2a6df)
  ) {
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

/**
 * テキストの推定幅を計算する (ピクセル単位)
 */
export function measureTextWidth(text: string, fontSizePt: number, bold: boolean): number {
  const baseSizePx = fontSizePt * PX_PER_PT;
  let totalWidth = 0;

  for (const char of text) {
    const codePoint = char.codePointAt(0)!;
    const category = categorizeChar(codePoint);
    totalWidth += baseSizePx * WIDTH_RATIO[category];
  }

  if (bold) {
    totalWidth *= BOLD_FACTOR;
  }

  return totalWidth;
}
