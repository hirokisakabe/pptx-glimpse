import type { TextMeasurer } from "./text-measurer.js";
import { measureTextWidth as defaultMeasureTextWidth } from "./utils/text-measure.js";

const PX_PER_PT = 96 / 72;
const BOLD_FACTOR = 1.05;

interface OpentypeGlyph {
  advanceWidth?: number;
}

export interface OpentypeFont {
  unitsPerEm: number;
  ascender: number;
  descender: number;
  stringToGlyphs(text: string): OpentypeGlyph[];
}

export class OpentypeTextMeasurer implements TextMeasurer {
  private fonts: Map<string, OpentypeFont>;
  private defaultFont: OpentypeFont | null;

  /**
   * @param fonts - フォント名 → opentype.js Font オブジェクトのマップ
   * @param defaultFont - フォールバックフォント（省略可）
   */
  constructor(fonts: Map<string, OpentypeFont>, defaultFont?: OpentypeFont) {
    this.fonts = fonts;
    this.defaultFont = defaultFont ?? null;
  }

  measureTextWidth(
    text: string,
    fontSizePt: number,
    bold: boolean,
    fontFamily?: string | null,
    fontFamilyEa?: string | null,
  ): number {
    const font = this.resolveFont(fontFamily) ?? this.resolveFont(fontFamilyEa) ?? this.defaultFont;
    if (!font) {
      return defaultMeasureTextWidth(text, fontSizePt, bold, fontFamily, fontFamilyEa);
    }
    const fontSizePx = fontSizePt * PX_PER_PT;
    const scale = fontSizePx / font.unitsPerEm;
    let totalWidth = 0;
    const glyphs = font.stringToGlyphs(text);
    for (const glyph of glyphs) {
      totalWidth += (glyph.advanceWidth ?? font.unitsPerEm * 0.6) * scale;
    }
    if (bold) totalWidth *= BOLD_FACTOR;
    return totalWidth;
  }

  getLineHeightRatio(fontFamily?: string | null, fontFamilyEa?: string | null): number {
    const font = this.resolveFont(fontFamily) ?? this.resolveFont(fontFamilyEa) ?? this.defaultFont;
    if (!font) return 1.2;
    return (font.ascender + Math.abs(font.descender)) / font.unitsPerEm;
  }

  private resolveFont(name: string | null | undefined): OpentypeFont | null {
    if (!name) return null;
    return this.fonts.get(name) ?? null;
  }
}
