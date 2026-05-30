import {
  isCjkCodePoint,
  measureTextWidth as defaultMeasureTextWidth,
} from "../utils/text-measure.js";
import { warn } from "../warning-logger.js";
import { getCjkFallbackFonts } from "./cjk-font-fallback.js";
import { getCurrentMappedFont } from "./font-mapping-context.js";
import type { TextMeasurer } from "./text-measurer.js";

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
  private warnedFonts = new Set<string>();

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
    const latinFont = this.resolveFont(fontFamily);
    const eaFont = this.resolveFont(fontFamilyEa);
    const fallbackFont = latinFont ?? eaFont ?? this.defaultFont;
    if (!fallbackFont) {
      return defaultMeasureTextWidth(text, fontSizePt, bold, fontFamily, fontFamilyEa);
    }
    const fontSizePx = fontSizePt * PX_PER_PT;
    // CJK 文字は東アジアフォントを優先、ラテン文字はラテンフォントを優先
    const latinFontResolved = latinFont ?? fallbackFont;
    const eaFontResolved = eaFont ?? fallbackFont;
    const boldLatinFont = bold ? this.resolveBoldFont(fontFamily) : null;

    // ユニーク文字ごとにキャッシュして stringToGlyphs 呼び出しを削減
    // full-text 一括呼び出しは GSUB リガチャでグリフ数が変わるため使用不可
    const latinGlyphCache = new Map<string, OpentypeGlyph | undefined>();
    const eaGlyphCache =
      eaFontResolved !== latinFontResolved
        ? new Map<string, OpentypeGlyph | undefined>()
        : latinGlyphCache;
    const boldGlyphCache = new Map<string, OpentypeGlyph | undefined>();

    let totalWidth = 0;
    for (const char of text) {
      const codePoint = char.codePointAt(0)!;
      const isEa = isCjkCodePoint(codePoint);
      const font = isEa ? eaFontResolved : latinFontResolved;
      const cache = isEa ? eaGlyphCache : latinGlyphCache;
      if (!cache.has(char)) {
        cache.set(char, font.stringToGlyphs(char)[0]);
      }
      const scale = fontSizePx / font.unitsPerEm;
      const glyph = cache.get(char);
      let charWidth = (glyph?.advanceWidth ?? font.unitsPerEm * 0.6) * scale;
      if (bold && !isEa) {
        if (boldLatinFont) {
          if (!boldGlyphCache.has(char)) {
            boldGlyphCache.set(char, boldLatinFont.stringToGlyphs(char)[0]);
          }
          const boldGlyph = boldGlyphCache.get(char);
          const boldScale = fontSizePx / boldLatinFont.unitsPerEm;
          charWidth = (boldGlyph?.advanceWidth ?? boldLatinFont.unitsPerEm * 0.6) * boldScale;
        } else {
          charWidth *= BOLD_FACTOR;
        }
      }
      totalWidth += charWidth;
    }
    return totalWidth;
  }

  getLineHeightRatio(fontFamily?: string | null, fontFamilyEa?: string | null): number {
    const font = this.resolveFont(fontFamily) ?? this.resolveFont(fontFamilyEa) ?? this.defaultFont;
    if (!font) return 1.2;
    return (font.ascender + Math.abs(font.descender)) / font.unitsPerEm;
  }

  getAscenderRatio(fontFamily?: string | null, fontFamilyEa?: string | null): number {
    const font = this.resolveFont(fontFamily) ?? this.resolveFont(fontFamilyEa) ?? this.defaultFont;
    if (!font) return 1.0;
    return font.ascender / font.unitsPerEm;
  }

  private resolveBoldFont(name: string | null | undefined): OpentypeFont | null {
    if (!name) return null;
    // 元の名前とフォントマッピング後の OSS 代替名の両方で Bold バリアントを探す
    const bases = [name];
    const mappedBase = getCurrentMappedFont(name);
    if (mappedBase && mappedBase !== name) bases.push(mappedBase);
    for (const base of bases) {
      for (const boldName of [`${base} Bold`, `${base}-Bold`]) {
        const direct = this.fonts.get(boldName);
        if (direct) return direct;
        const mapped = getCurrentMappedFont(boldName);
        if (mapped) {
          const mappedFont = this.fonts.get(mapped);
          if (mappedFont) return mappedFont;
        }
      }
    }
    return null;
  }

  private resolveFont(name: string | null | undefined): OpentypeFont | null {
    if (!name) return null;
    const direct = this.fonts.get(name);
    if (direct) return direct;

    // フォントマッピングで OSS 代替名を試行
    const mapped = getCurrentMappedFont(name);
    if (mapped) {
      const mappedFont = this.fonts.get(mapped);
      if (mappedFont) return mappedFont;

      // CJK フォールバックチェーン
      for (const fallback of getCjkFallbackFonts(mapped)) {
        const fallbackFont = this.fonts.get(fallback);
        if (fallbackFont) return fallbackFont;
      }
    }

    // フォント未検出の警告
    if (!this.warnedFonts.has(name)) {
      this.warnedFonts.add(name);
      warn("font.notFound", `Font not found: "${name}"`);
    }

    return null;
  }
}
