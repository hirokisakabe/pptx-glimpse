import {
  isCjkCodePoint,
  measureTextWidth as defaultMeasureTextWidth,
} from "../utils/text-measure.js";
import { warn } from "../warning-logger.js";
import { getCjkFallbackFonts } from "./cjk-font-fallback.js";
import type { FontMapping } from "./font-mapping.js";
import { getMappedFont } from "./font-mapping.js";
import { getCurrentMappedFont } from "./font-mapping-context.js";
import type { TextMeasurementContext, TextMeasurer } from "./text-measurer.js";

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
   * @param fonts - font name -> opentype.js Map of Font objects
   * @param defaultFont - fallback font (optional)
   */
  constructor(
    fonts: Map<string, OpentypeFont>,
    defaultFont?: OpentypeFont,
    private readonly fontMapping?: FontMapping,
  ) {
    this.fonts = fonts;
    this.defaultFont = defaultFont ?? null;
  }

  measureTextWidth(
    text: string,
    fontSizePt: number,
    bold: boolean,
    fontFamily?: string | null,
    fontFamilyEa?: string | null,
    context?: TextMeasurementContext,
  ): number {
    const latinFont = this.resolveFont(fontFamily, context);
    const eaFont = this.resolveFont(fontFamilyEa, context);
    const fallbackFont = latinFont ?? eaFont ?? this.defaultFont;
    if (!fallbackFont) {
      return defaultMeasureTextWidth(text, fontSizePt, bold, fontFamily, fontFamilyEa);
    }
    const fontSizePx = fontSizePt * PX_PER_PT;
    // CJK characters prefer East Asian fonts, Latin characters prefer Latin fonts.
    const latinFontResolved = latinFont ?? fallbackFont;
    const eaFontResolved = eaFont ?? fallbackFont;
    const boldLatinFont = bold ? this.resolveBoldFont(fontFamily, context) : null;

    // Cache each unique character to reduce stringToGlyphs calls
    // Full-text bulk calls cannot be used because the number of glyphs changes with the GSUB ligature.
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

  getLineHeightRatio(
    fontFamily?: string | null,
    fontFamilyEa?: string | null,
    context?: TextMeasurementContext,
  ): number {
    const font =
      this.resolveFont(fontFamily, context) ??
      this.resolveFont(fontFamilyEa, context) ??
      this.defaultFont;
    if (!font) return 1.2;
    return (font.ascender + Math.abs(font.descender)) / font.unitsPerEm;
  }

  getAscenderRatio(
    fontFamily?: string | null,
    fontFamilyEa?: string | null,
    context?: TextMeasurementContext,
  ): number {
    const font =
      this.resolveFont(fontFamily, context) ??
      this.resolveFont(fontFamilyEa, context) ??
      this.defaultFont;
    if (!font) return 1.0;
    return font.ascender / font.unitsPerEm;
  }

  private resolveBoldFont(
    name: string | null | undefined,
    context?: TextMeasurementContext,
  ): OpentypeFont | null {
    if (!name) return null;
    // Look for Bold variants by both the original name and the OSS alternate name after font mapping
    const bases = [name];
    const mappedBase = this.getMappedFont(name, context);
    if (mappedBase && mappedBase !== name) bases.push(mappedBase);
    for (const base of bases) {
      for (const boldName of [`${base} Bold`, `${base}-Bold`]) {
        const direct = this.fonts.get(boldName);
        if (direct) return direct;
        const mapped = this.getMappedFont(boldName, context);
        if (mapped) {
          const mappedFont = this.fonts.get(mapped);
          if (mappedFont) return mappedFont;
        }
      }
    }
    return null;
  }

  private resolveFont(
    name: string | null | undefined,
    context?: TextMeasurementContext,
  ): OpentypeFont | null {
    if (!name) return null;
    const direct = this.fonts.get(name);
    if (direct) return direct;

    // Try OSS replacement names from font mapping
    const mapped = this.getMappedFont(name, context);
    if (mapped) {
      const mappedFont = this.fonts.get(mapped);
      if (mappedFont) return mappedFont;

      // CJK fallback chain
      for (const fallback of getCjkFallbackFonts(mapped)) {
        const fallbackFont = this.fonts.get(fallback);
        if (fallbackFont) return fallbackFont;
      }
    }

    // Font-not-found warning
    if (context?.warningLogger) {
      if (context.fontWarningCache?.has(name)) return null;
      context.fontWarningCache?.add(name);
      context.warningLogger.warn("font.notFound", `Font not found: "${name}"`);
    } else if (!this.warnedFonts.has(name)) {
      this.warnedFonts.add(name);
      warn("font.notFound", `Font not found: "${name}"`);
    }

    return null;
  }

  private getMappedFont(
    name: string | null | undefined,
    context?: TextMeasurementContext,
  ): string | null {
    return context?.fontMapping !== undefined
      ? getMappedFont(name, context.fontMapping)
      : this.fontMapping !== undefined
        ? getMappedFont(name, this.fontMapping)
        : getCurrentMappedFont(name);
  }
}
