/**
 * Font resolver context for text-to-path conversion.
 * Provides access to a Font object with the opentype.js getPath() method.
 */

import type { WarningLogger } from "../warning-logger.js";
import { warn } from "../warning-logger.js";
import { getCjkFallbackFonts } from "./cjk-font-fallback.js";
import type { FontMapping } from "./font-mapping.js";
import { getMappedFont } from "./font-mapping.js";
import { getCurrentMappedFont } from "./font-mapping-context.js";

export interface OpentypePath {
  toPathData(decimalPlaces?: number): string;
}

export interface OpentypeFullFont {
  unitsPerEm: number;
  ascender: number;
  descender: number;
  getPath(text: string, x: number, y: number, fontSize: number): OpentypePath;
  getAdvanceWidth(text: string, fontSize: number): number;
}

export interface TextPathFontResolver {
  resolveFont(
    fontFamily: string | null | undefined,
    fontFamilyEa: string | null | undefined,
    jpanFallback?: string | null,
    context?: TextPathFontResolverContext,
  ): OpentypeFullFont | null;
}

export interface TextPathFontResolverContext {
  readonly fontMapping?: FontMapping;
  readonly warningLogger?: WarningLogger;
  readonly fontWarningCache?: Set<string>;
}

export class DefaultTextPathFontResolver implements TextPathFontResolver {
  private fonts: Map<string, OpentypeFullFont>;
  private defaultFont: OpentypeFullFont | null;
  private warnedFonts = new Set<string>();

  constructor(fonts: Map<string, OpentypeFullFont>, defaultFont?: OpentypeFullFont) {
    this.fonts = fonts;
    this.defaultFont = defaultFont ?? null;
  }

  resolveFont(
    fontFamily: string | null | undefined,
    fontFamilyEa: string | null | undefined,
    jpanFallback?: string | null,
    context?: TextPathFontResolverContext,
  ): OpentypeFullFont | null {
    if (fontFamily) {
      const font = this.findFont(fontFamily, context);
      if (font) return font;
    }
    if (fontFamilyEa) {
      const font = this.findFont(fontFamilyEa, context);
      if (font) return font;
    }
    if (jpanFallback) {
      const font = this.findFont(jpanFallback, context);
      if (font) return font;
    }

    // Font-not-found warning
    for (const name of [fontFamily, fontFamilyEa, jpanFallback]) {
      if (name) {
        const logger = context?.warningLogger;
        if (logger) {
          if (context.fontWarningCache?.has(name)) continue;
          context.fontWarningCache?.add(name);
          logger.warn("font.notFound", `Font not found: "${name}"`);
          continue;
        }
        if (!this.warnedFonts.has(name)) {
          this.warnedFonts.add(name);
          warn("font.notFound", `Font not found: "${name}"`);
        }
      }
    }

    return this.defaultFont;
  }

  private findFont(name: string, context?: TextPathFontResolverContext): OpentypeFullFont | null {
    const direct = this.fonts.get(name);
    if (direct) return direct;

    // Try OSS replacement names from font mapping
    const mapped =
      context?.fontMapping !== undefined
        ? getMappedFont(name, context.fontMapping)
        : getCurrentMappedFont(name);
    if (mapped) {
      const mappedFont = this.fonts.get(mapped);
      if (mappedFont) return mappedFont;

      // CJK fallback chain
      for (const fallback of getCjkFallbackFonts(mapped)) {
        const fallbackFont = this.fonts.get(fallback);
        if (fallbackFont) return fallbackFont;
      }
    }

    return null;
  }
}

let currentResolver: TextPathFontResolver | null = null;

export function setTextPathFontResolver(resolver: TextPathFontResolver): void {
  currentResolver = resolver;
}

export function getTextPathFontResolver(): TextPathFontResolver | null {
  return currentResolver;
}

export function resetTextPathFontResolver(): void {
  currentResolver = null;
}
