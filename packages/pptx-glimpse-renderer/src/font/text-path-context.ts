/**
 * テキスト→パス変換用のフォントリゾルバーコンテキスト。
 * opentype.js の getPath() メソッドを持つ Font オブジェクトへのアクセスを提供する。
 */

import { getCurrentMappedFont } from "./font-mapping-context.js";

export interface OpentypePath {
  toPathData(decimalPlaces?: number): string;
}

export interface OpentypeFullFont {
  unitsPerEm: number;
  ascender: number;
  descender: number;
  getPath(text: string, x: number, y: number, fontSize: number): OpentypePath;
}

export interface TextPathFontResolver {
  resolveFont(
    fontFamily: string | null | undefined,
    fontFamilyEa: string | null | undefined,
    jpanFallback?: string | null,
  ): OpentypeFullFont | null;
}

export class DefaultTextPathFontResolver implements TextPathFontResolver {
  private fonts: Map<string, OpentypeFullFont>;
  private defaultFont: OpentypeFullFont | null;

  constructor(fonts: Map<string, OpentypeFullFont>, defaultFont?: OpentypeFullFont) {
    this.fonts = fonts;
    this.defaultFont = defaultFont ?? null;
  }

  resolveFont(
    fontFamily: string | null | undefined,
    fontFamilyEa: string | null | undefined,
    jpanFallback?: string | null,
  ): OpentypeFullFont | null {
    if (fontFamily) {
      const font = this.findFont(fontFamily);
      if (font) return font;
    }
    if (fontFamilyEa) {
      const font = this.findFont(fontFamilyEa);
      if (font) return font;
    }
    if (jpanFallback) {
      const font = this.findFont(jpanFallback);
      if (font) return font;
    }
    return this.defaultFont;
  }

  private findFont(name: string): OpentypeFullFont | null {
    const direct = this.fonts.get(name);
    if (direct) return direct;

    // フォントマッピングで OSS 代替名を試行
    const mapped = getCurrentMappedFont(name);
    if (mapped) {
      const mappedFont = this.fonts.get(mapped);
      if (mappedFont) return mappedFont;
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
