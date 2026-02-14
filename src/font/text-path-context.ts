/**
 * テキスト→パス変換用のフォントリゾルバーコンテキスト。
 * opentype.js の getPath() メソッドを持つ Font オブジェクトへのアクセスを提供する。
 */

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
  ): OpentypeFullFont | null {
    if (fontFamily) {
      const font = this.fonts.get(fontFamily);
      if (font) return font;
    }
    if (fontFamilyEa) {
      const font = this.fonts.get(fontFamilyEa);
      if (font) return font;
    }
    return this.defaultFont;
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
