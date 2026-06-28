import {
  getAscenderRatio as defaultGetAscenderRatio,
  getLineHeightRatio as defaultGetLineHeightRatio,
  measureTextWidth as defaultMeasureTextWidth,
} from "../utils/text-measure.js";

/**
 * Interface for measuring text width and line height.
 * By default, measurement uses static font metrics.
 * Users can pass their own implementation to ConvertOptions.textMeasurer
 * Internal note.
 */
export interface TextMeasurer {
  /**
   * Internal note.
   * @param text - text to measure
   * @param fontSizePt - font size (points)
   * @param bold - whether the text is bold
   * @param fontFamily - font family name for Latin text
   * @param fontFamilyEa - font family name for East Asian text
   */
  measureTextWidth(
    text: string,
    fontSizePt: number,
    bold: boolean,
    fontFamily?: string | null,
    fontFamilyEa?: string | null,
  ): number;

  /**
   * Internal note.
   * Internal note.
   * @param fontFamily - font family name for Latin text
   * @param fontFamilyEa - font family name for East Asian text
   */
  getLineHeightRatio(fontFamily?: string | null, fontFamilyEa?: string | null): number;

  /**
   * Font ascender ratio (ascender / unitsPerEm) .
   * Internal note.
   * Internal note.
   * @param fontFamily - font family name for Latin text
   * @param fontFamilyEa - font family name for East Asian text
   */
  getAscenderRatio(fontFamily?: string | null, fontFamilyEa?: string | null): number;
}

export class DefaultTextMeasurer implements TextMeasurer {
  measureTextWidth(
    text: string,
    fontSizePt: number,
    bold: boolean,
    fontFamily?: string | null,
    fontFamilyEa?: string | null,
  ): number {
    return defaultMeasureTextWidth(text, fontSizePt, bold, fontFamily, fontFamilyEa);
  }

  getLineHeightRatio(fontFamily?: string | null, fontFamilyEa?: string | null): number {
    return defaultGetLineHeightRatio(fontFamily, fontFamilyEa);
  }

  getAscenderRatio(fontFamily?: string | null, fontFamilyEa?: string | null): number {
    return defaultGetAscenderRatio(fontFamily, fontFamilyEa);
  }
}

let currentMeasurer: TextMeasurer = new DefaultTextMeasurer();

export function setTextMeasurer(measurer: TextMeasurer): void {
  currentMeasurer = measurer;
}

export function getTextMeasurer(): TextMeasurer {
  return currentMeasurer;
}

export function resetTextMeasurer(): void {
  currentMeasurer = new DefaultTextMeasurer();
}
