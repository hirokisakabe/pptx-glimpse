import {
  measureTextWidth as defaultMeasureTextWidth,
  getLineHeightRatio as defaultGetLineHeightRatio,
  getAscenderRatio as defaultGetAscenderRatio,
} from "../utils/text-measure.js";

/**
 * テキストの幅と行高さを計測するインターフェース。
 * デフォルトでは静的フォントメトリクスによる計測が使われる。
 * ユーザーが独自の実装を ConvertOptions.textMeasurer に渡すことで、
 * Canvas API や opentype.js など任意の計測バックエンドに差し替えられる。
 */
export interface TextMeasurer {
  /**
   * テキストの推定幅をピクセル単位で返す。
   * @param text - 計測対象のテキスト
   * @param fontSizePt - フォントサイズ (ポイント)
   * @param bold - 太字かどうか
   * @param fontFamily - ラテン文字用フォントファミリー名
   * @param fontFamilyEa - 東アジア文字用フォントファミリー名
   */
  measureTextWidth(
    text: string,
    fontSizePt: number,
    bold: boolean,
    fontFamily?: string | null,
    fontFamilyEa?: string | null,
  ): number;

  /**
   * フォントの自然な行高さ比率 (行高さ / フォントサイズ) を返す。
   * メトリクスが不明な場合は 1.2 をフォールバック値として返す。
   * @param fontFamily - ラテン文字用フォントファミリー名
   * @param fontFamilyEa - 東アジア文字用フォントファミリー名
   */
  getLineHeightRatio(fontFamily?: string | null, fontFamilyEa?: string | null): number;

  /**
   * フォントの ascender 比率 (ascender / unitsPerEm) を返す。
   * 1行目のベースラインオフセット計算に使用する。
   * メトリクスが不明な場合は 1.0 をフォールバック値として返す。
   * @param fontFamily - ラテン文字用フォントファミリー名
   * @param fontFamilyEa - 東アジア文字用フォントファミリー名
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
