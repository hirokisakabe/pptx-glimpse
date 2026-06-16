/**
 * テーマのスクリプトベースフォント（Jpan）をモジュールレベルで保持する。
 * CJK テキストレンダリング時のフォールバックとして使用される。
 */

let jpanMajorFont: string | null = null;
let jpanMinorFont: string | null = null;

export function setScriptFonts(majorJpan: string | null, minorJpan: string | null): void {
  jpanMajorFont = majorJpan;
  jpanMinorFont = minorJpan;
}

export function resetScriptFonts(): void {
  jpanMajorFont = null;
  jpanMinorFont = null;
}

/**
 * CJK テキストのフォールバック用 Jpan フォントを返す。
 * major/minor の区別が不要な場合は major を優先する。
 */
export function getJpanFallbackFont(): string | null {
  return jpanMajorFont ?? jpanMinorFont;
}
