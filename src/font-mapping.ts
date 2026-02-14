/**
 * PPTX フォント名 → OSS 代替フォント (Google Fonts) のマッピング。
 * ユーザーが拡張・上書き可能。
 */

/** フォントマッピングテーブルの型 */
export type FontMapping = Record<string, string>;

/** デフォルトのフォントマッピングテーブル */
export const DEFAULT_FONT_MAPPING: Readonly<FontMapping> = {
  // ラテン文字フォント
  Calibri: "Carlito",
  Arial: "Arimo",
  "Times New Roman": "Tinos",
  "Courier New": "Cousine",
  Cambria: "Caladea",

  // 日本語ゴシック系 → Noto Sans JP
  メイリオ: "Noto Sans JP",
  Meiryo: "Noto Sans JP",
  游ゴシック: "Noto Sans JP",
  "Yu Gothic": "Noto Sans JP",
  "MS ゴシック": "Noto Sans JP",
  "MS Gothic": "Noto Sans JP",
  "MS Pゴシック": "Noto Sans JP",
  "MS PGothic": "Noto Sans JP",

  // 日本語明朝系 → Noto Serif JP
  "MS 明朝": "Noto Serif JP",
  "MS Mincho": "Noto Serif JP",
  "MS P明朝": "Noto Serif JP",
  "MS PMincho": "Noto Serif JP",
  游明朝: "Noto Serif JP",
  "Yu Mincho": "Noto Serif JP",
};

/**
 * デフォルトマッピングとユーザーマッピングをマージしたテーブルを生成する。
 * ユーザー指定が優先される。
 */
export function createFontMapping(userMapping?: FontMapping): FontMapping {
  if (!userMapping) return { ...DEFAULT_FONT_MAPPING };
  return { ...DEFAULT_FONT_MAPPING, ...userMapping };
}

/**
 * マッピングテーブルから OSS 代替フォント名を取得する。
 * 大文字小文字を区別せずにルックアップする。
 */
export function getMappedFont(
  fontFamily: string | null | undefined,
  mapping: FontMapping,
): string | null {
  if (!fontFamily) return null;

  const direct = mapping[fontFamily];
  if (direct !== undefined) return direct;

  // 大文字小文字を無視したフォールバック
  const lower = fontFamily.toLowerCase();
  for (const key of Object.keys(mapping)) {
    if (key.toLowerCase() === lower) {
      return mapping[key];
    }
  }

  return null;
}
