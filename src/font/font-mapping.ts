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
/**
 * 全角英数字・記号を半角に正規化する。
 * PPTX テーマでは「ＭＳ Ｐゴシック」のように全角が使われることがある。
 */
function normalizeFullWidth(s: string): string {
  return s
    .replace(/[\uff01-\uff5e]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0))
    .replace(/\u3000/g, " ");
}

export function getMappedFont(
  fontFamily: string | null | undefined,
  mapping: FontMapping,
): string | null {
  if (!fontFamily) return null;

  const direct = mapping[fontFamily];
  if (direct !== undefined) return direct;

  const normalized = normalizeFullWidth(fontFamily);

  // 正規化後の完全一致
  if (normalized !== fontFamily) {
    const directNormalized = mapping[normalized];
    if (directNormalized !== undefined) return directNormalized;
  }

  // 大文字小文字を無視したフォールバック
  const lower = normalized.toLowerCase();
  for (const key of Object.keys(mapping)) {
    if (normalizeFullWidth(key).toLowerCase() === lower) {
      return mapping[key];
    }
  }

  return null;
}
