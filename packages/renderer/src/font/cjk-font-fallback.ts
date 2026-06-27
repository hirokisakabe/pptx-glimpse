/**
 * CJK フォントの二次フォールバックチェーン。
 * マッピング先フォント (e.g., Noto Sans JP) がシステムに存在しない場合に
 * OS プリインストールの CJK フォントにフォールバックする。
 */
import { platform } from "node:os";

type CjkFallbackMap = Readonly<Record<string, readonly string[]>>;

const MACOS_FALLBACKS: CjkFallbackMap = {
  // ゴシック系
  "Noto Sans JP": ["Hiragino Sans", "Hiragino Kaku Gothic ProN"],
  "Noto Sans CJK JP": ["Hiragino Sans", "Hiragino Kaku Gothic ProN"],
  // 明朝系
  "Noto Serif CJK JP": ["Hiragino Mincho ProN"],
};

const WINDOWS_FALLBACKS: CjkFallbackMap = {
  // ゴシック系
  "Noto Sans JP": ["Yu Gothic", "Meiryo", "MS Gothic"],
  "Noto Sans CJK JP": ["Yu Gothic", "Meiryo", "MS Gothic"],
  // 明朝系
  "Noto Serif CJK JP": ["Yu Mincho", "MS Mincho"],
};

const EMPTY: readonly string[] = [];

let cachedFallbacks: CjkFallbackMap | null = null;

function getFallbackMap(): CjkFallbackMap {
  if (cachedFallbacks) return cachedFallbacks;

  const os = platform();
  switch (os) {
    case "darwin":
      cachedFallbacks = MACOS_FALLBACKS;
      break;
    case "win32":
      cachedFallbacks = WINDOWS_FALLBACKS;
      break;
    default:
      cachedFallbacks = {};
  }
  return cachedFallbacks;
}

/**
 * マッピング先フォント名に対する OS 固有の CJK フォールバックチェーンを返す。
 * フォールバックが定義されていない場合は空配列を返す。
 */
export function getCjkFallbackFonts(mappedFontName: string): readonly string[] {
  return getFallbackMap()[mappedFontName] ?? EMPTY;
}

/** テスト用: キャッシュをクリアする */
export function _resetCjkFallbackCache(): void {
  cachedFallbacks = null;
}
