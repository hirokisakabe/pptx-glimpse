/**
 * レンダリング中のフォントマッピングテーブルをモジュールレベルで保持する。
 * text-measurer.ts の setTextMeasurer/resetTextMeasurer パターンと同じ。
 */
import type { FontMapping } from "./font-mapping.js";
import { DEFAULT_FONT_MAPPING, getMappedFont } from "./font-mapping.js";

let currentMapping: FontMapping = { ...DEFAULT_FONT_MAPPING };

export function setFontMapping(mapping: FontMapping): void {
  currentMapping = mapping;
}

export function getFontMapping(): FontMapping {
  return currentMapping;
}

export function resetFontMapping(): void {
  currentMapping = { ...DEFAULT_FONT_MAPPING };
}

/**
 * 現在のマッピングテーブルから OSS 代替フォント名を取得するショートカット。
 */
export function getCurrentMappedFont(fontFamily: string | null | undefined): string | null {
  return getMappedFont(fontFamily, currentMapping);
}
