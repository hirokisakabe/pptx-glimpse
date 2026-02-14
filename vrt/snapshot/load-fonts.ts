/**
 * Docker コンテナ内のシステムフォントを fontBuffers として読み込むユーティリティ。
 * update-snapshots.ts と regression.test.ts の両方で使用する。
 */
import { readFileSync, readdirSync, existsSync, statSync } from "fs";
import { join, extname } from "path";
import type { FontMapping } from "../../src/font-mapping.js";

const FONT_SEARCH_DIRS = ["/usr/share/fonts/truetype", "/usr/share/fonts/opentype"];

/** フォントファイルを再帰的に検索 (.ttf, .otf のみ、Regular ウェイトのみ) */
function findRegularFontFiles(dir: string): string[] {
  if (!existsSync(dir)) return [];
  const results: string[] = [];
  for (const entry of readdirSync(dir)) {
    const fullPath = join(dir, entry);
    const stat = statSync(fullPath);
    if (stat.isDirectory()) {
      results.push(...findRegularFontFiles(fullPath));
    } else if (/\.(ttf|otf)$/i.test(entry) && /Regular/i.test(entry)) {
      results.push(fullPath);
    }
  }
  return results;
}

/** ファイル名からフォントファミリー名を抽出 (例: "LiberationSans-Regular.ttf" → "LiberationSans") */
function extractFontFamily(fileName: string): string {
  return fileName.replace(extname(fileName), "").replace(/-Regular$/i, "");
}

/** Docker コンテナ内のシステムフォント (Regular ウェイト) を fontBuffers として読み込む */
export function loadSystemFontBuffers(): Array<{ name: string; data: Uint8Array }> {
  const fontFiles = FONT_SEARCH_DIRS.flatMap(findRegularFontFiles);
  const buffers: Array<{ name: string; data: Uint8Array }> = [];

  for (const filePath of fontFiles.sort()) {
    const fileName = filePath.split("/").pop()!;
    const name = extractFontFamily(fileName);
    try {
      buffers.push({ name, data: readFileSync(filePath) });
    } catch {
      // Skip unreadable files
    }
  }

  return buffers;
}

/** VRT 用フォントマッピング (PPTX フォント名 → Docker 内フォント名) */
export const VRT_FONT_MAPPING: FontMapping = {
  // ラテン文字
  Calibri: "LiberationSans",
  Arial: "LiberationSans",
  "Times New Roman": "LiberationSerif",
  "Courier New": "LiberationMono",
  Cambria: "LiberationSerif",

  // 日本語ゴシック系
  メイリオ: "NotoSansCJKjp",
  Meiryo: "NotoSansCJKjp",
  游ゴシック: "NotoSansCJKjp",
  "Yu Gothic": "NotoSansCJKjp",
  "MS ゴシック": "NotoSansCJKjp",
  "MS Gothic": "NotoSansCJKjp",
  "MS Pゴシック": "NotoSansCJKjp",
  "MS PGothic": "NotoSansCJKjp",

  // 日本語明朝系
  "MS 明朝": "NotoSansCJKjp",
  "MS Mincho": "NotoSansCJKjp",
  "MS P明朝": "NotoSansCJKjp",
  "MS PMincho": "NotoSansCJKjp",
  游明朝: "NotoSansCJKjp",
  "Yu Mincho": "NotoSansCJKjp",
};
