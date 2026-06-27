import { existsSync, readdirSync } from "node:fs";
import { homedir, platform } from "node:os";
import { extname, join } from "node:path";

const FONT_EXTENSIONS = new Set([".ttf", ".otf"]);

/**
 * CJK TTC フォントの既知パターン。
 * TTC ファイルは大量にあるとメモリを消費するため、
 * CJK テキストに必要なもののみを選択的に読み込む。
 */
const CJK_TTC_PATTERNS = [
  "NotoSansCJK",
  "NotoSerifCJK",
  // macOS
  "Hiragino",
  "ヒラギノ",
  // Windows
  "YuGoth",
  "YuMin",
  "meiryo",
  "msgothic",
  "msmincho",
];

let cachedPaths: string[] | null = null;
let cachedAdditionalDirs: string[] | null = null;
let cachedSkipSystemFonts: boolean | null = null;

function getSystemFontDirs(): string[] {
  const os = platform();
  switch (os) {
    case "linux":
      return ["/usr/share/fonts", "/usr/local/share/fonts"];
    case "darwin":
      return ["/System/Library/Fonts", "/Library/Fonts", join(homedir(), "Library/Fonts")];
    case "win32":
      return ["C:\\Windows\\Fonts"];
    default:
      return [];
  }
}

function isCjkTtc(name: string): boolean {
  // macOS (APFS) はファイル名を NFD (分解形) で返すため、
  // NFC に正規化してからパターンマッチする。
  // 例: "ギ" が "キ" + 濁点(U+3099) に分解されている場合がある。
  const lower = name.normalize("NFC").toLowerCase();
  return lower.endsWith(".ttc") && CJK_TTC_PATTERNS.some((p) => lower.includes(p.toLowerCase()));
}

function walk(dir: string, result: string[]): void {
  if (!existsSync(dir)) return;
  try {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const fullPath = join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(fullPath, result);
      } else if (FONT_EXTENSIONS.has(extname(entry.name).toLowerCase()) || isCjkTtc(entry.name)) {
        result.push(fullPath);
      }
    }
  } catch {
    // Permission errors etc. — skip silently
  }
}

/**
 * OS のシステムフォントディレクトリ + 追加ディレクトリから
 * .ttf / .otf ファイルパスを収集する。
 *
 * skipSystemFonts が true の場合、システムフォントディレクトリをスキャンせず
 * additionalDirs のみを対象とする。
 *
 * 結果はモジュールレベルでキャッシュされ、同一引数での再呼び出しは即座に返る。
 */
export function collectFontFilePaths(additionalDirs?: string[], skipSystemFonts = false): string[] {
  const dirs = additionalDirs ?? [];
  const dirsKey = dirs.join("\0");
  const cachedKey = cachedAdditionalDirs?.join("\0") ?? null;

  if (cachedPaths !== null && dirsKey === cachedKey && cachedSkipSystemFonts === skipSystemFonts) {
    return cachedPaths;
  }

  const allDirs = skipSystemFonts ? dirs : [...getSystemFontDirs(), ...dirs];
  const result: string[] = [];
  for (const dir of allDirs) {
    walk(dir, result);
  }
  result.sort((a, b) => a.localeCompare(b));

  cachedPaths = result;
  cachedAdditionalDirs = dirs;
  cachedSkipSystemFonts = skipSystemFonts;
  return result;
}
