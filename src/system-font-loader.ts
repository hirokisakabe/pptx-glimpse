import fs from "node:fs";
import path from "node:path";
import os from "node:os";

const FONT_EXTENSIONS = new Set([".ttf", ".otf"]);

let cachedPaths: string[] | null = null;
let cachedAdditionalDirs: string[] | null = null;

function getSystemFontDirs(): string[] {
  const p = os.platform();
  switch (p) {
    case "linux":
      return ["/usr/share/fonts", "/usr/local/share/fonts"];
    case "darwin":
      return ["/System/Library/Fonts", "/Library/Fonts", path.join(os.homedir(), "Library/Fonts")];
    case "win32":
      return ["C:\\Windows\\Fonts"];
    default:
      return [];
  }
}

function walk(dir: string, result: string[]): void {
  if (!fs.existsSync(dir)) return;
  try {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(fullPath, result);
      } else if (FONT_EXTENSIONS.has(path.extname(entry.name).toLowerCase())) {
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
 * 結果はモジュールレベルでキャッシュされ、同一引数での再呼び出しは即座に返る。
 */
export function collectFontFilePaths(additionalDirs?: string[]): string[] {
  const dirs = additionalDirs ?? [];
  const dirsKey = dirs.join("\0");
  const cachedKey = cachedAdditionalDirs?.join("\0") ?? null;

  if (cachedPaths !== null && dirsKey === cachedKey) {
    return cachedPaths;
  }

  const allDirs = [...getSystemFontDirs(), ...dirs];
  const result: string[] = [];
  for (const dir of allDirs) {
    walk(dir, result);
  }

  cachedPaths = result;
  cachedAdditionalDirs = dirs;
  return result;
}
