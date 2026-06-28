import { existsSync, readdirSync } from "node:fs";
import { homedir, platform } from "node:os";
import { extname, join } from "node:path";

const FONT_EXTENSIONS = new Set([".ttf", ".otf"]);

/**
 * Known patterns for CJK TTC fonts.
 * Internal note.
 * so only fonts needed for CJK text are loaded selectively.
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
  if (os === "linux") return ["/usr/share/fonts", "/usr/local/share/fonts"];
  if (os === "darwin") {
    return ["/System/Library/Fonts", "/Library/Fonts", join(homedir(), "Library/Fonts")];
  }
  if (os === "win32") return ["C:\\Windows\\Fonts"];
  return [];
}

function isCjkTtc(name: string): boolean {
  // macOS (APFS) returns filenames as NFD (decomposed form) 、
  // Normalize to NFC before pattern matching.
  // Internal note.
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
 * From OS system font directories plus additional directories,
 * .ttf / .otf collects .ttf / .otf file paths.
 *
 * Internal note.
 * additionalDirs only.
 *
 * Results are cached at module level, and repeated calls with the same arguments return immediately.
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
