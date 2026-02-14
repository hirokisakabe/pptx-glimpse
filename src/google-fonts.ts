/**
 * Google Fonts からフォントを自動取得するユーティリティ。
 * ブラウザ環境で resvg-wasm にフォントバッファを渡すために使用する。
 */
import type { UsedFonts } from "./font-collector.js";
import type { FontMapping } from "./font-mapping.js";
import { createFontMapping, getMappedFont } from "./font-mapping.js";

/** fetchGoogleFonts のオプション */
export interface FetchGoogleFontsOptions {
  /** カスタムフォントマッピング（デフォルトマッピングにマージされる） */
  fontMapping?: FontMapping;
  /** カスタム fetch 関数（テスト用・プロキシ対応） */
  fetch?: typeof globalThis.fetch;
}

const GOOGLE_FONTS_CSS_API = "https://fonts.googleapis.com/css2";

/**
 * UsedFonts のフォント名にマッピングを適用し、Google Fonts で取得すべきフォント名一覧を返す。
 * 重複除去・ソート済み。
 */
export function resolveGoogleFontNames(usedFonts: UsedFonts, fontMapping?: FontMapping): string[] {
  const mapping = createFontMapping(fontMapping);
  const googleFonts = new Set<string>();

  for (const font of usedFonts.fonts) {
    const mapped = getMappedFont(font, mapping);
    if (mapped) {
      googleFonts.add(mapped);
    }
  }

  return [...googleFonts].sort();
}

/**
 * Google Fonts CSS レスポンスから @font-face の url(...) と font-family を抽出する。
 */
export function parseFontUrlsFromCss(css: string): Array<{ name: string; url: string }> {
  const results: Array<{ name: string; url: string }> = [];
  const faceRegex = /@font-face\s*\{([^}]+)\}/g;
  let faceMatch;

  while ((faceMatch = faceRegex.exec(css)) !== null) {
    const block = faceMatch[1];

    // font-family を抽出
    const familyMatch = block.match(/font-family:\s*'([^']+)'/);
    if (!familyMatch) continue;
    const name = familyMatch[1];

    // url(...) を抽出
    const urlMatch = block.match(/url\(([^)]+)\)/);
    if (!urlMatch) continue;
    const url = urlMatch[1];

    results.push({ name, url });
  }

  return results;
}

/**
 * PPTX の使用フォント情報から Google Fonts のフォントバッファを取得する。
 * resvg-wasm の fontBuffers オプションに渡せる形式で返す。
 *
 * Google Fonts に存在しないフォントはスキップされる。
 */
export async function fetchGoogleFonts(
  usedFonts: UsedFonts,
  options?: FetchGoogleFontsOptions,
): Promise<Array<{ name: string; data: Uint8Array }>> {
  const fetchFn = options?.fetch ?? globalThis.fetch;
  const fontNames = resolveGoogleFontNames(usedFonts, options?.fontMapping);

  if (fontNames.length === 0) return [];

  // Google Fonts CSS API から CSS を取得
  const css = await fetchGoogleFontsCss(fontNames, fetchFn);
  if (!css) return [];

  // CSS からフォント URL を抽出
  const fontEntries = parseFontUrlsFromCss(css);
  if (fontEntries.length === 0) return [];

  // フォントファイルを並列取得
  const results = await Promise.all(
    fontEntries.map(async ({ name, url }) => {
      try {
        const res = await fetchFn(url);
        if (!res.ok) return null;
        const buffer = await res.arrayBuffer();
        return { name, data: new Uint8Array(buffer) };
      } catch {
        return null;
      }
    }),
  );

  const filtered: Array<{ name: string; data: Uint8Array }> = [];
  for (const r of results) {
    if (r !== null) filtered.push(r);
  }
  return filtered;
}

/**
 * Google Fonts CSS API からフォント CSS を取得する。
 * 未知の User-Agent を送ることで TTF 形式のレスポンスを得る。
 */
async function fetchGoogleFontsCss(
  fontNames: string[],
  fetchFn: typeof globalThis.fetch,
): Promise<string | null> {
  const url = new URL(GOOGLE_FONTS_CSS_API);
  for (const name of fontNames) {
    url.searchParams.append("family", name);
  }

  try {
    const res = await fetchFn(url.toString(), {
      headers: { "User-Agent": "pptx-glimpse" },
    });
    if (!res.ok) return null;
    return await res.text();
  } catch {
    return null;
  }
}
