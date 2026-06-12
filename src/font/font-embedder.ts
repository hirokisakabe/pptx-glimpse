/**
 * フォント使用状況から @font-face 定義入りの <style> 要素を構築する。
 * ネイティブ <text> 出力モードで SVG にサブセット化フォントを埋め込むために使う。
 */

import { Buffer } from "node:buffer";

import { subsetFont } from "./font-subsetter.js";
import type { FontUsage } from "./font-usage-collector.js";
import { getJpanFallbackFont } from "./script-font-context.js";
import type { TextPathFontResolver } from "./text-path-context.js";

/**
 * CSS 文字列リテラル用にフォント名をエスケープする。
 * <style> の中身は XML テキストノードでもあるため、XML 特殊文字は CSS 16進エスケープで表現する。
 */
function escapeCssFamilyName(name: string): string {
  return name
    .replace(/[\\"]/g, (c) => `\\${c}`)
    .replace(/[<>&]/g, (c) => `\\${c.codePointAt(0)!.toString(16)} `);
}

/**
 * 収集したフォント使用状況をサブセット化し、@font-face 定義の <style> 要素を返す。
 * 埋め込めるフォントが 1 つも無い場合は空文字列を返す。
 */
export async function buildFontFaceStyle(
  usages: Map<string, FontUsage>,
  fontResolver: TextPathFontResolver,
): Promise<string> {
  const faces: string[] = [];
  const jpanFallback = getJpanFallbackFont();

  for (const [familyName, usage] of usages) {
    const font = fontResolver.resolveFont(
      usage.fonts[0],
      usage.fonts[1] ?? null,
      usage.fonts[2] ?? jpanFallback,
    );
    if (!font) continue;

    const buffer = await subsetFont(font, usage.chars, familyName);
    if (!buffer) continue;

    const base64 = Buffer.from(buffer).toString("base64");
    faces.push(
      `@font-face{font-family:"${escapeCssFamilyName(familyName)}";src:url(data:font/otf;base64,${base64}) format("opentype");}`,
    );
  }

  if (faces.length === 0) return "";
  return `<style type="text/css">${faces.join("")}</style>`;
}
