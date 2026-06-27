/**
 * opentype.js を使ってフォントを使用文字のみにサブセット化する。
 * ネイティブ <text> 出力モードの @font-face 埋め込み用。
 */

import { unsafeTypeAssertion } from "../unsafe-type-assertion.js";
import { warn } from "../warning-logger.js";
import type { OpentypeFullFont } from "./text-path-context.js";

interface OpentypeGlyph {
  index: number;
  name?: string | null;
  unicode?: number;
  unicodes?: number[];
  advanceWidth?: number;
  path: unknown;
}

interface SubsettableFont {
  unitsPerEm: number;
  ascender: number;
  descender: number;
  charToGlyph(char: string): OpentypeGlyph | null;
  glyphs: { get(index: number): OpentypeGlyph };
}

interface OpentypeCtors {
  Font: new (options: Record<string, unknown>) => { toArrayBuffer(): ArrayBuffer };
  Glyph: new (options: Record<string, unknown>) => unknown;
}

/**
 * opentype.js のコンストラクタを動的 import でロードする。
 * opentype.js がインストールされていない場合は null を返す。
 */
async function tryLoadOpentypeCtors(): Promise<OpentypeCtors | null> {
  try {
    // Use a variable to prevent bundlers from statically resolving this import
    const specifier = "opentype.js";
    // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
    const mod: OpentypeCtors = await import(/* @vite-ignore */ specifier);
    return { Font: mod.Font, Glyph: mod.Glyph };
  } catch {
    return null;
  }
}

function glyphName(glyph: OpentypeGlyph, firstUnicode: number): string {
  if (glyph.name) return glyph.name;
  return `uni${firstUnicode.toString(16).toUpperCase().padStart(4, "0")}`;
}

/**
 * フォントを指定文字のみにサブセット化し、OTF (CFF) バイナリとして返す。
 *
 * - フォントにグリフが存在しない文字 (.notdef になる文字) はサブセットに含めない。
 *   ブラウザ側で font-family の後続フォールバックに委ねるため。
 * - 対象文字が 1 つも無い場合やサブセット化に失敗した場合は null を返す。
 */
export async function subsetFont(
  font: OpentypeFullFont,
  chars: Set<string>,
  familyName: string,
): Promise<Uint8Array | null> {
  const opentype = await tryLoadOpentypeCtors();
  if (!opentype) return null;

  const source = unsafeTypeAssertion<SubsettableFont>(font);
  if (typeof source.charToGlyph !== "function" || !source.glyphs) return null;

  // グリフインデックス → { グリフ, 担当ユニコード集合 }。
  // 同一グリフに複数文字がマップされるケース (例: 合字なし統合) をまとめる。
  const glyphMap = new Map<number, { glyph: OpentypeGlyph; unicodes: Set<number> }>();
  for (const char of chars) {
    const codePoint = char.codePointAt(0);
    if (codePoint === undefined) continue;
    let glyph: OpentypeGlyph | null;
    try {
      glyph = source.charToGlyph(char);
    } catch {
      continue;
    }
    // index 0 (.notdef) はフォント未収録 → サブセットから除外
    if (!glyph || !glyph.index) continue;
    const entry = glyphMap.get(glyph.index);
    if (entry) {
      entry.unicodes.add(codePoint);
    } else {
      glyphMap.set(glyph.index, { glyph, unicodes: new Set([codePoint]) });
    }
  }

  if (glyphMap.size === 0) return null;

  try {
    const notdefSource = source.glyphs.get(0);
    const glyphs: unknown[] = [
      new opentype.Glyph({
        name: ".notdef",
        advanceWidth: notdefSource?.advanceWidth ?? source.unitsPerEm / 2,
        path: notdefSource?.path,
      }),
    ];

    for (const { glyph, unicodes } of glyphMap.values()) {
      const unicodeList = [...unicodes].sort((a, b) => a - b);
      glyphs.push(
        new opentype.Glyph({
          name: glyphName(glyph, unicodeList[0]),
          unicode: unicodeList[0],
          unicodes: unicodeList,
          advanceWidth: glyph.advanceWidth ?? 0,
          path: glyph.path,
        }),
      );
    }

    const subset = new opentype.Font({
      familyName: familyName || "EmbeddedFont",
      styleName: "Regular",
      unitsPerEm: source.unitsPerEm,
      ascender: source.ascender,
      // opentype.js は descender に負値を要求する
      descender: source.descender < 0 ? source.descender : -1,
      glyphs,
    });

    return new Uint8Array(subset.toArrayBuffer());
  } catch (e) {
    warn(
      "font.subsetFailed",
      `Failed to subset font "${familyName}": ${e instanceof Error ? e.message : String(e)}`,
    );
    return null;
  }
}
