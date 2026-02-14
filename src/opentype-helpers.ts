/**
 * opentype.js を使ってフォントバッファから OpentypeTextMeasurer を構築するヘルパー。
 */
import type { OpentypeFont } from "./opentype-text-measurer.js";
import { OpentypeTextMeasurer } from "./opentype-text-measurer.js";
import type { FontMapping } from "./font-mapping.js";
import { createFontMapping } from "./font-mapping.js";
import type { OpentypeFullFont, TextPathFontResolver } from "./text-path-context.js";
import { DefaultTextPathFontResolver } from "./text-path-context.js";

/** フォントバッファの入力形式（ConvertOptions.fonts.fontBuffers と互換） */
export interface FontBuffer {
  name?: string;
  data: ArrayBuffer | Uint8Array;
}

/**
 * opentype.js を動的 import でロードする。
 * opentype.js がインストールされていない場合は null を返す。
 */
async function tryLoadOpentype(): Promise<{
  parse: (buffer: ArrayBuffer) => OpentypeFont;
} | null> {
  try {
    // Use a variable to prevent bundlers from statically resolving this optional import
    const specifier = "opentype.js";
    // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
    const mod: { parse: (buffer: ArrayBuffer) => OpentypeFont } = await import(
      /* @vite-ignore */ specifier
    );
    return { parse: mod.parse };
  } catch {
    return null;
  }
}

/**
 * フォントマッピングの逆引きテーブルを構築する。
 * OSS フォント名 → PPTX フォント名[] のマッピング。
 * 例: "Carlito" → ["Calibri"]
 */
function buildReverseMapping(mapping: FontMapping): Map<string, string[]> {
  const reverse = new Map<string, string[]>();
  for (const [pptxName, ossName] of Object.entries(mapping)) {
    const existing = reverse.get(ossName) ?? [];
    existing.push(pptxName);
    reverse.set(ossName, existing);
  }
  return reverse;
}

/**
 * フォントバッファ配列から OpentypeTextMeasurer を構築する。
 *
 * 内部で opentype.js を動的 import してフォントをパースする。
 * opentype.js が利用不可な場合は null を返す。
 *
 * Map には以下のキーで登録する:
 * 1. バッファの name フィールド（例: "Carlito"）
 * 2. フォントマッピングの逆引きで得られる PPTX 名（例: "Calibri"）
 *
 * これにより OpentypeTextMeasurer.resolveFont("Calibri") が
 * Carlito のフォントデータで解決できる。
 */
export async function createOpentypeTextMeasurerFromBuffers(
  fontBuffers: FontBuffer[],
  fontMapping?: FontMapping,
): Promise<OpentypeTextMeasurer | null> {
  if (fontBuffers.length === 0) return null;

  const opentype = await tryLoadOpentype();
  if (!opentype) return null;

  const mapping = createFontMapping(fontMapping);
  const reverseMap = buildReverseMapping(mapping);
  const fonts = new Map<string, OpentypeFont>();
  let firstFont: OpentypeFont | null = null;

  for (const buffer of fontBuffers) {
    try {
      let arrayBuffer: ArrayBuffer;
      if (buffer.data instanceof ArrayBuffer) {
        arrayBuffer = buffer.data;
      } else {
        // Uint8Array → ArrayBuffer: slice で独立した ArrayBuffer を取得
        arrayBuffer = buffer.data.buffer.slice(
          buffer.data.byteOffset,
          buffer.data.byteOffset + buffer.data.byteLength,
        ) as ArrayBuffer;
      }
      const font = opentype.parse(arrayBuffer);

      if (!firstFont) firstFont = font;

      // バッファの name があればそのキーで登録
      if (buffer.name) {
        fonts.set(buffer.name, font);

        // 逆引きで PPTX フォント名も登録
        const pptxNames = reverseMap.get(buffer.name);
        if (pptxNames) {
          for (const pptxName of pptxNames) {
            if (!fonts.has(pptxName)) {
              fonts.set(pptxName, font);
            }
          }
        }
      }
    } catch {
      // パース失敗のフォントはスキップ
    }
  }

  if (fonts.size === 0 && !firstFont) return null;

  return new OpentypeTextMeasurer(fonts, firstFont ?? undefined);
}

export interface OpentypeSetup {
  measurer: OpentypeTextMeasurer;
  fontResolver: TextPathFontResolver;
}

/**
 * フォントバッファ配列から OpentypeTextMeasurer と TextPathFontResolver を同時に構築する。
 *
 * opentype.parse() が返すオブジェクトは OpentypeFont と OpentypeFullFont の両方を満たすため、
 * 同じ Font オブジェクトを measurer と fontResolver の両方に渡す。
 */
export async function createOpentypeSetupFromBuffers(
  fontBuffers: FontBuffer[],
  fontMapping?: FontMapping,
): Promise<OpentypeSetup | null> {
  if (fontBuffers.length === 0) return null;

  const opentype = await tryLoadOpentype();
  if (!opentype) return null;

  const mapping = createFontMapping(fontMapping);
  const reverseMap = buildReverseMapping(mapping);
  const measurerFonts = new Map<string, OpentypeFont>();
  const resolverFonts = new Map<string, OpentypeFullFont>();
  let firstMeasurerFont: OpentypeFont | null = null;
  let firstResolverFont: OpentypeFullFont | null = null;

  for (const buffer of fontBuffers) {
    try {
      let arrayBuffer: ArrayBuffer;
      if (buffer.data instanceof ArrayBuffer) {
        arrayBuffer = buffer.data;
      } else {
        arrayBuffer = buffer.data.buffer.slice(
          buffer.data.byteOffset,
          buffer.data.byteOffset + buffer.data.byteLength,
        ) as ArrayBuffer;
      }
      // opentype.parse() の返却値は OpentypeFont & OpentypeFullFont を満たす
      const font = opentype.parse(arrayBuffer);
      const fullFont = font as unknown as OpentypeFullFont;

      if (!firstMeasurerFont) firstMeasurerFont = font;
      if (!firstResolverFont) firstResolverFont = fullFont;

      if (buffer.name) {
        measurerFonts.set(buffer.name, font);
        resolverFonts.set(buffer.name, fullFont);

        const pptxNames = reverseMap.get(buffer.name);
        if (pptxNames) {
          for (const pptxName of pptxNames) {
            if (!measurerFonts.has(pptxName)) {
              measurerFonts.set(pptxName, font);
              resolverFonts.set(pptxName, fullFont);
            }
          }
        }
      }
    } catch {
      // パース失敗のフォントはスキップ
    }
  }

  if (measurerFonts.size === 0 && !firstMeasurerFont) return null;

  const measurer = new OpentypeTextMeasurer(measurerFonts, firstMeasurerFont ?? undefined);
  const fontResolver = new DefaultTextPathFontResolver(
    resolverFonts,
    firstResolverFont ?? undefined,
  );

  return { measurer, fontResolver };
}
