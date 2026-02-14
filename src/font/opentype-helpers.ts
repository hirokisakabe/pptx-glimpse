/**
 * opentype.js を使ってフォントを読み込み OpentypeTextMeasurer を構築するヘルパー。
 */
import { readFile } from "node:fs/promises";
import type { OpentypeFont } from "./opentype-text-measurer.js";
import { OpentypeTextMeasurer } from "./opentype-text-measurer.js";
import type { FontMapping } from "./font-mapping.js";
import { createFontMapping } from "./font-mapping.js";
import type { OpentypeFullFont, TextPathFontResolver } from "./text-path-context.js";
import { DefaultTextPathFontResolver } from "./text-path-context.js";
import { collectFontFilePaths } from "./system-font-loader.js";

/** フォントバッファの入力形式 */
export interface FontBuffer {
  name?: string;
  data: ArrayBuffer | Uint8Array;
}

interface OpentypeFontWithNames extends OpentypeFont {
  names: {
    fontFamily?: Record<string, string>;
  };
}

/**
 * opentype.js を動的 import でロードする。
 * opentype.js がインストールされていない場合は null を返す。
 */
async function tryLoadOpentype(): Promise<{
  parse: (buffer: ArrayBuffer) => OpentypeFontWithNames;
} | null> {
  try {
    // Use a variable to prevent bundlers from statically resolving this import
    const specifier = "opentype.js";
    // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
    const mod: { parse: (buffer: ArrayBuffer) => OpentypeFontWithNames } = await import(
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
 * ArrayBuffer | Uint8Array → ArrayBuffer に変換する。
 * Uint8Array の場合は slice で独立した ArrayBuffer を取得する。
 */
function toArrayBuffer(data: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (data instanceof ArrayBuffer) return data;
  return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
}

/**
 * フォントバッファ配列から OpentypeTextMeasurer を構築する。
 *
 * 内部で opentype.js を動的 import してフォントをパースする。
 * opentype.js が利用不可な場合は null を返す。
 */
export async function createOpentypeTextMeasurerFromBuffers(
  fontBuffers: FontBuffer[],
  fontMapping?: FontMapping,
): Promise<OpentypeTextMeasurer | null> {
  const setup = await createOpentypeSetupFromBuffers(fontBuffers, fontMapping);
  return setup?.measurer ?? null;
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
      const arrayBuffer = toArrayBuffer(buffer.data);
      const font = opentype.parse(arrayBuffer);

      if (!firstMeasurerFont) firstMeasurerFont = font;
      if (!firstResolverFont) firstResolverFont = font as unknown as OpentypeFullFont;

      if (buffer.name) {
        registerFont(buffer.name, font, reverseMap, measurerFonts, resolverFonts);
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

function registerFont(
  name: string,
  font: OpentypeFontWithNames,
  reverseMap: Map<string, string[]>,
  measurerFonts: Map<string, OpentypeFont>,
  resolverFonts: Map<string, OpentypeFullFont>,
): void {
  const fullFont = font as unknown as OpentypeFullFont;
  if (!measurerFonts.has(name)) {
    measurerFonts.set(name, font);
    resolverFonts.set(name, fullFont);
  }

  // 逆引きで PPTX フォント名も登録
  const pptxNames = reverseMap.get(name);
  if (pptxNames) {
    for (const pptxName of pptxNames) {
      if (!measurerFonts.has(pptxName)) {
        measurerFonts.set(pptxName, font);
        resolverFonts.set(pptxName, fullFont);
      }
    }
  }
}

/**
 * システムフォント + 追加ディレクトリから OpentypeTextMeasurer と TextPathFontResolver を構築する。
 *
 * 1. collectFontFilePaths() でフォントファイルパスを収集
 * 2. 各ファイルを readFile + opentype.parse でパース
 * 3. フォント名をキーとしてマップに登録（逆引きマッピング含む）
 */
export async function createOpentypeSetupFromSystem(
  additionalFontDirs?: string[],
  fontMapping?: FontMapping,
): Promise<OpentypeSetup | null> {
  const opentype = await tryLoadOpentype();
  if (!opentype) return null;

  const fontFilePaths = collectFontFilePaths(additionalFontDirs);
  if (fontFilePaths.length === 0) return null;

  const mapping = createFontMapping(fontMapping);
  const reverseMap = buildReverseMapping(mapping);
  const measurerFonts = new Map<string, OpentypeFont>();
  const resolverFonts = new Map<string, OpentypeFullFont>();
  let firstMeasurerFont: OpentypeFont | null = null;
  let firstResolverFont: OpentypeFullFont | null = null;

  for (const filePath of fontFilePaths) {
    try {
      const data = await readFile(filePath);
      const arrayBuffer = toArrayBuffer(data);
      const font = opentype.parse(arrayBuffer);

      if (!firstMeasurerFont) firstMeasurerFont = font;
      if (!firstResolverFont) firstResolverFont = font as unknown as OpentypeFullFont;

      // names.fontFamily からフォント名を取得して登録
      const fontFamily = font.names.fontFamily;
      if (fontFamily) {
        for (const name of Object.values(fontFamily)) {
          registerFont(name, font, reverseMap, measurerFonts, resolverFonts);
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
