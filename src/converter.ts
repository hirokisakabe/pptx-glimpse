import { readFileSync, statSync } from "node:fs";

import type { FontMapping } from "./font/font-mapping.js";
import { createFontMapping } from "./font/font-mapping.js";
import { resetFontMapping, setFontMapping } from "./font/font-mapping-context.js";
import { createOpentypeSetupFromSystem } from "./font/opentype-helpers.js";
import { resetScriptFonts, setScriptFonts } from "./font/script-font-context.js";
import { collectFontFilePaths } from "./font/system-font-loader.js";
import { resetTextMeasurer, setTextMeasurer } from "./font/text-measurer.js";
import { resetTextPathFontResolver, setTextPathFontResolver } from "./font/text-path-context.js";
import type { ShapeElement, SlideElement } from "./model/shape.js";
import { clearXmlCache, enableXmlCache } from "./parser/xml-parser.js";
import { svgToPng } from "./png/png-converter.js";
import { parsePptxData, parseSlideWithLayout } from "./pptx-data-parser.js";
import { renderSlideToSvg } from "./renderer/svg-renderer.js";
import { DEFAULT_OUTPUT_WIDTH } from "./utils/constants.js";
import type { LogLevel } from "./warning-logger.js";
import { flushWarnings, initWarningLogger, warn } from "./warning-logger.js";

export interface ConvertOptions {
  /** 変換対象のスライド番号 (1始まり)。未指定で全スライド */
  slides?: number[];
  /** 出力画像の幅 (ピクセル)。デフォルト: 960 */
  width?: number;
  /** 出力画像の高さ (ピクセル)。widthと同時指定時はwidthが優先 */
  height?: number;
  /** 警告ログレベル。デフォルト: "off" */
  logLevel?: LogLevel;
  /** 追加のフォントディレクトリパス。システムフォントに加えて検索する */
  fontDirs?: string[];
  /** PPTX フォント名 → OSS 代替フォントのカスタムマッピング。デフォルトマッピングにマージされる */
  fontMapping?: FontMapping;
  /** true のとき OS のシステムフォントをスキャンせず fontDirs のみを使用する */
  skipSystemFonts?: boolean;
}

export interface SlideSvg {
  slideNumber: number;
  svg: string;
}

export interface SlideImage {
  slideNumber: number;
  png: Buffer;
  width: number;
  height: number;
}

export async function convertPptxToSvg(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideSvg[]> {
  const setup = await createOpentypeSetupFromSystem(
    options?.fontDirs,
    options?.fontMapping,
    options?.skipSystemFonts,
  );
  if (setup) {
    setTextMeasurer(setup.measurer);
    setTextPathFontResolver(setup.fontResolver);
  }
  setFontMapping(createFontMapping(options?.fontMapping));
  enableXmlCache();
  try {
    initWarningLogger(options?.logLevel ?? "off");

    const data = parsePptxData(input);
    setScriptFonts(data.theme.fontScheme.majorFontJpan, data.theme.fontScheme.minorFontJpan);

    // Filter slides if specified
    const targetSlides = options?.slides
      ? data.slidePaths.filter((s) => options.slides!.includes(s.slideNumber))
      : data.slidePaths;

    if (data.slidePaths.length === 0) {
      warn("presentation.noSlides", "No slides found in the PPTX file");
    }

    // Parse and render each slide
    const results: SlideSvg[] = [];
    for (const { slideNumber, path } of targetSlides) {
      const parsed = parseSlideWithLayout(slideNumber, path, data);
      if (!parsed) continue;

      const { slide, layoutElements, layoutShowMasterSp, masterElements } = parsed;

      // Merge shapes: master (back) → layout → slide (front)
      const effectiveMasterElements =
        slide.showMasterSp && layoutShowMasterSp ? masterElements : [];
      slide.elements = mergeElements(effectiveMasterElements, layoutElements, slide.elements);

      const svg = renderSlideToSvg(slide, data.presInfo.slideSize);
      results.push({ slideNumber, svg });
    }

    flushWarnings();

    return results;
  } finally {
    clearXmlCache();
    resetTextMeasurer();
    resetTextPathFontResolver();
    resetFontMapping();
    resetScriptFonts();
  }
}

/**
 * resvg に渡すフォントバッファのキャッシュ。
 * collectFontFilePaths と同じキャッシュキーで管理する。
 */
let cachedFontBuffers: Uint8Array[] | null = null;
let cachedFontBuffersKey: string | null = null;

/**
 * TTF/OTF フォントファイルを読み込んでバッファとして返す。
 * resvg-wasm は fontFiles (ファイルパス) を解釈できないため、
 * fontBuffers (生バイト) として渡す必要がある。
 * 合計サイズが MAX_TOTAL_FONT_BUFFER_BYTES を超えた時点で読み込みを打ち切る。
 */
const MAX_TOTAL_FONT_BUFFER_BYTES = 100 * 1024 * 1024; // 100MB

function loadFontBuffers(fontDirs?: string[], skipSystemFonts?: boolean): Uint8Array[] {
  const key = `${(fontDirs ?? []).join("\0")}\n${skipSystemFonts ?? false}`;
  if (cachedFontBuffers !== null && cachedFontBuffersKey === key) {
    return cachedFontBuffers;
  }

  const allPaths = collectFontFilePaths(fontDirs, skipSystemFonts);
  // TTC は resvg-wasm では不安定なため TTF/OTF のみを対象とする
  const ttfOtfPaths = allPaths.filter((p) => {
    const lower = p.toLowerCase();
    return lower.endsWith(".ttf") || lower.endsWith(".otf");
  });

  // ファイルサイズ昇順に並べ、小さいフォントから優先的に読み込む
  const pathsWithSize: { path: string; size: number }[] = [];
  for (const p of ttfOtfPaths) {
    try {
      pathsWithSize.push({ path: p, size: statSync(p).size });
    } catch {
      // 読み取れないファイルはスキップ
    }
  }
  pathsWithSize.sort((a, b) => a.size - b.size);

  const buffers: Uint8Array[] = [];
  let totalSize = 0;
  for (const { path, size } of pathsWithSize) {
    if (totalSize + size > MAX_TOTAL_FONT_BUFFER_BYTES) break;
    try {
      buffers.push(new Uint8Array(readFileSync(path)));
      totalSize += size;
    } catch {
      // 読み取り失敗はスキップ
    }
  }

  cachedFontBuffers = buffers;
  cachedFontBuffersKey = key;
  return buffers;
}

export async function convertPptxToPng(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideImage[]> {
  const svgResults = await convertPptxToSvg(input, options);

  const width = options?.width ?? DEFAULT_OUTPUT_WIDTH;
  const height = options?.height;

  // resvg に渡すフォントバッファを収集する（チャートの <text> 要素を描画するため）
  const fontBuffers = loadFontBuffers(options?.fontDirs, options?.skipSystemFonts);

  const results: SlideImage[] = [];
  for (const { slideNumber, svg } of svgResults) {
    const pngResult = await svgToPng(svg, { width, height, fontBuffers });
    results.push({
      slideNumber,
      png: pngResult.png,
      width: pngResult.width,
      height: pngResult.height,
    });
  }

  return results;
}

function mergeElements(
  masterElements: SlideElement[],
  layoutElements: SlideElement[],
  slideElements: SlideElement[],
): SlideElement[] {
  // Placeholder shapes in master and layout are templates (position/style definitions).
  // Their text content should never appear on actual slides.
  // Only non-placeholder shapes (decorative elements, logos, etc.) are shown.
  const filterTemplatePlaceholders = (elements: SlideElement[]) =>
    elements.filter((el) => {
      if (el.type !== "shape") return true;
      return !el.placeholderType;
    });

  // Placeholder shapes on the slide itself are templates that the user has not
  // filled in when their TextBody contains no run text. PowerPoint hides them
  // entirely (the "Click to add title" prompt lives in the layout and is never
  // copied into the slide), so we drop them too.
  const filterEmptySlidePlaceholders = (elements: SlideElement[]) =>
    elements.filter((el) => !(el.type === "shape" && isEmptyPlaceholder(el)));

  const filteredMaster = filterTemplatePlaceholders(masterElements);
  const filteredLayout = filterTemplatePlaceholders(layoutElements);
  const filteredSlide = filterEmptySlidePlaceholders(slideElements);

  return [...filteredMaster, ...filteredLayout, ...filteredSlide];
}

function isEmptyPlaceholder(shape: ShapeElement): boolean {
  if (!shape.placeholderType) return false;
  const paragraphs = shape.textBody?.paragraphs;
  if (!paragraphs || paragraphs.length === 0) return true;
  return !paragraphs.some((p) => p.runs.some((r) => r.text.length > 0));
}
