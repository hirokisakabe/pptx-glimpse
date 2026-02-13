import { renderSlideToSvg } from "./renderer/svg-renderer.js";
import { svgToPng } from "./png/png-converter.js";
import { DEFAULT_OUTPUT_WIDTH } from "./utils/constants.js";
import type { SlideElement } from "./model/shape.js";
import type { LogLevel } from "./warning-logger.js";
import { initWarningLogger, flushWarnings } from "./warning-logger.js";
import type { TextMeasurer } from "./text-measurer.js";
import { setTextMeasurer, resetTextMeasurer } from "./text-measurer.js";
import { parsePptxData, parseSlideWithLayout } from "./pptx-data-parser.js";
import type { FontMapping } from "./font-mapping.js";
import { createFontMapping } from "./font-mapping.js";
import { setFontMapping, resetFontMapping } from "./font-mapping-context.js";

export interface FontOptions {
  /** resvg-wasm に渡すフォントファイルパス (Node.js 向け) */
  fontFiles?: string[];
  /** resvg-wasm に渡すフォントディレクトリパス (Node.js 向け) */
  fontDirs?: string[];
  /** resvg-wasm に渡すフォントバッファ (ブラウザ/Node.js 両対応) */
  fontBuffers?: Array<{ name?: string; data: ArrayBuffer | Uint8Array }>;
  /** resvg-wasm でシステムフォントを読み込むか (デフォルト: false) */
  loadSystemFonts?: boolean;
  /** resvg-wasm のデフォルトフォントファミリー */
  defaultFontFamily?: string;
  /** resvg-wasm の sans-serif フォントファミリー */
  sansSerifFamily?: string;
  /** resvg-wasm の serif フォントファミリー */
  serifFamily?: string;
}

export interface ConvertOptions {
  /** 変換対象のスライド番号 (1始まり)。未指定で全スライド */
  slides?: number[];
  /** 出力画像の幅 (ピクセル)。デフォルト: 960 */
  width?: number;
  /** 出力画像の高さ (ピクセル)。widthと同時指定時はwidthが優先 */
  height?: number;
  /** 警告ログレベル。デフォルト: "off" */
  logLevel?: LogLevel;
  /** テキスト計測のカスタム実装 */
  textMeasurer?: TextMeasurer;
  /** PNG 変換時のフォント設定 (resvg-wasm オプション) */
  fonts?: FontOptions;
  /** PPTX フォント名 → OSS 代替フォントのカスタムマッピング。デフォルトマッピングにマージされる */
  fontMapping?: FontMapping;
}

export interface SlideSvg {
  slideNumber: number;
  svg: string;
}

export interface SlideImage {
  slideNumber: number;
  png: Uint8Array;
  width: number;
  height: number;
}

export async function convertPptxToSvg(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideSvg[]> {
  if (options?.textMeasurer) {
    setTextMeasurer(options.textMeasurer);
  }
  setFontMapping(createFontMapping(options?.fontMapping));
  try {
    initWarningLogger(options?.logLevel ?? "off");

    const data = await parsePptxData(input);

    // Filter slides if specified
    const targetSlides = options?.slides
      ? data.slidePaths.filter((s) => options.slides!.includes(s.slideNumber))
      : data.slidePaths;

    // Parse and render each slide
    const results: SlideSvg[] = [];
    for (const { slideNumber, path } of targetSlides) {
      const parsed = parseSlideWithLayout(slideNumber, path, data);
      if (!parsed) continue;

      const { slide, layoutElements, layoutShowMasterSp } = parsed;

      // Merge shapes: master (back) → layout → slide (front)
      const effectiveMasterElements =
        slide.showMasterSp && layoutShowMasterSp ? data.masterElements : [];
      slide.elements = mergeElements(effectiveMasterElements, layoutElements, slide.elements);

      const svg = renderSlideToSvg(slide, data.presInfo.slideSize);
      results.push({ slideNumber, svg });
    }

    flushWarnings();

    return results;
  } finally {
    resetTextMeasurer();
    resetFontMapping();
  }
}

export async function convertPptxToPng(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideImage[]> {
  const svgResults = await convertPptxToSvg(input, options);

  const width = options?.width ?? DEFAULT_OUTPUT_WIDTH;
  const height = options?.height;

  const results: SlideImage[] = [];
  for (const { slideNumber, svg } of svgResults) {
    const pngResult = await svgToPng(svg, { width, height, fonts: options?.fonts });
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
  const filterPlaceholders = (elements: SlideElement[]) =>
    elements.filter((el) => {
      if (el.type !== "shape") return true;
      return !el.placeholderType;
    });

  const filteredMaster = filterPlaceholders(masterElements);
  const filteredLayout = filterPlaceholders(layoutElements);

  return [...filteredMaster, ...filteredLayout, ...slideElements];
}
