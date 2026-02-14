import { renderSlideToSvg } from "./renderer/svg-renderer.js";
import { svgToPng } from "./png/png-converter.js";
import { DEFAULT_OUTPUT_WIDTH } from "./utils/constants.js";
import type { SlideElement } from "./model/shape.js";
import type { LogLevel } from "./warning-logger.js";
import { initWarningLogger, flushWarnings } from "./warning-logger.js";
import { setTextMeasurer, resetTextMeasurer } from "./font/text-measurer.js";
import { parsePptxData, parseSlideWithLayout } from "./pptx-data-parser.js";
import type { FontMapping } from "./font/font-mapping.js";
import { createFontMapping } from "./font/font-mapping.js";
import { setFontMapping, resetFontMapping } from "./font/font-mapping-context.js";
import { createOpentypeSetupFromSystem } from "./font/opentype-helpers.js";
import { setTextPathFontResolver, resetTextPathFontResolver } from "./font/text-path-context.js";

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
  const setup = await createOpentypeSetupFromSystem(options?.fontDirs, options?.fontMapping);
  if (setup) {
    setTextMeasurer(setup.measurer);
    setTextPathFontResolver(setup.fontResolver);
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
    resetTextPathFontResolver();
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
    const pngResult = await svgToPng(svg, { width, height });
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
