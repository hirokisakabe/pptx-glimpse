import type { PptxSourceModel } from "@pptx-glimpse/document";
import { DEFAULT_OUTPUT_WIDTH } from "@pptx-glimpse/renderer";

import {
  type ConvertOptions,
  convertPptxToSvg as convertPptxToSvgBase,
  renderPptxSourceModelToSvg as renderPptxSourceModelToSvgBase,
  type SupportCoverage,
  type SvgConversionReport,
  type SystemFontSetupLoader,
} from "./svg-converter.js";

export type {
  ConversionDiagnostic,
  ConvertOptions,
  SlideSupportCoverage,
  SlideSvg,
  SupportCoverage,
  SupportCoverageCounts,
  SvgConversionReport,
} from "./svg-converter.js";
export type { PptxSourceModel } from "@pptx-glimpse/document";

/**
 * PNG conversion result for one slide.
 */
export interface SlideImage {
  /**
   * Original slide number in the PPTX file, using 1-based numbering.
   */
  slideNumber: number;
  /**
   * PNG image bytes for the rendered slide.
   */
  png: Uint8Array;
  /**
   * Actual output image width in pixels after rasterization.
   */
  width: number;
  /**
   * Actual output image height in pixels after rasterization.
   */
  height: number;
}

export interface PngConversionReport {
  readonly slides: readonly SlideImage[];
  readonly diagnostics: SvgConversionReport["diagnostics"];
  readonly supportCoverage: SupportCoverage;
}

/**
 * Convert a PPTX file to SVG documents.
 *
 * @param input PPTX binary data.
 * @param options Conversion options. `slides` uses 1-based slide numbers; when
 * omitted, all slides are converted.
 * @returns A conversion report containing converted slides, diagnostics, and support coverage.
 *
 * Text is emitted as SVG paths by default for portable rendering. Set
 * `textOutput: "text"` to emit native `<text>` elements with embedded subset
 * fonts for inline browser SVG use. Font directories and font mapping options
 * control how PPTX font names are resolved for text measurement and text output.
 */
export async function convertPptxToSvg(
  input: Uint8Array,
  options?: ConvertOptions,
): Promise<SvgConversionReport> {
  return convertPptxToSvgBase(input, options, loadSystemFontSetup);
}

/**
 * Render SVG documents from an already parsed PptxSourceModel.
 *
 * Use this with `readPptx()` from `@pptx-glimpse/document` to repeatedly render
 * slides without unzipping and parsing the PPTX bytes again.
 */
export async function renderPptxSourceModelToSvg(
  source: PptxSourceModel,
  options?: ConvertOptions,
): Promise<SvgConversionReport> {
  return renderPptxSourceModelToSvgBase(source, options, loadSystemFontSetup);
}

/**
 * Convert a PPTX file to PNG images.
 *
 * @param input PPTX binary data.
 * @param options Conversion options. `slides` uses 1-based slide numbers; when
 * omitted, all slides are converted.
 * @returns A conversion report containing converted PNG slides, diagnostics, and support coverage.
 *
 * PNG conversion first renders each slide to SVG and then rasterizes it with
 * resvg. The `textOutput` option is intentionally ignored: PNG rendering always
 * uses path-based text output because resvg does not interpret the embedded
 * `@font-face` rules used by SVG text mode. Font directories and font mapping
 * options are still used to resolve glyph outlines and text metrics.
 */
export async function convertPptxToPng(
  input: Uint8Array,
  options?: ConvertOptions,
): Promise<PngConversionReport> {
  const svgResult = await convertPptxToSvg(input, {
    ...options,
    textOutput: "path",
  });
  const width = options?.width ?? DEFAULT_OUTPUT_WIDTH;
  const height = options?.height;
  const fontBuffers = await loadPngFontBuffers(options);

  const slides: SlideImage[] = [];
  for (const { slideNumber, svg } of svgResult.slides) {
    const pngResult = await convertSvgToPng(svg, { width, height, fontBuffers });
    slides.push({
      slideNumber,
      png: toPlainUint8Array(pngResult.png),
      width: pngResult.width,
      height: pngResult.height,
    });
  }

  return {
    slides,
    diagnostics: svgResult.diagnostics,
    supportCoverage: svgResult.supportCoverage,
  };
}

const loadSystemFontSetup: SystemFontSetupLoader = async (options) => {
  if (!shouldLoadSystemFonts(options)) {
    return null;
  }

  const { createOpentypeSetupFromSystem } = await import(
    /* @vite-ignore */ "@pptx-glimpse/renderer/node"
  );
  return createOpentypeSetupFromSystem(
    options?.fontDirs,
    options?.fontMapping,
    options?.skipSystemFonts,
  );
};

async function loadPngFontBuffers(options: ConvertOptions | undefined): Promise<Uint8Array[]> {
  if (options?.fonts !== undefined) {
    return options.fonts.map((font) => toUint8Array(font.data));
  }
  if (!shouldLoadSystemFonts(options)) {
    return [];
  }

  const { loadFontBuffersFromSystem } = await import(/* @vite-ignore */ "./node-font-loader.js");
  return loadFontBuffersFromSystem(options?.fontDirs, options?.skipSystemFonts);
}

async function convertSvgToPng(
  svg: string,
  options: { width?: number; height?: number; fontBuffers?: Uint8Array[] },
) {
  const { svgToPng } = await import("@pptx-glimpse/renderer/png");
  return svgToPng(svg, options);
}

function shouldLoadSystemFonts(options: ConvertOptions | undefined): boolean {
  return options?.skipSystemFonts !== true || (options?.fontDirs?.length ?? 0) > 0;
}

function toUint8Array(data: ArrayBuffer | Uint8Array): Uint8Array {
  return data instanceof Uint8Array ? data : new Uint8Array(data);
}

function toPlainUint8Array(data: Uint8Array): Uint8Array {
  return new Uint8Array(data);
}
