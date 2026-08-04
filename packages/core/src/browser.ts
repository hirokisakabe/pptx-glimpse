/**
 * Public browser API for PPTX rendering, inspection, and editing.
 *
 * Browser callers provide font bytes directly and must initialize the PNG WebAssembly runtime
 * before calling {@link convertPptxToPng}. SVG conversion and editing do not require that step.
 *
 * @module browser
 */

import type { PptxSourceModel } from "@pptx-glimpse/document";
import { DEFAULT_OUTPUT_WIDTH } from "@pptx-glimpse/renderer";
import {
  initResvgWasm as initRendererResvgWasm,
  svgToPng,
} from "@pptx-glimpse/renderer/png/browser";

import {
  affectedSlidePartPaths,
  initializePptxEditorSession,
  type PptxEditorRenderOptions,
  PptxEditorSession as BasePptxEditorSession,
} from "./pptx-editor-session.js";
import {
  type ConvertOptions,
  convertPptxToSvg as convertPptxToSvgBase,
  renderPptxComputedViewToSvg as renderPptxComputedViewToSvgForEditor,
  renderPptxSourceModelToSvg as renderPptxSourceModelToSvgForEditor,
} from "./svg-converter.js";

/** Headless read/edit/render/write session using the browser renderer. */
export class PptxEditorSession extends BasePptxEditorSession {
  private constructor(source: PptxSourceModel, renderOptions: PptxEditorRenderOptions) {
    super(source, renderOptions, {
      renderToSvg: renderPptxSourceModelToSvgForEditor,
      renderComputedToSvg: renderPptxComputedViewToSvgForEditor,
      resolveAffectedSlides: affectedSlidePartPaths,
    });
  }

  static override async create(
    input: Uint8Array,
    renderOptions: PptxEditorRenderOptions = {},
  ): Promise<PptxEditorSession> {
    return initializePptxEditorSession(
      input,
      (source) => new PptxEditorSession(source, renderOptions),
    );
  }
}

/** Parse PPTX bytes and create a rendered browser editor session. */
export function createPptxEditorSession(
  input: Uint8Array,
  renderOptions?: PptxEditorRenderOptions,
): Promise<PptxEditorSession> {
  return PptxEditorSession.create(input, renderOptions);
}

export type { PngConversionReport, SlideImage } from "./converter.js";
export type { UsedFonts } from "./font/font-collector.js";
export { collectUsedFonts } from "./font/font-collector.js";
export type {
  PptxEditorAddConnectorOptions,
  PptxEditorAddTextBoxOptions,
  PptxEditorErrorCode,
  PptxEditorHistoryState,
  PptxEditorImageReplacementInfo,
  PptxEditorRenderOptions,
  PptxEditorSaveResponse,
  PptxEditorSelectionInfo,
  PptxEditorShapeBoundsPx,
  PptxEditorShapeInfo,
  PptxEditorSlideLayoutCatalogEntry,
  PptxEditorSlideMasterCatalogEntry,
  PptxEditorSlidesResponse,
  PptxEditorSlideSvg,
  PptxEditorTextBodyView,
  PptxEditorTextParagraphView,
  PptxEditorTextRunInfo,
  PptxEditorTextRunView,
  PptxEditorTemplatePreviewErrorCode,
  PptxEditorTemplatePreviewFailure,
  PptxEditorTemplatePreviewResult,
  PptxEditorTemplatePreviewSuccess,
  PptxEditorTemplatePreviewTargetKind,
} from "./pptx-editor-session.js";
export { isPptxEditorError, PptxEditorError } from "./pptx-editor-session.js";
export type {
  ConversionDiagnostic,
  ConvertOptions,
  PptxSourceModel,
  SlideSupportCoverage,
  SlideSvg,
  SupportCoverage,
  SupportCoverageCounts,
  SvgConversionReport,
} from "./svg-converter.js";
export { convertPptxToSvg, renderPptxSourceModelToSvg } from "./svg-converter.js";
export type { SourceHandle, UpdateThemeSchemeInput } from "@pptx-glimpse/document";
export type {
  EditorCommand,
  EditorCommandWarning,
  EditorOperationErrorCode,
  EditorOperationFailure,
  GroupableSourceShape,
  GroupShapesCommand,
  UngroupShapeCommand,
  UpdateThemeSchemeCommand,
} from "@pptx-glimpse/editor";
export type { FontMapping } from "@pptx-glimpse/renderer";
export type { FontBuffer, OpentypeSetup } from "@pptx-glimpse/renderer";
export type { LogLevel, WarningEntry, WarningSummary } from "@pptx-glimpse/renderer";
export { createFontMapping, DEFAULT_FONT_MAPPING, getMappedFont } from "@pptx-glimpse/renderer";
export {
  clearFontCache,
  createOpentypeSetupFromBuffers,
  createOpentypeTextMeasurerFromBuffers,
} from "@pptx-glimpse/renderer";
export { getWarningEntries, getWarningSummary } from "@pptx-glimpse/renderer";

/**
 * Accepted WebAssembly inputs for browser PNG rasterization.
 */
export type ResvgWasmInput = ArrayBuffer | Uint8Array | Response;

/**
 * Convert PPTX bytes to PNG images in a browser.
 *
 * Call {@link initResvgWasm} once before using this function. Font directories and system font
 * scanning are unavailable in browsers; provide font bytes through `options.fonts`.
 *
 * @param input PPTX binary data.
 * @param options Conversion options. PNG conversion requests path-based text; unresolved fonts
 * can still fall back to native SVG text before rasterization.
 * @returns Converted slides, diagnostics, and support coverage.
 * @throws An error if resvg WebAssembly has not been initialized.
 */
export async function convertPptxToPng(
  input: Uint8Array,
  options?: ConvertOptions,
): Promise<import("./converter.js").PngConversionReport> {
  const svgResult = await convertPptxToSvgBase(input, {
    ...options,
    textOutput: "path",
  });
  const width = options?.width ?? DEFAULT_OUTPUT_WIDTH;
  const height = options?.height;
  const fontBuffers = options?.fonts?.map((font) => toUint8Array(font.data)) ?? [];

  const slides: import("./converter.js").SlideImage[] = [];
  for (const { slideNumber, svg } of svgResult.slides) {
    const pngResult = await svgToPng(svg, { width, height, fontBuffers });
    slides.push({
      slideNumber,
      png: new Uint8Array(pngResult.png),
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

/**
 * Initialize browser PNG rasterization with resvg WebAssembly.
 *
 * SVG conversion and editor sessions do not require this initialization.
 *
 * @param wasm WebAssembly bytes or a fetched `Response`.
 * @returns A promise that resolves after initialization.
 */
export function initResvgWasm(wasm: ResvgWasmInput): Promise<void> {
  return initRendererResvgWasm(wasm);
}

function toUint8Array(data: ArrayBuffer | Uint8Array): Uint8Array {
  return data instanceof Uint8Array ? data : new Uint8Array(data);
}
