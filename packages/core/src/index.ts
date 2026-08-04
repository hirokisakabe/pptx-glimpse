/**
 * Public Node.js API for PPTX rendering, inspection, and editing.
 *
 * Node.js callers can use filesystem-backed font discovery as well as caller-supplied font
 * buffers. See the browser entry point for browser-specific PNG initialization requirements.
 *
 * @module node
 */

import type { PptxSourceModel } from "@pptx-glimpse/document";

import { renderPptxSourceModelToSvg as renderPptxSourceModelToSvgForEditor } from "./converter.js";
import {
  affectedSlidePartPaths,
  initializePptxEditorSession,
  type PptxEditorRenderOptions,
  PptxEditorSession as BasePptxEditorSession,
} from "./pptx-editor-session.js";

/** Headless read/edit/render/write session using the Node renderer. */
export class PptxEditorSession extends BasePptxEditorSession {
  private constructor(source: PptxSourceModel, renderOptions: PptxEditorRenderOptions) {
    super(source, renderOptions, {
      renderToSvg: renderPptxSourceModelToSvgForEditor,
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

/** Parse PPTX bytes and create a rendered Node editor session. */
export function createPptxEditorSession(
  input: Uint8Array,
  renderOptions?: PptxEditorRenderOptions,
): Promise<PptxEditorSession> {
  return PptxEditorSession.create(input, renderOptions);
}

export type {
  ConversionDiagnostic,
  ConvertOptions,
  PngConversionReport,
  PptxSourceModel,
  SlideImage,
  SlideSupportCoverage,
  SlideSvg,
  SupportCoverage,
  SupportCoverageCounts,
  SvgConversionReport,
} from "./converter.js";
export { convertPptxToPng, convertPptxToSvg, renderPptxSourceModelToSvg } from "./converter.js";
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
} from "./pptx-editor-session.js";
export { isPptxEditorError, PptxEditorError } from "./pptx-editor-session.js";
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
 * Accepted WebAssembly inputs for PNG rasterization.
 *
 * Node.js can normally initialize the bundled resvg module lazily. Pass bytes or a response when
 * the host runtime requires explicit initialization.
 */
export type ResvgWasmInput = ArrayBuffer | Uint8Array | Response;

/**
 * Initialize the resvg WebAssembly runtime used by {@link convertPptxToPng}.
 *
 * @param wasm Optional WebAssembly bytes or response. When omitted, the Node.js renderer loads
 * its bundled module.
 * @returns A promise that resolves after initialization.
 */
export async function initResvgWasm(wasm?: ResvgWasmInput): Promise<void> {
  const { initResvgWasm: initRendererResvgWasm } = await import("@pptx-glimpse/renderer/png");
  return initRendererResvgWasm(wasm);
}
