import { renderPptxSourceModelToSvg as renderPptxSourceModelToSvgForEditor } from "./converter.js";
import { configurePptxEditorSessionRenderer } from "./pptx-editor-session.js";

configurePptxEditorSessionRenderer(renderPptxSourceModelToSvgForEditor);

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
  PptxEditorHistoryState,
  PptxEditorImageReplacementInfo,
  PptxEditorRenderOptions,
  PptxEditorSaveResponse,
  PptxEditorSelectionInfo,
  PptxEditorShapeBoundsPx,
  PptxEditorShapeInfo,
  PptxEditorSlidesResponse,
  PptxEditorSlideSvg,
  PptxEditorTextBodyInfo,
  PptxEditorTextBodyView,
  PptxEditorTextParagraphView,
  PptxEditorTextRunInfo,
  PptxEditorTextRunView,
} from "./pptx-editor-session.js";
export { createPptxEditorSession, PptxEditorSession } from "./pptx-editor-session.js";
export type { SourceHandle } from "@pptx-glimpse/document";
export type { EditorCommand, EditorCommandWarning } from "@pptx-glimpse/editor";
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

export type ResvgWasmInput = ArrayBuffer | Uint8Array | Response;

export async function initResvgWasm(wasm?: ResvgWasmInput): Promise<void> {
  const { initResvgWasm: initRendererResvgWasm } = await import("@pptx-glimpse/renderer/png");
  return initRendererResvgWasm(wasm);
}
