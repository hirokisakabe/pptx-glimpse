export type {
  BrowserEditorAddConnectorOptions,
  BrowserEditorAddTextBoxOptions,
  BrowserEditorHistoryState,
  BrowserEditorImageReplacementInfo,
  BrowserEditorRenderOptions,
  BrowserEditorSaveResponse,
  BrowserEditorSelectionInfo,
  BrowserEditorShapeBoundsPx,
  BrowserEditorShapeInfo,
  BrowserEditorSlidesResponse,
  BrowserEditorSlideSvg,
  BrowserEditorTextBodyInfo,
  BrowserEditorTextBodyView,
  BrowserEditorTextParagraphView,
  BrowserEditorTextRunInfo,
  BrowserEditorTextRunView,
} from "./browser-editor.js";
export { BrowserPptxEditorSession, createBrowserPptxEditorSession } from "./browser-editor.js";
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
