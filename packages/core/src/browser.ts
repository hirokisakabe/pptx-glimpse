import { DEFAULT_OUTPUT_WIDTH } from "@pptx-glimpse/renderer";
import {
  initResvgWasm as initRendererResvgWasm,
  type ResvgWasmInput,
  svgToPng,
} from "@pptx-glimpse/renderer/png/browser";

import { type ConvertOptions, convertPptxToSvg as convertPptxToSvgBase } from "./svg-converter.js";

export type {
  BrowserEditorAddTextBoxOptions,
  BrowserEditorHistoryState,
  BrowserEditorRenderOptions,
  BrowserEditorSaveResponse,
  BrowserEditorSelectionInfo,
  BrowserEditorShapeBoundsPx,
  BrowserEditorShapeInfo,
  BrowserEditorSlidesResponse,
  BrowserEditorTextBodyInfo,
  BrowserEditorTextRunInfo,
} from "./browser-editor.js";
export { BrowserPptxEditorSession, createBrowserPptxEditorSession } from "./browser-editor.js";
export type { PngConversionReport, SlideImage } from "./converter.js";
export type { UsedFonts } from "./font/font-collector.js";
export { collectUsedFonts } from "./font/font-collector.js";
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
export type { SourceHandle } from "@pptx-glimpse/document";
export type { EditorCommand, EditorCommandWarning } from "@pptx-glimpse/editor-core";
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
export type { ResvgWasmInput } from "@pptx-glimpse/renderer/png/browser";

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

export function initResvgWasm(wasm: ResvgWasmInput): Promise<void> {
  return initRendererResvgWasm(wasm);
}

function toUint8Array(data: ArrayBuffer | Uint8Array): Uint8Array {
  return data instanceof Uint8Array ? data : new Uint8Array(data);
}
