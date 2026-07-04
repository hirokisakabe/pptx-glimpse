export type { PngConversionReport, SlideImage } from "./converter.js";
export type { UsedFonts } from "./font/font-collector.js";
export { collectUsedFonts } from "./font/font-collector.js";
export type {
  ConversionDiagnostic,
  ConvertOptions,
  SlideSupportCoverage,
  SlideSvg,
  SupportCoverage,
  SupportCoverageCounts,
  SvgConversionReport,
} from "./svg-converter.js";
export { convertPptxToSvg } from "./svg-converter.js";
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
export type { ResvgWasmInput } from "@pptx-glimpse/renderer/png";

export function convertPptxToPng(): Promise<never> {
  return Promise.reject(
    new Error(
      "convertPptxToPng is not available from the browser entry. Use convertPptxToSvg or the Node.js entry.",
    ),
  );
}

export async function initResvgWasm(
  wasm: import("@pptx-glimpse/renderer/png").ResvgWasmInput,
): Promise<void> {
  const { initResvgWasm: initRendererResvgWasm } = await import("@pptx-glimpse/renderer/png");
  return initRendererResvgWasm(wasm);
}
