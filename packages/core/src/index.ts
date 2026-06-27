export type { ConvertOptions, SlideImage, SlideSvg } from "./converter.js";
export { convertPptxToPng, convertPptxToSvg } from "./converter.js";
export type { UsedFonts } from "./font/font-collector.js";
export { collectUsedFonts } from "./font/font-collector.js";
export type { FontMapping } from "@pptx-glimpse/renderer";
export type { FontBuffer, OpentypeSetup } from "@pptx-glimpse/renderer";
export type { LogLevel, WarningEntry, WarningSummary } from "@pptx-glimpse/renderer";
export { createFontMapping, DEFAULT_FONT_MAPPING, getMappedFont } from "@pptx-glimpse/renderer";
export {
  clearFontCache,
  createOpentypeSetupFromBuffers,
  createOpentypeTextMeasurerFromBuffers,
} from "@pptx-glimpse/renderer";
export { initResvgWasm } from "@pptx-glimpse/renderer";
export { getWarningEntries, getWarningSummary } from "@pptx-glimpse/renderer";
