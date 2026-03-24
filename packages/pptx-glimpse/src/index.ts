export type { ConvertOptions, SlideImage, SlideSvg } from "./converter.js";
export { convertPptxToPng, convertPptxToSvg } from "./converter.js";
export type { UsedFonts } from "./font-collector.js";
export { collectUsedFonts } from "./font-collector.js";

// Re-export from pptx-glimpse-renderer
export type { FontBuffer, FontMapping, OpentypeSetup } from "pptx-glimpse-renderer";
export type { LogLevel, WarningEntry, WarningSummary } from "pptx-glimpse-renderer";
export {
  createFontMapping,
  createOpentypeSetupFromBuffers,
  createOpentypeTextMeasurerFromBuffers,
  DEFAULT_FONT_MAPPING,
  getMappedFont,
} from "pptx-glimpse-renderer";
export { getWarningEntries, getWarningSummary } from "pptx-glimpse-renderer";
