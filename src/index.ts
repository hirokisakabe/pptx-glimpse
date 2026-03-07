export type { ConvertOptions, SlideImage, SlideSvg } from "./converter.js";
export { convertPptxToPng, convertPptxToSvg } from "./converter.js";
export type { UsedFonts } from "./font/font-collector.js";
export { collectUsedFonts } from "./font/font-collector.js";
export type { FontMapping } from "./font/font-mapping.js";
export { createFontMapping, DEFAULT_FONT_MAPPING, getMappedFont } from "./font/font-mapping.js";
export type { FontBuffer, OpentypeSetup } from "./font/opentype-helpers.js";
export {
  createOpentypeSetupFromBuffers,
  createOpentypeTextMeasurerFromBuffers,
} from "./font/opentype-helpers.js";
export type { LogLevel, WarningEntry, WarningSummary } from "./warning-logger.js";
export { getWarningEntries, getWarningSummary } from "./warning-logger.js";
