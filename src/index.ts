export { convertPptxToPng, convertPptxToSvg } from "./converter.js";
export type { ConvertOptions, SlideImage, SlideSvg } from "./converter.js";
export type { LogLevel, WarningSummary, WarningEntry } from "./warning-logger.js";
export { getWarningSummary, getWarningEntries } from "./warning-logger.js";
export { collectUsedFonts } from "./font/font-collector.js";
export type { UsedFonts } from "./font/font-collector.js";
export { DEFAULT_FONT_MAPPING, createFontMapping, getMappedFont } from "./font/font-mapping.js";
export type { FontMapping } from "./font/font-mapping.js";
