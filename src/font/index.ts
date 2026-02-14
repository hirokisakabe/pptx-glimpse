export { DEFAULT_FONT_MAPPING, createFontMapping, getMappedFont } from "./font-mapping.js";
export type { FontMapping } from "./font-mapping.js";
export { setFontMapping, getFontMapping, resetFontMapping } from "./font-mapping-context.js";
export { getCurrentMappedFont } from "./font-mapping-context.js";
export { collectUsedFonts } from "./font-collector.js";
export type { UsedFonts } from "./font-collector.js";
export {
  setTextMeasurer,
  getTextMeasurer,
  resetTextMeasurer,
  DefaultTextMeasurer,
} from "./text-measurer.js";
export type { TextMeasurer } from "./text-measurer.js";
export {
  setTextPathFontResolver,
  getTextPathFontResolver,
  resetTextPathFontResolver,
  DefaultTextPathFontResolver,
} from "./text-path-context.js";
export type { TextPathFontResolver, OpentypeFullFont, OpentypePath } from "./text-path-context.js";
export {
  createOpentypeTextMeasurerFromBuffers,
  createOpentypeSetupFromBuffers,
  createOpentypeSetupFromSystem,
} from "./opentype-helpers.js";
export type { FontBuffer, OpentypeSetup } from "./opentype-helpers.js";
export { OpentypeTextMeasurer } from "./opentype-text-measurer.js";
export type { OpentypeFont } from "./opentype-text-measurer.js";
export { collectFontFilePaths } from "./system-font-loader.js";
