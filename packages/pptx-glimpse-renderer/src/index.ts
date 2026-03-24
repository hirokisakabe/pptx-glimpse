// --- model ---
export type {
  ChartData,
  ChartElement,
  ChartLegend,
  ChartSeries,
  ChartType,
} from "./model/chart.js";
export type {
  BiLevelEffect,
  BlipEffects,
  BlurEffect,
  ClrChangeEffect,
  DuotoneEffect,
  EffectList,
  Glow,
  InnerShadow,
  LumEffect,
  OuterShadow,
  SoftEdge,
} from "./model/effect.js";
export type {
  Fill,
  GradientFill,
  GradientStop,
  ImageFill,
  ImageFillTile,
  NoFill,
  PatternFill,
  ResolvedColor,
  SolidFill,
} from "./model/fill.js";
export type { ImageElement, SrcRect, StretchFillRect, TileInfo } from "./model/image.js";
export type {
  ArrowEndpoint,
  ArrowSize,
  ArrowType,
  DashStyle,
  LineCap,
  LineJoin,
  Outline,
} from "./model/line.js";
export type { EmbeddedFont, Protection, SlideSize } from "./model/presentation.js";
export type {
  ConnectorElement,
  CustomGeometry,
  CustomGeometryPath,
  Geometry,
  GroupElement,
  PresetGeometry,
  ShapeElement,
  SlideElement,
  Transform,
} from "./model/shape.js";
export type { Background, Slide } from "./model/slide.js";
export type {
  CellBorders,
  TableCell,
  TableColumn,
  TableData,
  TableElement,
  TableRow,
} from "./model/table.js";
export type {
  AutoNumScheme,
  BodyProperties,
  BulletType,
  DefaultParagraphLevelProperties,
  DefaultRunProperties,
  DefaultTextStyle,
  Hyperlink,
  Paragraph,
  ParagraphProperties,
  PlaceholderStyleInfo,
  RunProperties,
  SpacingValue,
  TabStop,
  TextBody,
  TextOutline,
  TextRun,
  TextVerticalType,
  TxStyles,
} from "./model/text.js";
export type {
  ColorMap,
  ColorScheme,
  ColorSchemeKey,
  FontScheme,
  FormatScheme,
  Theme,
} from "./model/theme.js";

// --- utils ---
export { uint8ArrayToBase64 } from "./utils/base64.js";
export {
  DEFAULT_DPI,
  DEFAULT_OUTPUT_WIDTH,
  EMU_PER_INCH,
  EMU_PER_POINT,
  ROTATION_UNIT,
} from "./utils/constants.js";
export { emuToPixels, emuToPoints, hundredthPointToPoint, rotationToDegrees } from "./utils/emu.js";
export {
  getAscenderRatio,
  getLineHeightRatio,
  isCjkCodePoint,
  measureTextWidth,
} from "./utils/text-measure.js";
export { wrapParagraph } from "./utils/text-wrap.js";
export type { Emu, HundredthPt, Pt } from "./utils/unit-types.js";
export { asEmu, asHundredthPt } from "./utils/unit-types.js";

// --- data ---
export type { FontMetrics } from "./data/font-metrics.js";
export { getFontMetrics, getMetricsFallbackFont } from "./data/font-metrics.js";

// --- font ---
export type { FontMapping } from "./font/font-mapping.js";
export { createFontMapping, DEFAULT_FONT_MAPPING, getMappedFont } from "./font/font-mapping.js";
export {
  getCurrentMappedFont,
  resetFontMapping,
  setFontMapping,
} from "./font/font-mapping-context.js";
export type { FontBuffer, OpentypeSetup } from "./font/opentype-helpers.js";
export {
  createOpentypeSetupFromBuffers,
  createOpentypeSetupFromSystem,
  createOpentypeTextMeasurerFromBuffers,
} from "./font/opentype-helpers.js";
export type { OpentypeFont } from "./font/opentype-text-measurer.js";
export { OpentypeTextMeasurer } from "./font/opentype-text-measurer.js";
export {
  getJpanFallbackFont,
  resetScriptFonts,
  setScriptFonts,
} from "./font/script-font-context.js";
export { collectFontFilePaths } from "./font/system-font-loader.js";
export type { TextMeasurer } from "./font/text-measurer.js";
export {
  DefaultTextMeasurer,
  getTextMeasurer,
  resetTextMeasurer,
  setTextMeasurer,
} from "./font/text-measurer.js";
export type {
  OpentypeFullFont,
  OpentypePath,
  TextPathFontResolver,
} from "./font/text-path-context.js";
export {
  DefaultTextPathFontResolver,
  getTextPathFontResolver,
  resetTextPathFontResolver,
  setTextPathFontResolver,
} from "./font/text-path-context.js";
export { extractTtcFonts, isTtcBuffer } from "./font/ttc-parser.js";

// --- renderer ---
export type { RenderResult } from "./renderer/render-result.js";
export { renderShape } from "./renderer/shape-renderer.js";
export { renderSlideToSvg } from "./renderer/svg-renderer.js";

// --- png ---
export { svgToPng } from "./png/png-converter.js";

// --- warning-logger ---
export type { LogLevel, WarningEntry, WarningSummary } from "./warning-logger.js";
export {
  debug,
  flushWarnings,
  getLogLevel,
  getWarningEntries,
  getWarningSummary,
  initWarningLogger,
  warn,
} from "./warning-logger.js";
