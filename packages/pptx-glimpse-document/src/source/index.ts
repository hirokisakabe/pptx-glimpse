/**
 * CleanDoc source model 型の barrel re-export。
 */

export type { CleanDocEdit, CleanDocSource, CleanDocTextRunEdit } from "./clean-doc-source.js";
export type { Diagnostic, DiagnosticSeverity } from "./diagnostics.js";
export { findTextRunBySourceHandle, replaceTextRunPlainText } from "./editing.js";
export type {
  PartPath,
  RawSidecarId,
  RelationshipId,
  SourceHandle,
  SourceNodeId,
} from "./handles.js";
export { asPartPath, asRawSidecarId, asRelationshipId, asSourceNodeId } from "./handles.js";
export type {
  ContentTypeDefault,
  ContentTypeOverride,
  ContentTypes,
  MediaPart,
  PackageGraph,
  PackagePartRef,
  PartRelationships,
  Relationship,
  RelationshipTargetMode,
} from "./package-graph.js";
export type {
  SlideSize,
  SourceBackground,
  SourceColorMap,
  SourcePresentation,
  SourceSlide,
  SourceSlideLayout,
  SourceSlideMaster,
  SourceTheme,
  SourceThemeColorScheme,
  SourceThemeFontScheme,
  SourceThemeFormatScheme,
} from "./presentation.js";
export type { RawOoxmlNode, RawPackagePart, RawSidecar } from "./raw.js";
export type {
  SourceArrowEndpoint,
  SourceArrowSize,
  SourceArrowType,
  SourceCellBorders,
  SourceChart,
  SourceColor,
  SourceColorTransform,
  SourceConnector,
  SourceCustomGeometry,
  SourceCustomGeometryPath,
  SourceDashStyle,
  SourceEffectList,
  SourceFill,
  SourceGeometry,
  SourceGlow,
  SourceGradient,
  SourceGradientStop,
  SourceGroup,
  SourceImage,
  SourceImageCrop,
  SourceImageFillTile,
  SourceImageStretch,
  SourceInnerShadow,
  SourceLineCap,
  SourceLineJoin,
  SourceOuterShadow,
  SourceOutline,
  SourceParagraph,
  SourceParagraphProperties,
  SourcePlaceholder,
  SourcePresetGeometry,
  SourceRawShapeNode,
  SourceRunProperties,
  SourceShape,
  SourceShapeNode,
  SourceShapeStyle,
  SourceSmartArt,
  SourceSoftEdge,
  SourceStyleReference,
  SourceTable,
  SourceTableCell,
  SourceTableColumn,
  SourceTableData,
  SourceTableRow,
  SourceTextAlign,
  SourceTextBody,
  SourceTextBodyProperties,
  SourceTextRun,
  SourceTransform,
  SourceVerticalAnchor,
} from "./shapes.js";
export type { Emu, HundredthPt, OoxmlAngle, OoxmlPercent, Pt } from "./units.js";
export { asEmu, asHundredthPt, asOoxmlAngle, asOoxmlPercent, asPt } from "./units.js";
