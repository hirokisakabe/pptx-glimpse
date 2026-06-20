/**
 * CleanDoc source model 型の barrel re-export。
 */

export type { CleanDocSource } from "./clean-doc-source.js";
export type { Diagnostic, DiagnosticSeverity } from "./diagnostics.js";
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
  SourcePresentation,
  SourceSlide,
  SourceSlideLayout,
  SourceSlideMaster,
  SourceTheme,
} from "./presentation.js";
export type { RawOoxmlNode, RawPackagePart, RawSidecar } from "./raw.js";
export type {
  SourceColor,
  SourceColorTransform,
  SourceFill,
  SourceImage,
  SourceParagraph,
  SourceParagraphProperties,
  SourcePlaceholder,
  SourcePresetGeometry,
  SourceRawShapeNode,
  SourceRunProperties,
  SourceShape,
  SourceShapeNode,
  SourceTextAlign,
  SourceTextBody,
  SourceTextRun,
  SourceTransform,
} from "./shapes.js";
export type { Emu, HundredthPt, OoxmlAngle, OoxmlPercent, Pt } from "./units.js";
export { asEmu, asHundredthPt, asOoxmlAngle, asOoxmlPercent, asPt } from "./units.js";
