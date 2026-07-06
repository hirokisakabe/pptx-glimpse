/**
 * Top-level types for the PptxSourceModel source model.
 *
 * This is the canonical PPTX document representation owned by
 * `@pptx-glimpse/document`. Rather than exposing package parts directly as the public
 * API, it groups presentation, slides, layouts, masters, themes, relationships, media,
 * and content types as OOXML source semantics. Upper layers such as core, editor-core,
 * and pom may consume this package, but this package must not depend on them. Renderer
 * output is produced by the core adapter, and PptxSourceModel does not know about it.
 *
 * This model is the source of truth for writer, editor, and round-trip workflows. It
 * keeps source-local values, relationship ids, part paths, element ordering, typed
 * PPTX-domain units, stable source handles, diagnostics, and raw preservation hooks.
 * Unsupported OOXML, vendor extensions, mc:AlternateContent, and unsupported DrawingML
 * are not mixed into the typed operation API. They are preserved as raw sidecars or raw
 * package parts for structural round-tripping.
 * Editing therefore carries a deliberate double representation: supported changes
 * update typed nodes immediately and append edit records, while the writer later
 * patches preserved raw bytes. Operations that cannot soundly merge those pending
 * edits into raw material, such as duplicating a slide with dirty slide-part edits, are
 * rejected at runtime instead of guessing.
 *
 * New-content edits (new slides, text boxes, connectors) finalize their XML and id
 * numbering at edit time and record them on the edit; the writer only applies
 * insertion positions. To keep the edited in-memory model and the written XML derived
 * from that single finalized fragment, `source/shape-xml.ts` and the edit-time slide
 * id numbering intentionally reference the package-local reader. This is the one
 * sanctioned source -> reader dependency; it stays inside `@pptx-glimpse/document`
 * and does not change the package's external dependency direction.
 *
 * PptxSourceModel must not include renderer-specific fallbacks, environment-specific
 * font substitution, SVG/PNG output, pixel-output values, or pom authoring primitives.
 * Slide/layout/master/theme cascades, relationship resolution, theme color resolution,
 * placeholder and text style resolution, and similar effective values are derived from
 * the source as a non-mutating computed view.
 */

import type { Diagnostic } from "./diagnostics.js";
import type { PartPath, RelationshipId, SourceHandle } from "./handles.js";
import type { PackageGraph } from "./package-graph.js";
import type {
  SourcePresentation,
  SourceSlide,
  SourceSlideLayout,
  SourceSlideMaster,
  SourceTheme,
} from "./presentation.js";
import type { Emu, Pt } from "./units.js";

export interface PptxSourceModel {
  /** Structure of package part / relationship / content type / media. */
  readonly packageGraph: PackageGraph;
  readonly presentation: SourcePresentation;
  readonly slides: readonly SourceSlide[];
  readonly slideLayouts: readonly SourceSlideLayout[];
  readonly slideMasters: readonly SourceSlideMaster[];
  readonly themes: readonly SourceTheme[];
  /** Diagnostics about document correctness. */
  readonly diagnostics: readonly Diagnostic[];
  /** typed PptxSourceModel operation and dirty scope. The writer determines the minimum update range from this. */
  readonly edits?: readonly PptxSourceModelEdit[];
}

export type PptxSourceModelEdit =
  | PptxSourceModelTextRunEdit
  | PptxSourceModelTextRunPropertiesEdit
  | PptxSourceModelParagraphTextEdit
  | PptxSourceModelShapeTransformEdit
  | PptxSourceModelAddTextBoxEdit
  | PptxSourceModelAddConnectorEdit
  | PptxSourceModelDeleteShapeEdit
  | PptxSourceModelReplaceImageEdit
  | PptxSourceModelAddEmptySlideFromLayoutEdit
  | PptxSourceModelDuplicateSlideEdit
  | PptxSourceModelMoveSlideEdit
  | PptxSourceModelDeleteSlideEdit;

export interface PptxSourceModelTextRunEdit {
  readonly kind: "replaceTextRunPlainText";
  readonly handle: SourceHandle;
  readonly text: string;
}

export type EditableTextRunProperty =
  | "bold"
  | "italic"
  | "underline"
  | "fontSize"
  | "color"
  | "typeface";

export interface EditableTextRunProperties {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly fontSize?: Pt;
  readonly color?: { readonly kind: "srgb"; readonly hex: string };
  readonly typeface?: string;
}

export interface PptxSourceModelTextRunPropertiesEdit {
  readonly kind: "updateTextRunProperties";
  readonly handle: SourceHandle;
  readonly set?: EditableTextRunProperties;
  readonly clear?: readonly EditableTextRunProperty[];
}

export interface PptxSourceModelParagraphTextEdit {
  readonly kind: "replaceParagraphPlainText";
  readonly handle: SourceHandle;
  readonly text: string;
}

export interface PptxSourceModelShapeTransformEdit {
  readonly kind: "updateShapeTransform";
  readonly handle: SourceHandle;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
}

export interface PptxSourceModelAddTextBoxEdit {
  readonly kind: "addTextBox";
  readonly slidePartPath: PartPath;
  readonly shapeId: string;
  /** Serialized `p:sp` fragment finalized at edit time. The writer only splices it. */
  readonly xml: string;
}

export type ConnectorPresetGeometry = "straightConnector1" | "bentConnector3" | "curvedConnector3";

export interface PptxSourceModelAddConnectorEdit {
  readonly kind: "addConnector";
  readonly slidePartPath: PartPath;
  readonly shapeId: string;
  readonly startShapeId?: string;
  readonly endShapeId?: string;
  /** Serialized `p:cxnSp` fragment finalized at edit time. The writer only splices it. */
  readonly xml: string;
}

export interface PptxSourceModelDeleteShapeEdit {
  readonly kind: "deleteShape";
  readonly handle: SourceHandle;
}

export interface PptxSourceModelReplaceImageEdit {
  readonly kind: "replaceImage";
  readonly handle: SourceHandle;
  readonly mediaPartPath: PartPath;
  readonly contentType: string;
  readonly sharedReferenceCount: number;
}

export interface PptxSourceModelAddEmptySlideFromLayoutEdit {
  readonly kind: "addEmptySlideFromLayout";
  readonly layoutPartPath: PartPath;
  readonly newSlidePartPath: PartPath;
  readonly newRelationshipId: RelationshipId;
  /** Numeric `p:sldId@id` assigned at edit time. The writer only applies it. */
  readonly newSlideNumericId: number;
}

export interface PptxSourceModelDuplicateSlideEdit {
  readonly kind: "duplicateSlide";
  readonly sourceSlidePartPath: PartPath;
  readonly sourceRelationshipId: RelationshipId;
  readonly newSlidePartPath: PartPath;
  readonly newRelationshipId: RelationshipId;
  /** Numeric `p:sldId@id` assigned at edit time. The writer only applies it. */
  readonly newSlideNumericId: number;
}

export interface PptxSourceModelMoveSlideEdit {
  readonly kind: "moveSlide";
  readonly slidePartPath: PartPath;
  readonly relationshipId: RelationshipId;
  /** Zero-based final index in the slide list at the time this edit is applied. */
  readonly toIndex: number;
}

export interface PptxSourceModelDeleteSlideEdit {
  readonly kind: "deleteSlide";
  readonly slidePartPath: PartPath;
  readonly relationshipId: RelationshipId;
}
