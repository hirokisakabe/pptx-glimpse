/**
 * Top-level types for the PptxSourceModel source model.
 *
 * This is the canonical PPTX document representation owned by
 * `@pptx-glimpse/document`. Rather than exposing package parts directly as the public
 * API, it groups presentation, slides, layouts, masters, themes, relationships, media,
 * and content types as OOXML source semantics. Upper layers such as core, editor,
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
 * New-content edits (new slides, layouts, text boxes, shapes, connectors, pictures) finalize their XML and id
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
 *
 * Repository-level package and dependency boundaries are documented in the
 * [architecture overview](../../../../docs/architecture/overview.md).
 */

import type { Diagnostic } from "./diagnostics.js";
import type { PartPath, RelationshipId, SourceHandle, SourceNodeId } from "./handles.js";
import type { PackageGraph } from "./package-graph.js";
import type {
  SourcePresentation,
  SourceSlide,
  SourceSlideLayout,
  SourceSlideMaster,
  SourceTheme,
} from "./presentation.js";
import type { SourceImageCrop, SourceParagraphProperties, SourceTransform } from "./shapes.js";
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
  | PptxSourceModelParagraphPropertiesEdit
  | PptxSourceModelParagraphTextEdit
  | PptxSourceModelShapeTransformEdit
  | PptxSourceModelShapeFillEdit
  | PptxSourceModelShapeOutlineEdit
  | PptxSourceModelPictureCropEdit
  | PptxSourceModelTableCellPropertiesEdit
  | PptxSourceModelAddShapeEdit
  | PptxSourceModelAddTextBoxEdit
  | PptxSourceModelAddConnectorEdit
  | PptxSourceModelAddPictureEdit
  | PptxSourceModelAddChartEdit
  | PptxSourceModelUpdateChartDataEdit
  | PptxSourceModelUpdateScatterChartDataEdit
  | PptxSourceModelUpdateBubbleChartDataEdit
  | PptxSourceModelAddTableEdit
  | PptxSourceModelReorderShapesEdit
  | PptxSourceModelMoveShapesEdit
  | PptxSourceModelMoveShapesAcrossSlidesEdit
  | PptxSourceModelGroupShapesEdit
  | PptxSourceModelUngroupShapeEdit
  | PptxSourceModelDeleteShapeEdit
  | PptxSourceModelReplaceImageEdit
  | PptxSourceModelAddSlideLayoutEdit
  | PptxSourceModelCloneSlideLayoutEdit
  | PptxSourceModelAddEmptySlideFromLayoutEdit
  | PptxSourceModelDuplicateSlideEdit
  | PptxSourceModelMoveSlideEdit
  | PptxSourceModelDeleteSlideEdit
  | PptxSourceModelUpdateThemeSchemeEdit
  | PptxSourceModelSetBackgroundEdit
  | PptxSourceModelSetSlideBackgroundEdit;

export interface PptxSourceModelTextRunEdit {
  readonly kind: "replaceTextRunPlainText";
  readonly handle: SourceHandle;
  readonly text: string;
}

/** @inline */
export type EditableTextRunProperty =
  | "bold"
  | "italic"
  | "underline"
  | "fontSize"
  | "color"
  | "typeface";

/** @inline */
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

/** @inline */
export type EditableParagraphProperty = "align" | "level" | "bullet";

/** @inline */
export interface EditableParagraphProperties {
  readonly align?: SourceParagraphProperties["align"];
  readonly level?: SourceParagraphProperties["level"];
  readonly bullet?: SourceParagraphProperties["bullet"];
}

export interface PptxSourceModelParagraphPropertiesEdit {
  readonly kind: "updateParagraphProperties";
  readonly handle: SourceHandle;
  readonly set?: EditableParagraphProperties;
  readonly clear?: readonly EditableParagraphProperty[];
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

/** @inline */
export type EditableShapeFill =
  | { readonly kind: "none" }
  | { readonly kind: "solid"; readonly color: { readonly kind: "srgb"; readonly hex: string } };

/** @inline */
export interface EditableShapeOutline {
  readonly width?: Emu;
  readonly fill?: EditableShapeFill;
}

export interface PptxSourceModelShapeFillEdit {
  readonly kind: "updateShapeFill";
  readonly handle: SourceHandle;
  readonly fill: EditableShapeFill;
}

export interface PptxSourceModelShapeOutlineEdit {
  readonly kind: "updateShapeOutline";
  readonly handle: SourceHandle;
  readonly outline: EditableShapeOutline;
}

export interface PptxSourceModelPictureCropEdit {
  readonly kind: "updatePictureCrop";
  readonly handle: SourceHandle;
  /** Omitted to remove the targeted `a:srcRect`. */
  readonly crop?: SourceImageCrop;
}

/** @inline */
export type EditableTableCellProperty =
  | "fill"
  | "borderTop"
  | "borderBottom"
  | "borderLeft"
  | "borderRight"
  | "marginLeft"
  | "marginRight"
  | "marginTop"
  | "marginBottom";

/** @inline */
export interface EditableTableCellBorder {
  readonly width?: Emu;
  readonly fill?: EditableShapeFill;
}

/** @inline */
export interface EditableTableCellProperties {
  readonly fill?: EditableShapeFill;
  readonly borders?: {
    readonly top?: EditableTableCellBorder;
    readonly bottom?: EditableTableCellBorder;
    readonly left?: EditableTableCellBorder;
    readonly right?: EditableTableCellBorder;
  };
  readonly marginLeft?: Emu;
  readonly marginRight?: Emu;
  readonly marginTop?: Emu;
  readonly marginBottom?: Emu;
}

/** Zero-based physical cell address within one native Table. @inline */
export interface TableCellAddress {
  readonly tableHandle: SourceHandle;
  readonly rowIndex: number;
  readonly cellIndex: number;
}

export interface PptxSourceModelTableCellPropertiesEdit {
  readonly kind: "updateTableCellProperties";
  readonly address: TableCellAddress;
  readonly set?: EditableTableCellProperties;
  readonly clear?: readonly EditableTableCellProperty[];
}

export interface PptxSourceModelAddTextBoxEdit {
  readonly kind: "addTextBox";
  readonly slidePartPath: PartPath;
  readonly shapeId: string;
  /** Serialized `p:sp` fragment finalized at edit time. The writer only splices it. */
  readonly xml: string;
}

export interface PptxSourceModelAddShapeEdit {
  readonly kind: "addShape";
  readonly slidePartPath: PartPath;
  readonly shapeId: string;
  /** Serialized `p:sp` fragment finalized at edit time. The writer only splices it. */
  readonly xml: string;
}

/** @inline */
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

export interface PptxSourceModelAddPictureEdit {
  readonly kind: "addPicture";
  readonly slidePartPath: PartPath;
  readonly shapeId: string;
  readonly relationshipId: RelationshipId;
  readonly mediaPartPath: PartPath;
  readonly contentType: string;
  /** Serialized `p:pic` fragment finalized at edit time. The writer only splices it. */
  readonly xml: string;
}

export interface PptxSourceModelAddChartEdit {
  readonly kind: "addChart";
  readonly slidePartPath: PartPath;
  readonly shapeId: string;
  readonly relationshipId: RelationshipId;
  readonly chartPartPath: PartPath;
  readonly workbookPartPath: PartPath;
  /** Serialized `p:graphicFrame` fragment finalized at edit time. */
  readonly xml: string;
}

export interface PptxSourceModelUpdateChartDataEdit {
  readonly kind: "updateChartData";
  readonly handle: SourceHandle;
  readonly chartPartPath: PartPath;
  readonly workbookPartPath: PartPath;
}

export interface PptxSourceModelUpdateScatterChartDataEdit {
  readonly kind: "updateScatterChartData";
  readonly handle: SourceHandle;
  readonly chartPartPath: PartPath;
  readonly workbookPartPath: PartPath;
}

export interface PptxSourceModelUpdateBubbleChartDataEdit {
  readonly kind: "updateBubbleChartData";
  readonly handle: SourceHandle;
  readonly chartPartPath: PartPath;
  readonly workbookPartPath: PartPath;
}

export interface PptxSourceModelAddTableEdit {
  readonly kind: "addTable";
  readonly slidePartPath: PartPath;
  readonly shapeId: string;
  /** Serialized `p:graphicFrame` fragment finalized at edit time. */
  readonly xml: string;
}

export interface PptxSourceModelAddSlideLayoutEdit {
  readonly kind: "addSlideLayout";
  readonly masterPartPath: PartPath;
  readonly newLayoutPartPath: PartPath;
  readonly newRelationshipId: RelationshipId;
  readonly newLayoutNumericId: number;
  /** Existing relationship-order layouts materialized when the master had no id list. */
  readonly initialLayoutEntries: readonly {
    readonly relationshipId: RelationshipId;
    readonly numericId: number;
  }[];
}

export interface PptxSourceModelCloneSlideLayoutEdit {
  readonly kind: "cloneSlideLayout";
  readonly masterPartPath: PartPath;
  readonly sourceLayoutPartPath: PartPath;
  readonly newLayoutPartPath: PartPath;
  readonly newRelationshipId: RelationshipId;
  readonly newLayoutNumericId: number;
  /** Zero-based insertion position within the master's layout id list. */
  readonly insertAt: number;
  /** Existing relationship-order layouts materialized when the master had no id list. */
  readonly initialLayoutEntries: readonly {
    readonly relationshipId: RelationshipId;
    readonly numericId: number;
  }[];
}

export interface PptxSourceModelReorderShapesEdit {
  readonly kind: "reorderShapes";
  readonly targetPartPath: PartPath;
  /** Direct-child container group id. Omitted for the root `p:spTree`. */
  readonly parentGroupId?: string;
  readonly shapeIds: readonly string[];
}

export interface PptxSourceModelMoveShapesEdit {
  readonly kind: "moveShapes";
  readonly targetPartPath: PartPath;
  /** Source direct-child container group id. Omitted for the root `p:spTree`. */
  readonly parentGroupId?: string;
  /** True when source and destination are different direct-child containers. */
  readonly crossParent?: true;
  /** Destination direct-child group id for a cross-parent move. Omitted for the root `p:spTree`. */
  readonly destinationParentGroupId?: string;
  /** Finalized integer OOXML transforms for roots re-expressed across affine parents. */
  readonly transformedRoots?: readonly {
    readonly shapeId: string;
    readonly transform: SourceTransform;
  }[];
  /** Consecutive drawing ids in source document order. */
  readonly shapeIds: readonly string[];
  /** Direct-child anchor id. Omitted to move the block to the drawing-order end. */
  readonly beforeShapeId?: string;
}

/** Finalized two-part XML patch for a slide-root drawing move. */
export interface PptxSourceModelMoveShapesAcrossSlidesEdit {
  readonly kind: "moveShapesAcrossSlides";
  readonly sourcePartPath: PartPath;
  readonly destinationPartPath: PartPath;
  readonly sourceShapeIds: readonly SourceNodeId[];
  readonly destinationShapeIds: readonly SourceNodeId[];
  readonly nodeIdMappings: readonly {
    readonly before: SourceNodeId;
    readonly after: SourceNodeId;
  }[];
  readonly relationshipIdMappings: readonly {
    readonly before: RelationshipId;
    readonly after: RelationshipId;
  }[];
  readonly beforeShapeId?: SourceNodeId;
}

export interface PptxSourceModelGroupShapesEdit {
  readonly kind: "groupShapes";
  readonly targetPartPath: PartPath;
  /** Immediate parent group id. Omitted for the root `p:spTree`. */
  readonly parentGroupId?: string;
  readonly shapeIds: readonly string[];
  readonly groupId: string;
  /** Serialized minimal `p:grpSp` header finalized at edit time. */
  readonly xml: string;
}

export interface PptxSourceModelUngroupShapeEdit {
  readonly kind: "ungroupShape";
  readonly targetPartPath: PartPath;
  readonly groupId: string;
}

export interface PptxSourceModelDeleteShapeEdit {
  readonly kind: "deleteShape";
  readonly handle: SourceHandle;
}

export interface PptxSourceModelReplaceImageEdit {
  readonly kind: "replaceImage";
  readonly handle: SourceHandle;
  /**
   * Single-reference parts are updated in place; shared parts use an isolated copy.
   * Omitted legacy edit records are interpreted as in-place replacements.
   */
  readonly mode?: "inPlace" | "copyOnWrite";
  /** Media part referenced before this edit. Defaults to `mediaPartPath` for legacy records. */
  readonly sourceMediaPartPath?: PartPath;
  /** Relationship referenced by the target picture before this edit. */
  readonly sourceRelationshipId?: RelationshipId;
  /** Media part that contains the replacement bytes after this edit. */
  readonly mediaPartPath: PartPath;
  readonly contentType: string;
  readonly sharedReferenceCount: number;
  /** New relationship assigned to the target picture for copy-on-write edits. */
  readonly replacementRelationshipId?: RelationshipId;
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

export type ThemeColorSlot =
  | "dk1"
  | "lt1"
  | "dk2"
  | "lt2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hlink"
  | "folHlink";

export type ThemeFontSetKind = "major" | "minor";

export interface ThemeFontSetPatch {
  readonly latin?: string;
  readonly eastAsian?: string;
  readonly complexScript?: string;
}

export interface PptxSourceModelUpdateThemeSchemeEdit {
  readonly kind: "updateThemeScheme";
  readonly themePartPath: PartPath;
  readonly colorScheme?: Readonly<Partial<Record<ThemeColorSlot, string>>>;
  readonly fontScheme?: Readonly<Partial<Record<ThemeFontSetKind, ThemeFontSetPatch>>>;
}

export type PptxSourceModelSetBackgroundEdit = {
  readonly kind: "setBackground";
  /** Slide, layout, or master part whose direct `p:bg` is replaced or removed. */
  readonly targetPartPath: PartPath;
} & (
  | {
      /** Omitted only when clearing the direct background. */
      readonly xml?: never;
      readonly relationshipId?: never;
      readonly mediaPartPath?: never;
      readonly contentType?: never;
    }
  | {
      /** Serialized `p:bg` fragment finalized at edit time. */
      readonly xml: string;
      readonly relationshipId: RelationshipId;
      readonly mediaPartPath: PartPath;
      readonly contentType: string;
    }
  | {
      /** Serialized `p:bg` fragment finalized at edit time. */
      readonly xml: string;
      readonly relationshipId?: never;
      readonly mediaPartPath?: never;
      readonly contentType?: never;
    }
);

/** @deprecated Compatibility journal entry produced by `setSlideBackground`. */
export type PptxSourceModelSetSlideBackgroundEdit = {
  readonly kind: "setSlideBackground";
  readonly slidePartPath: PartPath;
  readonly xml: string;
} & (
  | {
      readonly relationshipId: RelationshipId;
      readonly mediaPartPath: PartPath;
      readonly contentType: string;
    }
  | {
      readonly relationshipId?: never;
      readonly mediaPartPath?: never;
      readonly contentType?: never;
    }
);
