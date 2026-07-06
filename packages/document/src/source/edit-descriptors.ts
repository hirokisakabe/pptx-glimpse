/**
 * Edit-kind descriptor table.
 *
 * Every `PptxSourceModelEdit` kind has cross-cutting traits that are consumed
 * from several call sites: which part paths it reserves for new-part numbering,
 * which XML part the writer must reserialize (its dirty part), whether it
 * targets a given shape, which part removals invalidate it, which shape id it
 * occupies for numeric shape-id numbering, which `p:sldIdLst` operation it
 * contributes to presentation.xml, and which slide/shape it inserted. This
 * mapped table is the single place those traits are declared: adding a new
 * edit kind to the union fails to compile until the kind declares every trait,
 * instead of silently falling through hand-maintained switch/filter chains in
 * editing.ts and write-pptx.ts. The writer keeps exactly one exhaustive apply
 * switch for XML patching; everything else about an edit kind belongs here.
 */

import type { PartPath, RelationshipId, SourceHandle } from "./handles.js";
import type { PptxSourceModelEdit } from "./pptx-source-model.js";

/** Declarative `p:sldIdLst` operation an edit contributes to presentation.xml. */
export type SlideTopologyOperation =
  | {
      readonly kind: "appendSlide";
      readonly newRelationshipId: RelationshipId;
      readonly newSlideNumericId: number;
    }
  | {
      readonly kind: "insertSlideAfter";
      readonly sourceRelationshipId: RelationshipId;
      readonly newRelationshipId: RelationshipId;
      readonly newSlideNumericId: number;
    }
  | {
      readonly kind: "removeSlide";
      readonly relationshipId: RelationshipId;
    }
  | {
      readonly kind: "moveSlide";
      readonly relationshipId: RelationshipId;
      readonly toIndex: number;
    };

/** Shape inserted by an edit, identified by its slide part and shape id. */
interface InsertedShapeRef {
  readonly slidePartPath: PartPath;
  readonly shapeId: string;
}

type EditOfKind<K extends PptxSourceModelEdit["kind"]> = Extract<
  PptxSourceModelEdit,
  { readonly kind: K }
>;

interface EditKindDescriptor<E extends PptxSourceModelEdit> {
  /** Part paths the edit names; kept unavailable when numbering new parts. */
  readonly reservedPartPaths: (edit: E) => readonly PartPath[];
  /** XML part the writer must reserialize to apply the edit, if any. */
  readonly dirtyPartPath: (edit: E) => PartPath | undefined;
  /** Whether the edit targets the shape identified by the handle. */
  readonly targetsShape: (edit: E, shapeHandle: SourceHandle) => boolean;
  /** Part paths whose removal from the package invalidates the edit. */
  readonly invalidatingPartPaths: (edit: E) => readonly PartPath[];
  /** Shape id the edit occupies on the slide part for numeric id numbering. */
  readonly reservedShapeId: (edit: E, slidePartPath: PartPath) => string | undefined;
  /** `p:sldIdLst` operation the writer applies to presentation.xml, if any. */
  readonly slideTopologyOperation: (edit: E) => SlideTopologyOperation | undefined;
  /** Slide part the edit inserted; deleting that slide cancels the edit. */
  readonly insertedSlidePartPath: (edit: E) => PartPath | undefined;
  /** Shape the edit inserted; deleting that shape cancels the edit. */
  readonly insertedShape: (edit: E) => InsertedShapeRef | undefined;
}

const EDIT_KIND_DESCRIPTORS: {
  readonly [K in PptxSourceModelEdit["kind"]]: EditKindDescriptor<EditOfKind<K>>;
} = {
  replaceTextRunPlainText: {
    reservedPartPaths: () => [],
    dirtyPartPath: (edit) => edit.handle.partPath,
    targetsShape: (edit, shapeHandle) =>
      edit.handle.partPath === shapeHandle.partPath &&
      textRunShapeId(edit.handle) === shapeHandle.nodeId,
    invalidatingPartPaths: (edit) => [edit.handle.partPath],
    reservedShapeId: () => undefined,
    slideTopologyOperation: () => undefined,
    insertedSlidePartPath: () => undefined,
    insertedShape: () => undefined,
  },
  updateTextRunProperties: {
    reservedPartPaths: () => [],
    dirtyPartPath: (edit) => edit.handle.partPath,
    targetsShape: (edit, shapeHandle) =>
      edit.handle.partPath === shapeHandle.partPath &&
      textRunShapeId(edit.handle) === shapeHandle.nodeId,
    invalidatingPartPaths: (edit) => [edit.handle.partPath],
    reservedShapeId: () => undefined,
    slideTopologyOperation: () => undefined,
    insertedSlidePartPath: () => undefined,
    insertedShape: () => undefined,
  },
  updateParagraphProperties: {
    reservedPartPaths: () => [],
    dirtyPartPath: (edit) => edit.handle.partPath,
    targetsShape: (edit, shapeHandle) =>
      edit.handle.partPath === shapeHandle.partPath &&
      paragraphShapeId(edit.handle) === shapeHandle.nodeId,
    invalidatingPartPaths: (edit) => [edit.handle.partPath],
    reservedShapeId: () => undefined,
    slideTopologyOperation: () => undefined,
    insertedSlidePartPath: () => undefined,
    insertedShape: () => undefined,
  },
  replaceParagraphPlainText: {
    reservedPartPaths: () => [],
    dirtyPartPath: (edit) => edit.handle.partPath,
    targetsShape: (edit, shapeHandle) =>
      edit.handle.partPath === shapeHandle.partPath &&
      paragraphShapeId(edit.handle) === shapeHandle.nodeId,
    invalidatingPartPaths: (edit) => [edit.handle.partPath],
    reservedShapeId: () => undefined,
    slideTopologyOperation: () => undefined,
    insertedSlidePartPath: () => undefined,
    insertedShape: () => undefined,
  },
  updateShapeTransform: {
    reservedPartPaths: () => [],
    dirtyPartPath: (edit) => edit.handle.partPath,
    targetsShape: (edit, shapeHandle) => sourceHandlesEqual(edit.handle, shapeHandle),
    invalidatingPartPaths: (edit) => [edit.handle.partPath],
    reservedShapeId: () => undefined,
    slideTopologyOperation: () => undefined,
    insertedSlidePartPath: () => undefined,
    insertedShape: () => undefined,
  },
  updateShapeFill: {
    reservedPartPaths: () => [],
    dirtyPartPath: (edit) => edit.handle.partPath,
    targetsShape: (edit, shapeHandle) => sourceHandlesEqual(edit.handle, shapeHandle),
    invalidatingPartPaths: (edit) => [edit.handle.partPath],
    reservedShapeId: () => undefined,
    slideTopologyOperation: () => undefined,
    insertedSlidePartPath: () => undefined,
    insertedShape: () => undefined,
  },
  updateShapeOutline: {
    reservedPartPaths: () => [],
    dirtyPartPath: (edit) => edit.handle.partPath,
    targetsShape: (edit, shapeHandle) => sourceHandlesEqual(edit.handle, shapeHandle),
    invalidatingPartPaths: (edit) => [edit.handle.partPath],
    reservedShapeId: () => undefined,
    slideTopologyOperation: () => undefined,
    insertedSlidePartPath: () => undefined,
    insertedShape: () => undefined,
  },
  addTextBox: {
    reservedPartPaths: (edit) => [edit.slidePartPath],
    dirtyPartPath: (edit) => edit.slidePartPath,
    targetsShape: (edit, shapeHandle) =>
      edit.slidePartPath === shapeHandle.partPath && edit.shapeId === String(shapeHandle.nodeId),
    invalidatingPartPaths: (edit) => [edit.slidePartPath],
    reservedShapeId: (edit, slidePartPath) =>
      edit.slidePartPath === slidePartPath ? edit.shapeId : undefined,
    slideTopologyOperation: () => undefined,
    insertedSlidePartPath: () => undefined,
    insertedShape: (edit) => ({ slidePartPath: edit.slidePartPath, shapeId: edit.shapeId }),
  },
  addConnector: {
    reservedPartPaths: (edit) => [edit.slidePartPath],
    dirtyPartPath: (edit) => edit.slidePartPath,
    targetsShape: (edit, shapeHandle) =>
      edit.slidePartPath === shapeHandle.partPath && edit.shapeId === String(shapeHandle.nodeId),
    invalidatingPartPaths: (edit) => [edit.slidePartPath],
    reservedShapeId: (edit, slidePartPath) =>
      edit.slidePartPath === slidePartPath ? edit.shapeId : undefined,
    slideTopologyOperation: () => undefined,
    insertedSlidePartPath: () => undefined,
    insertedShape: (edit) => ({ slidePartPath: edit.slidePartPath, shapeId: edit.shapeId }),
  },
  deleteShape: {
    reservedPartPaths: () => [],
    dirtyPartPath: (edit) => edit.handle.partPath,
    targetsShape: (edit, shapeHandle) => sourceHandlesEqual(edit.handle, shapeHandle),
    invalidatingPartPaths: (edit) => [edit.handle.partPath],
    reservedShapeId: (edit, slidePartPath) =>
      edit.handle.partPath === slidePartPath ? edit.handle.nodeId : undefined,
    slideTopologyOperation: () => undefined,
    insertedSlidePartPath: () => undefined,
    insertedShape: () => undefined,
  },
  replaceImage: {
    reservedPartPaths: () => [],
    dirtyPartPath: () => undefined,
    targetsShape: (edit, shapeHandle) => sourceHandlesEqual(edit.handle, shapeHandle),
    invalidatingPartPaths: (edit) => [edit.handle.partPath],
    reservedShapeId: () => undefined,
    slideTopologyOperation: () => undefined,
    insertedSlidePartPath: () => undefined,
    insertedShape: () => undefined,
  },
  addEmptySlideFromLayout: {
    reservedPartPaths: (edit) => [edit.layoutPartPath, edit.newSlidePartPath],
    dirtyPartPath: () => undefined,
    targetsShape: () => false,
    invalidatingPartPaths: (edit) => [edit.newSlidePartPath],
    reservedShapeId: () => undefined,
    slideTopologyOperation: (edit) => ({
      kind: "appendSlide",
      newRelationshipId: edit.newRelationshipId,
      newSlideNumericId: edit.newSlideNumericId,
    }),
    insertedSlidePartPath: (edit) => edit.newSlidePartPath,
    insertedShape: () => undefined,
  },
  duplicateSlide: {
    reservedPartPaths: (edit) => [edit.sourceSlidePartPath, edit.newSlidePartPath],
    dirtyPartPath: () => undefined,
    targetsShape: () => false,
    invalidatingPartPaths: (edit) => [edit.newSlidePartPath],
    reservedShapeId: () => undefined,
    slideTopologyOperation: (edit) => ({
      kind: "insertSlideAfter",
      sourceRelationshipId: edit.sourceRelationshipId,
      newRelationshipId: edit.newRelationshipId,
      newSlideNumericId: edit.newSlideNumericId,
    }),
    insertedSlidePartPath: (edit) => edit.newSlidePartPath,
    insertedShape: () => undefined,
  },
  moveSlide: {
    reservedPartPaths: (edit) => [edit.slidePartPath],
    dirtyPartPath: () => undefined,
    targetsShape: () => false,
    invalidatingPartPaths: (edit) => [edit.slidePartPath],
    reservedShapeId: () => undefined,
    slideTopologyOperation: (edit) => ({
      kind: "moveSlide",
      relationshipId: edit.relationshipId,
      toIndex: edit.toIndex,
    }),
    insertedSlidePartPath: () => undefined,
    insertedShape: () => undefined,
  },
  deleteSlide: {
    reservedPartPaths: (edit) => [edit.slidePartPath],
    dirtyPartPath: () => undefined,
    targetsShape: () => false,
    invalidatingPartPaths: (edit) => [edit.slidePartPath],
    reservedShapeId: () => undefined,
    slideTopologyOperation: (edit) => ({
      kind: "removeSlide",
      relationshipId: edit.relationshipId,
    }),
    insertedSlidePartPath: () => undefined,
    insertedShape: () => undefined,
  },
};

function descriptorFor(edit: PptxSourceModelEdit): EditKindDescriptor<PptxSourceModelEdit> {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion -- The mapped table guarantees the entry at `edit.kind` accepts exactly that edit shape; TypeScript cannot correlate the indexed access with the discriminant, so this single boundary widens the descriptor parameter to the full edit union.
  return EDIT_KIND_DESCRIPTORS[edit.kind] as EditKindDescriptor<PptxSourceModelEdit>;
}

export function editReservedPartPaths(edit: PptxSourceModelEdit): readonly PartPath[] {
  return descriptorFor(edit).reservedPartPaths(edit);
}

export function editDirtyPartPath(edit: PptxSourceModelEdit): PartPath | undefined {
  return descriptorFor(edit).dirtyPartPath(edit);
}

export function editTargetsShape(edit: PptxSourceModelEdit, shapeHandle: SourceHandle): boolean {
  return descriptorFor(edit).targetsShape(edit, shapeHandle);
}

export function editInvalidatingPartPaths(edit: PptxSourceModelEdit): readonly PartPath[] {
  return descriptorFor(edit).invalidatingPartPaths(edit);
}

export function editReservedShapeId(
  edit: PptxSourceModelEdit,
  slidePartPath: PartPath,
): string | undefined {
  return descriptorFor(edit).reservedShapeId(edit, slidePartPath);
}

export function editSlideTopologyOperation(
  edit: PptxSourceModelEdit,
): SlideTopologyOperation | undefined {
  return descriptorFor(edit).slideTopologyOperation(edit);
}

export function editInsertedSlidePartPath(edit: PptxSourceModelEdit): PartPath | undefined {
  return descriptorFor(edit).insertedSlidePartPath(edit);
}

export function editInsertedShape(edit: PptxSourceModelEdit): InsertedShapeRef | undefined {
  return descriptorFor(edit).insertedShape(edit);
}

export function sourceHandlesEqual(left: SourceHandle | undefined, right: SourceHandle): boolean {
  if (left === undefined) return false;
  return (
    left.partPath === right.partPath &&
    left.nodeId === right.nodeId &&
    left.relationshipId === right.relationshipId &&
    left.orderingSlot === right.orderingSlot
  );
}

function textRunShapeId(handle: SourceHandle): string | undefined {
  return /^text:shape:([^:]+):p:\d+:r:\d+$/.exec(String(handle.nodeId ?? ""))?.[1];
}

function paragraphShapeId(handle: SourceHandle): string | undefined {
  return /^text:shape:([^:]+):p:\d+$/.exec(String(handle.nodeId ?? ""))?.[1];
}
