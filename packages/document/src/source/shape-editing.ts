/**
 * Existing shape mutation operations for PptxSourceModel.
 *
 * This module owns lookup and mutation of existing shape nodes. New-content shape
 * authoring lives in shape-authoring.ts so validation and XML-finalization changes do
 * not expand the change surface of existing-node editing.
 */

import { editInsertedShape, editTargetsShape, sourceHandlesEqual } from "./edit-descriptors.js";
import type {
  EditableShapeFill,
  EditableShapeOutline,
  PptxSourceModel,
  PptxSourceModelEdit,
  PptxSourceModelShapeOutlineEdit,
  SourceConnector,
  SourceFill,
  SourceHandle,
  SourceOutline,
  SourceShape,
  SourceShapeNode,
} from "./index.js";
import type { UpdateShapeTransformInput } from "./shape-transform.js";

type TransformableShapeNode = Exclude<SourceShapeNode, { readonly kind: "raw" }>;

export type { UpdateShapeTransformInput } from "./shape-transform.js";

type StyleEditableShapeNode = SourceShape | SourceConnector;

export function findShapeNodeBySourceHandle(
  source: PptxSourceModel,
  handle: SourceHandle,
): SourceShapeNode | undefined {
  const matches = [...source.slides, ...source.slideLayouts, ...source.slideMasters].flatMap(
    (target) => findShapeNodesInTree(target.shapes, handle),
  );
  if (matches.length > 1) {
    throw duplicateNodeIdError("findShapeNodeBySourceHandle", handle);
  }
  return matches[0]?.node;
}

export function updateShapeTransform(
  source: PptxSourceModel,
  handle: SourceHandle,
  transform: UpdateShapeTransformInput,
): PptxSourceModel {
  if (handle.nodeId === undefined) {
    throw new Error("updateShapeTransform: shape transform edit requires a node id");
  }

  const target = requireUniqueSlideShapeTarget(source, handle, "updateShapeTransform");
  assertNotAlternateContentTarget(target, "updateShapeTransform");
  if (!hasEditableTransform(target.node)) {
    throw new Error("updateShapeTransform: shape handle does not reference a shape with xfrm");
  }
  if (shapeTransformPositionAndSizeEqual(target.node.transform, transform)) return source;
  const slides = replaceSlideShapeNode(source, handle, (shape) => {
    if (!hasEditableTransform(shape)) return shape;
    return {
      ...shape,
      transform: {
        ...shape.transform,
        offsetX: transform.offsetX,
        offsetY: transform.offsetY,
        width: transform.width,
        height: transform.height,
      },
    };
  });

  return {
    ...source,
    slides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "updateShapeTransform",
        handle,
        offsetX: transform.offsetX,
        offsetY: transform.offsetY,
        width: transform.width,
        height: transform.height,
      },
    ],
  };
}

export function setShapeFill(
  source: PptxSourceModel,
  handle: SourceHandle,
  fill: EditableShapeFill,
): PptxSourceModel {
  assertEditableShapeFill(fill, "setShapeFill");
  if (handle.nodeId === undefined) {
    throw new Error("setShapeFill: shape fill edit requires a node id");
  }

  const target = requireUniqueSlideShapeTarget(source, handle, "setShapeFill");
  assertNotAlternateContentTarget(target, "setShapeFill");
  if (target.node.kind !== "shape") {
    throw new Error("setShapeFill: only sp shapes support fill edits");
  }
  const nextFill = toSourceFill(fill);
  if (sourceFillEqual(target.node.fill, nextFill)) return source;
  const slides = replaceSlideShapeNode(source, handle, (shape) =>
    shape.kind === "shape"
      ? ({
          ...shape,
          fill: nextFill,
        } satisfies SourceShape)
      : shape,
  );

  return {
    ...source,
    slides,
    edits: appendShapeFillEdit(source.edits ?? [], handle, fill),
  };
}

export function setShapeOutline(
  source: PptxSourceModel,
  handle: SourceHandle,
  outline: EditableShapeOutline,
): PptxSourceModel {
  assertEditableShapeOutline(outline, "setShapeOutline");
  if (handle.nodeId === undefined) {
    throw new Error("setShapeOutline: shape outline edit requires a node id");
  }

  const target = requireUniqueSlideShapeTarget(source, handle, "setShapeOutline");
  assertNotAlternateContentTarget(target, "setShapeOutline");
  if (target.node.kind !== "shape" && target.node.kind !== "connector") {
    throw new Error("setShapeOutline: only sp and cxnSp shapes support outline edits");
  }
  const nextOutline = patchSourceOutline(target.node.outline, outline);
  if (sourceOutlineEqual(target.node.outline, nextOutline)) return source;
  const slides = replaceSlideShapeNode(source, handle, (shape) =>
    shape.kind === "shape" || shape.kind === "connector"
      ? ({
          ...shape,
          outline: nextOutline,
        } satisfies StyleEditableShapeNode)
      : shape,
  );

  return {
    ...source,
    slides,
    edits: appendShapeOutlineEdit(source.edits ?? [], handle, outline),
  };
}

export function deleteShape(source: PptxSourceModel, handle: SourceHandle): PptxSourceModel {
  if (handle.nodeId === undefined) {
    throw new Error("deleteShape: shape delete requires a node id");
  }

  const target = requireUniqueSlideShapeTarget(source, handle, "deleteShape");
  if (target.nested) {
    throw new Error("deleteShape: nested group shape deletion is not supported");
  }
  if (target.node.kind !== "shape" && target.node.kind !== "connector") {
    throw new Error("deleteShape: only top-level sp or cxnSp shapes can be deleted");
  }
  assertNotAlternateContentTarget(target, "deleteShape");

  const slides = source.slides.map((slide) => {
    let slideChanged = false;
    const nextShapes = slide.shapes.filter((shape) => {
      if (!sourceHandlesEqual(shape.handle, handle)) return true;
      slideChanged = true;
      return false;
    });
    return slideChanged ? { ...slide, shapes: nextShapes } : slide;
  });

  const referencingConnector = findConnectorReferencingShape(source, handle);
  if (referencingConnector !== undefined) {
    throw new Error(
      `deleteShape: shape is referenced by connector '${referencingConnector.name ?? referencingConnector.nodeId ?? "unknown"}'`,
    );
  }

  const retainedEdits = (source.edits ?? []).filter((edit) => !editTargetsShape(edit, handle));
  const deletedInsertedShape = (source.edits ?? []).some((edit) => {
    const inserted = editInsertedShape(edit);
    return (
      inserted !== undefined &&
      inserted.slidePartPath === handle.partPath &&
      inserted.shapeId === String(handle.nodeId)
    );
  });

  return {
    ...source,
    slides,
    edits: deletedInsertedShape
      ? retainedEdits
      : [
          ...retainedEdits,
          {
            kind: "deleteShape",
            handle,
          },
        ],
  };
}

function shapeTransformPositionAndSizeEqual(
  current: TransformableShapeNode["transform"],
  next: UpdateShapeTransformInput,
): boolean {
  return (
    current?.offsetX === next.offsetX &&
    current.offsetY === next.offsetY &&
    current.width === next.width &&
    current.height === next.height
  );
}

function patchSourceOutline(
  current: SourceOutline | undefined,
  patch: EditableShapeOutline,
): SourceOutline {
  return {
    ...(current ?? {}),
    ...(patch.width !== undefined ? { width: patch.width } : {}),
    ...(patch.fill !== undefined ? { fill: toSourceFill(patch.fill) } : {}),
  };
}

function appendShapeFillEdit(
  edits: readonly PptxSourceModelEdit[],
  handle: SourceHandle,
  fill: EditableShapeFill,
): PptxSourceModelEdit[] {
  const retainedEdits = edits.filter(
    (edit) => edit.kind !== "updateShapeFill" || !sourceHandlesEqual(edit.handle, handle),
  );
  return [...retainedEdits, { kind: "updateShapeFill", handle, fill }];
}

function appendShapeOutlineEdit(
  edits: readonly PptxSourceModelEdit[],
  handle: SourceHandle,
  outline: EditableShapeOutline,
): PptxSourceModelEdit[] {
  let outlineEdit: PptxSourceModelShapeOutlineEdit = {
    kind: "updateShapeOutline",
    handle,
    outline,
  };
  const retainedEdits: PptxSourceModelEdit[] = [];

  for (const edit of edits) {
    if (edit.kind !== "updateShapeOutline" || !sourceHandlesEqual(edit.handle, handle)) {
      retainedEdits.push(edit);
      continue;
    }
    outlineEdit = {
      ...outlineEdit,
      outline: mergeEditableShapeOutline(edit.outline, outlineEdit.outline),
    };
  }

  return [...retainedEdits, outlineEdit];
}

function mergeEditableShapeOutline(
  base: EditableShapeOutline,
  patch: EditableShapeOutline,
): EditableShapeOutline {
  return {
    ...(base.width !== undefined ? { width: base.width } : {}),
    ...(base.fill !== undefined ? { fill: base.fill } : {}),
    ...(patch.width !== undefined ? { width: patch.width } : {}),
    ...(patch.fill !== undefined ? { fill: patch.fill } : {}),
  };
}

function toSourceFill(fill: EditableShapeFill): SourceFill {
  if (fill.kind === "none") return { kind: "none" };
  return {
    kind: "solid",
    color: { kind: "srgb", hex: fill.color.hex },
  };
}

function assertEditableShapeFill(fill: EditableShapeFill, operationName: string): void {
  if (fill.kind === "none") return;
  if (fill.kind !== "solid") {
    throw new Error(`${operationName}: only solid and none fills are supported`);
  }
  if (fill.color.kind !== "srgb") {
    throw new Error(`${operationName}: only srgb solid fill colors are supported`);
  }
  if (!/^[0-9A-Fa-f]{6}$/.test(fill.color.hex)) {
    throw new Error(`${operationName}: srgb fill color must be a 6-digit hex value`);
  }
}

function assertEditableShapeOutline(outline: EditableShapeOutline, operationName: string): void {
  if (outline.width === undefined && outline.fill === undefined) {
    throw new Error(`${operationName}: outline must set width or fill`);
  }
  if (outline.width !== undefined) {
    assertPositiveFiniteEmu(outline.width, operationName, "width");
  }
  if (outline.fill !== undefined) assertEditableShapeFill(outline.fill, operationName);
}

function sourceFillEqual(left: SourceFill | undefined, right: SourceFill | undefined): boolean {
  return stableValueEqual(left ?? {}, right ?? {});
}

function sourceOutlineEqual(
  left: SourceOutline | undefined,
  right: SourceOutline | undefined,
): boolean {
  return stableValueEqual(left ?? {}, right ?? {});
}

function stableValueEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right)) return false;
    if (left.length !== right.length) return false;
    return left.every((value, index) => stableValueEqual(value, right[index]));
  }
  if (isPlainRecord(left) || isPlainRecord(right)) {
    if (!isPlainRecord(left) || !isPlainRecord(right)) return false;
    const leftKeys = Object.keys(left).sort();
    const rightKeys = Object.keys(right).sort();
    if (!stableValueEqual(leftKeys, rightKeys)) return false;
    return leftKeys.every((key) => stableValueEqual(left[key], right[key]));
  }
  return false;
}

function isPlainRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function hasEditableTransform(shape: SourceShapeNode): shape is TransformableShapeNode & {
  readonly transform: NonNullable<TransformableShapeNode["transform"]>;
} {
  return shape.kind !== "raw" && shape.transform !== undefined;
}

function assertPositiveFiniteEmu(value: unknown, operationName: string, fieldName: string): void {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw new Error(`${operationName}: ${fieldName} must be a finite positive EMU value`);
  }
}

interface ShapeNodeMatch {
  readonly node: SourceShapeNode;
  readonly nested: boolean;
  readonly insideAlternateContent: boolean;
}

function findShapeNodesInTree(
  shapes: readonly SourceShapeNode[],
  handle: SourceHandle,
  depth = 0,
  insideAlternateContent = false,
): ShapeNodeMatch[] {
  const matches: ShapeNodeMatch[] = [];
  for (const shape of shapes) {
    const shapeInsideAlternateContent =
      insideAlternateContent || hasAlternateContentWrapperSidecar(shape);
    if (sourceHandlesEqual(shape.handle, handle)) {
      matches.push({
        node: shape,
        nested: depth > 0,
        insideAlternateContent: shapeInsideAlternateContent,
      });
    }
    if (shape.kind === "group") {
      matches.push(
        ...findShapeNodesInTree(shape.children, handle, depth + 1, shapeInsideAlternateContent),
      );
    }
  }
  return matches;
}

function requireUniqueSlideShapeTarget(
  source: PptxSourceModel,
  handle: SourceHandle,
  operationName: string,
): ShapeNodeMatch {
  const matches = source.slides.flatMap((slide) => findShapeNodesInTree(slide.shapes, handle));
  if (matches.length > 1) throw duplicateNodeIdError(operationName, handle);
  const match = matches[0];
  if (match === undefined) {
    throw new Error(`${operationName}: shape handle was not found in PptxSourceModel source`);
  }
  return match;
}

function duplicateNodeIdError(operationName: string, handle: SourceHandle): Error {
  return new Error(
    `${operationName}: duplicate node id '${String(handle.nodeId)}' in drawing part '${handle.partPath}' is not supported`,
  );
}

function assertNotAlternateContentTarget(target: ShapeNodeMatch, operationName: string): void {
  if (target.insideAlternateContent) {
    throw new Error(`${operationName}: shapes inside AlternateContent are not supported`);
  }
}

function replaceSlideShapeNode(
  source: PptxSourceModel,
  handle: SourceHandle,
  replace: (shape: SourceShapeNode) => SourceShapeNode,
): PptxSourceModel["slides"] {
  return source.slides.map((slide) => {
    const shapes = replaceShapeNodeInTree(slide.shapes, handle, replace);
    return shapes === slide.shapes ? slide : { ...slide, shapes };
  });
}

function replaceShapeNodeInTree(
  shapes: readonly SourceShapeNode[],
  handle: SourceHandle,
  replace: (shape: SourceShapeNode) => SourceShapeNode,
): readonly SourceShapeNode[] {
  let changed = false;
  const nextShapes = shapes.map((shape) => {
    if (sourceHandlesEqual(shape.handle, handle)) {
      const next = replace(shape);
      if (next !== shape) changed = true;
      return next;
    }
    if (shape.kind !== "group") return shape;
    const children = replaceShapeNodeInTree(shape.children, handle, replace);
    if (children === shape.children) return shape;
    changed = true;
    return { ...shape, children };
  });
  return changed ? nextShapes : shapes;
}

function hasAlternateContentWrapperSidecar(shape: SourceShapeNode): boolean {
  if (shape.kind === "raw") return false;
  return (
    shape.rawSidecars?.some(
      (sidecar) =>
        sidecar.node.name === "mc:AlternateContent" && sidecar.orderingSlot === undefined,
    ) ?? false
  );
}

function findConnectorReferencingShape(
  source: PptxSourceModel,
  handle: SourceHandle,
): SourceConnector | undefined {
  if (handle.nodeId === undefined) return undefined;
  for (const slide of source.slides) {
    if (slide.partPath !== handle.partPath) continue;
    for (const shape of slide.shapes) {
      if (shape.kind !== "connector") continue;
      if (
        shape.connection?.start?.shapeId === handle.nodeId ||
        shape.connection?.end?.shapeId === handle.nodeId
      ) {
        return shape;
      }
    }
  }
  return undefined;
}
