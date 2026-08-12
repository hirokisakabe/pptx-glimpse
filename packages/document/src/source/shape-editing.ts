/**
 * Existing shape mutation operations for PptxSourceModel.
 *
 * This module owns lookup and mutation of existing shape nodes. New-content shape
 * authoring lives in shape-authoring.ts so validation and XML-finalization changes do
 * not expand the change surface of existing-node editing.
 */

import { getAttr, getChild, localName, parseXml, type XmlNode } from "../reader/xml.js";
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
import { removePackageParts, removePartRelationship } from "./package-graph-mutations.js";
import { resolveInternalRelationshipTarget } from "./package-paths.js";
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

  const target = requireUniqueDrawingShapeTarget(source, handle, "updateShapeTransform");
  assertNotAlternateContentTarget(target, "updateShapeTransform");
  if (!hasEditableTransform(target.node)) {
    throw new Error("updateShapeTransform: shape handle does not reference a shape with xfrm");
  }
  if (shapeTransformPositionAndSizeEqual(target.node.transform, transform)) return source;
  const drawingParts = replaceDrawingShapeNode(source, handle, (shape) => {
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
    ...drawingParts,
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

  const target = requireUniqueDrawingShapeTarget(source, handle, "setShapeFill");
  assertNotAlternateContentTarget(target, "setShapeFill");
  if (target.node.kind !== "shape") {
    throw new Error("setShapeFill: only sp shapes support fill edits");
  }
  const nextFill = toSourceFill(fill);
  if (sourceFillEqual(target.node.fill, nextFill)) return source;
  const drawingParts = replaceDrawingShapeNode(source, handle, (shape) =>
    shape.kind === "shape"
      ? ({
          ...shape,
          fill: nextFill,
        } satisfies SourceShape)
      : shape,
  );

  return {
    ...source,
    ...drawingParts,
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

  const target = requireUniqueDrawingShapeTarget(source, handle, "setShapeOutline");
  assertNotAlternateContentTarget(target, "setShapeOutline");
  if (target.node.kind !== "shape" && target.node.kind !== "connector") {
    throw new Error("setShapeOutline: only sp and cxnSp shapes support outline edits");
  }
  const nextOutline = patchSourceOutline(target.node.outline, outline);
  if (sourceOutlineEqual(target.node.outline, nextOutline)) return source;
  const drawingParts = replaceDrawingShapeNode(source, handle, (shape) =>
    shape.kind === "shape" || shape.kind === "connector"
      ? ({
          ...shape,
          outline: nextOutline,
        } satisfies StyleEditableShapeNode)
      : shape,
  );

  return {
    ...source,
    ...drawingParts,
    edits: appendShapeOutlineEdit(source.edits ?? [], handle, outline),
  };
}

export function deleteShape(source: PptxSourceModel, handle: SourceHandle): PptxSourceModel {
  if (handle.nodeId === undefined) {
    throw new Error("deleteShape: shape delete requires a node id");
  }

  const target = requireUniqueSlideShapeTarget(source, handle, "deleteShape");
  if (
    target.node.kind !== "shape" &&
    target.node.kind !== "connector" &&
    target.node.kind !== "image" &&
    target.node.kind !== "table" &&
    target.node.kind !== "chart" &&
    target.node.kind !== "group"
  ) {
    throw new Error(
      "deleteShape: only sp, cxnSp, pic, native table/chart graphicFrame, or grpSp drawings can be deleted",
    );
  }
  assertNotAlternateContentTarget(target, "deleteShape");

  const cleanupPlan = planDrawingDeletionCleanup(source, handle, target.node);

  const slides = source.slides.map((slide) => {
    if (slide.partPath !== handle.partPath) return slide;
    const shapes = deleteShapeNodeFromTree(slide.shapes, handle);
    return shapes === slide.shapes ? slide : { ...slide, shapes };
  });

  const referencingConnector = findConnectorReferencingDeletedSubtree(source, handle, target.node);
  if (referencingConnector !== undefined) {
    throw new Error(
      `deleteShape: shape is referenced by connector '${referencingConnector.name ?? referencingConnector.nodeId ?? "unknown"}'`,
    );
  }

  const retainedEdits = (source.edits ?? []).filter(
    (edit) =>
      (target.nested && isDrawingTopologyRequiredForNestedDelete(edit, handle)) ||
      isPendingGroupCreationForHandle(edit, handle) ||
      !editTargetsShape(edit, handle),
  );
  const deletedInsertedShape = isDirectlyAddedDrawing(source.edits ?? [], handle);

  return {
    ...source,
    slides,
    packageGraph: cleanupPlan,
    edits:
      deletedInsertedShape && !target.nested
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

function deleteShapeNodeFromTree(
  shapes: readonly SourceShapeNode[],
  handle: SourceHandle,
): readonly SourceShapeNode[] {
  let changed = false;
  const nextShapes = shapes.flatMap((shape): readonly SourceShapeNode[] => {
    if (sourceHandlesEqual(shape.handle, handle)) {
      changed = true;
      return [];
    }
    if (shape.kind !== "group") return [shape];
    const children = deleteShapeNodeFromTree(shape.children, handle);
    if (children === shape.children) return [shape];
    changed = true;
    return [{ ...shape, children }];
  });
  return changed ? nextShapes : shapes;
}

function isDrawingTopologyRequiredForNestedDelete(
  edit: PptxSourceModelEdit,
  handle: SourceHandle,
): boolean {
  if (edit.kind === "groupShapes") {
    return (
      edit.targetPartPath === handle.partPath &&
      (edit.groupId === String(handle.nodeId) || edit.shapeIds.includes(String(handle.nodeId)))
    );
  }
  const inserted = editInsertedShape(edit);
  return inserted?.slidePartPath === handle.partPath && inserted.shapeId === String(handle.nodeId);
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

export function requireUniqueSlideShapeTarget(
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

function requireUniqueDrawingShapeTarget(
  source: PptxSourceModel,
  handle: SourceHandle,
  operationName: string,
): ShapeNodeMatch {
  const matches = [...source.slides, ...source.slideLayouts, ...source.slideMasters].flatMap(
    (target) => findShapeNodesInTree(target.shapes, handle),
  );
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

export function assertNotAlternateContentTarget(
  target: ShapeNodeMatch,
  operationName: string,
): void {
  if (target.insideAlternateContent) {
    throw new Error(`${operationName}: shapes inside AlternateContent are not supported`);
  }
}

export function replaceSlideShapeNode(
  source: PptxSourceModel,
  handle: SourceHandle,
  replace: (shape: SourceShapeNode) => SourceShapeNode,
): PptxSourceModel["slides"] {
  return source.slides.map((slide) => {
    const shapes = replaceShapeNodeInTree(slide.shapes, handle, replace);
    return shapes === slide.shapes ? slide : { ...slide, shapes };
  });
}

function replaceDrawingShapeNode(
  source: PptxSourceModel,
  handle: SourceHandle,
  replace: (shape: SourceShapeNode) => SourceShapeNode,
): Pick<PptxSourceModel, "slides" | "slideLayouts" | "slideMasters"> {
  const replaceTargets = <
    T extends {
      readonly partPath: SourceHandle["partPath"];
      readonly shapes: readonly SourceShapeNode[];
    },
  >(
    targets: readonly T[],
  ): readonly T[] =>
    targets.map((target) => {
      if (target.partPath !== handle.partPath) return target;
      const shapes = replaceShapeNodeInTree(target.shapes, handle, replace);
      return shapes === target.shapes ? target : { ...target, shapes };
    });

  return {
    slides: replaceTargets(source.slides),
    slideLayouts: replaceTargets(source.slideLayouts),
    slideMasters: replaceTargets(source.slideMasters),
  };
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

function findConnectorReferencingDeletedSubtree(
  source: PptxSourceModel,
  handle: SourceHandle,
  target: SourceShapeNode,
): SourceConnector | undefined {
  const deletedIds = collectSourceShapeIds(target);
  for (const slide of source.slides) {
    if (slide.partPath !== handle.partPath) continue;
    for (const shape of flattenSourceShapeTree(slide.shapes)) {
      if (shape.kind !== "connector" || deletedIds.has(String(shape.nodeId))) continue;
      if (
        deletedIds.has(String(shape.connection?.start?.shapeId)) ||
        deletedIds.has(String(shape.connection?.end?.shapeId))
      ) {
        return shape;
      }
    }
  }
  return undefined;
}

function planDrawingDeletionCleanup(
  source: PptxSourceModel,
  handle: SourceHandle,
  target: SourceShapeNode,
): PptxSourceModel["packageGraph"] {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === handle.partPath);
  if (rawPart?.kind !== "binary") {
    throw new Error(
      `deleteShape: drawing part '${handle.partPath}' has no preserved XML for reference validation`,
    );
  }

  let root: XmlNode;
  try {
    root = parseXml(new TextDecoder().decode(rawPart.bytes));
  } catch (cause) {
    throw new Error(
      `deleteShape: drawing part '${handle.partPath}' could not be parsed for reference validation`,
      { cause },
    );
  }
  const directAddition = isDirectlyAddedDrawing(source.edits ?? [], handle);
  const pendingGroupCreation = (source.edits ?? []).some((edit) =>
    isPendingGroupCreationForHandle(edit, handle),
  );
  const typedRelationshipIds = collectSourceRelationshipIds(target);
  for (const node of flattenSourceShapeTree([target])) {
    if (node.handle === undefined) continue;
    for (const relationshipId of collectPendingAddedRelationshipIds(
      source.edits ?? [],
      node.handle,
    )) {
      typedRelationshipIds.add(relationshipId);
    }
  }
  const spTree = getChild(getChild(getChild(root, "sld"), "cSld"), "spTree");
  if (spTree === undefined) {
    if (
      directAddition ||
      (typedRelationshipIds.size === 0 &&
        target.kind !== "group" &&
        target.kind !== "image" &&
        target.kind !== "chart" &&
        (!("rawSidecars" in target) || target.rawSidecars === undefined))
    ) {
      return cleanupDirectAdditionRelationships(source, handle, target, root);
    }
    throw new Error(`deleteShape: drawing part '${handle.partPath}' has no shape tree`);
  }

  const nodeId = String(handle.nodeId);
  const matches = findXmlDrawingMatches(spTree, nodeId);
  if (matches.some((match) => match.insideAlternateContent)) {
    throw new Error("deleteShape: shapes inside AlternateContent are not supported");
  }
  const xmlTarget = matches.find((match) => !match.insideAlternateContent)?.node;
  if (matches.filter((match) => !match.insideAlternateContent).length > 1) {
    throw duplicateNodeIdError("deleteShape", handle);
  }
  if (xmlTarget === undefined && !directAddition && !pendingGroupCreation) {
    throw new Error(
      `deleteShape: drawing '${nodeId}' was not found in preserved owner XML for validation`,
    );
  }

  const deletedIds = collectSourceShapeIds(target);
  const currentXmlTargets = new Set<XmlNode>();
  if (xmlTarget !== undefined) currentXmlTargets.add(xmlTarget);
  if (xmlTarget === undefined && pendingGroupCreation && target.kind === "group") {
    for (const child of flattenSourceShapeTree(target.children)) {
      if (child.nodeId === undefined) continue;
      const childMatches = findXmlDrawingMatches(spTree, String(child.nodeId)).filter(
        (match) => !match.insideAlternateContent,
      );
      if (childMatches.length > 1)
        throw duplicateNodeIdError("deleteShape", child.handle ?? handle);
      if (childMatches[0] !== undefined) currentXmlTargets.add(childMatches[0].node);
    }
  }
  for (const deletedXmlTarget of currentXmlTargets) {
    collectXmlDrawingIds(deletedXmlTarget, deletedIds);
  }
  const skippedXmlSubtrees = new Set([
    ...collectPreviouslyDeletedXmlSubtrees(source.edits ?? [], handle, spTree),
    ...currentXmlTargets,
  ]);
  const rawReferencingConnector = findXmlConnectorReferencingIds(
    spTree,
    skippedXmlSubtrees,
    deletedIds,
  );
  if (rawReferencingConnector !== undefined) {
    throw new Error(`deleteShape: shape is referenced by connector '${rawReferencingConnector}'`);
  }

  const candidateRelationshipIds = typedRelationshipIds;
  for (const deletedXmlTarget of currentXmlTargets) {
    collectRelationshipAttributeValues(
      deletedXmlTarget,
      candidateRelationshipIds,
      new Set(),
      findNamespaceBindingsForNode(root, deletedXmlTarget),
    );
  }
  if (candidateRelationshipIds.size === 0) return source.packageGraph;

  const remainingRelationshipIds = collectEffectiveRemainingRelationshipIds(
    source.edits ?? [],
    handle,
    root,
    spTree,
    skippedXmlSubtrees,
  );
  const ownerRelationships = source.packageGraph.relationships.find(
    (group) => group.sourcePartPath === handle.partPath,
  );
  if (ownerRelationships === undefined) return source.packageGraph;

  let graph = source.packageGraph;
  const reachableTargets = [];
  for (const relationship of ownerRelationships.relationships) {
    if (
      !candidateRelationshipIds.has(String(relationship.id)) ||
      remainingRelationshipIds.has(String(relationship.id))
    ) {
      continue;
    }
    const targetPartPath = resolveInternalRelationshipTarget(handle.partPath, relationship);
    graph = removePartRelationship(graph, handle.partPath, relationship.id);
    if (targetPartPath !== undefined) reachableTargets.push(targetPartPath);
  }
  return removeUnreferencedReachableParts(graph, reachableTargets);
}

function collectEffectiveRemainingRelationshipIds(
  edits: readonly PptxSourceModelEdit[],
  ownerHandle: SourceHandle,
  root: XmlNode,
  shapeTree: XmlNode,
  skippedSubtrees: ReadonlySet<XmlNode>,
): ReadonlySet<string> {
  const pendingReplacements = edits.flatMap((edit) => {
    if (
      edit.kind !== "replaceImage" ||
      edit.mode !== "copyOnWrite" ||
      edit.handle.partPath !== ownerHandle.partPath ||
      edit.handle.nodeId === undefined ||
      edit.replacementRelationshipId === undefined
    ) {
      return [];
    }
    const matches = findXmlDrawingMatches(shapeTree, String(edit.handle.nodeId));
    if (matches.some((match) => match.insideAlternateContent)) {
      throw new Error("deleteShape: pending image replacement is inside AlternateContent");
    }
    const supportedMatches = matches.filter((match) => !match.insideAlternateContent);
    if (supportedMatches.length > 1) {
      throw duplicateNodeIdError("deleteShape", edit.handle);
    }
    if (supportedMatches[0] === undefined) {
      throw new Error(
        `deleteShape: pending image replacement '${String(edit.handle.nodeId)}' was not found in preserved owner XML`,
      );
    }
    return [{ edit, node: supportedMatches[0].node }];
  });
  const replacementSubtrees = new Set(pendingReplacements.map(({ node }) => node));
  const excludedSubtrees = new Set([...skippedSubtrees, ...replacementSubtrees]);
  const output = new Set<string>();
  collectRelationshipAttributeValues(root, output, excludedSubtrees);

  for (const { edit, node } of pendingReplacements) {
    if (isNodeWithinAnySubtree(node, skippedSubtrees)) continue;
    const effectiveIds = new Set<string>();
    collectRelationshipAttributeValues(
      node,
      effectiveIds,
      new Set(),
      findNamespaceBindingsForNode(root, node),
    );
    const sourceRelationshipId = edit.sourceRelationshipId ?? edit.handle.relationshipId;
    if (sourceRelationshipId !== undefined) effectiveIds.delete(String(sourceRelationshipId));
    effectiveIds.add(String(edit.replacementRelationshipId));
    for (const relationshipId of effectiveIds) output.add(relationshipId);
  }
  return output;
}

function collectPreviouslyDeletedXmlSubtrees(
  edits: readonly PptxSourceModelEdit[],
  ownerHandle: SourceHandle,
  shapeTree: XmlNode,
): ReadonlySet<XmlNode> {
  const deletedIds = new Set<string>();
  const pendingGroupChildren = new Map<string, readonly string[]>();
  for (const edit of edits) {
    if (edit.kind === "deleteShape" && edit.handle.partPath === ownerHandle.partPath) {
      if (edit.handle.nodeId !== undefined) deletedIds.add(String(edit.handle.nodeId));
    }
    if (edit.kind === "groupShapes" && edit.targetPartPath === ownerHandle.partPath) {
      pendingGroupChildren.set(edit.groupId, edit.shapeIds);
    }
  }

  const pending = [...deletedIds];
  while (pending.length > 0) {
    const groupId = pending.shift();
    if (groupId === undefined) continue;
    for (const childId of pendingGroupChildren.get(groupId) ?? []) {
      if (deletedIds.has(childId)) continue;
      deletedIds.add(childId);
      pending.push(childId);
    }
  }

  const subtrees = new Set<XmlNode>();
  for (const nodeId of deletedIds) {
    const matches = findXmlDrawingMatches(shapeTree, nodeId);
    if (matches.some((match) => match.insideAlternateContent)) {
      throw new Error("deleteShape: prior deleted shape is inside AlternateContent");
    }
    const supportedMatches = matches.filter((match) => !match.insideAlternateContent);
    if (supportedMatches.length > 1) {
      throw new Error(
        `deleteShape: duplicate node id '${nodeId}' in the target drawing part is not supported`,
      );
    }
    if (supportedMatches[0] !== undefined) subtrees.add(supportedMatches[0].node);
  }
  return subtrees;
}

function cleanupDirectAdditionRelationships(
  source: PptxSourceModel,
  handle: SourceHandle,
  target: SourceShapeNode,
  ownerRoot: XmlNode,
): PptxSourceModel["packageGraph"] {
  const candidateRelationshipIds = collectSourceRelationshipIds(target);
  for (const relationshipId of collectPendingAddedRelationshipIds(source.edits ?? [], handle)) {
    candidateRelationshipIds.add(relationshipId);
  }
  if (candidateRelationshipIds.size === 0) return source.packageGraph;
  const retainedRelationshipIds = new Set<string>();
  collectRelationshipAttributeValues(ownerRoot, retainedRelationshipIds);
  const ownerRelationships = source.packageGraph.relationships.find(
    (group) => group.sourcePartPath === handle.partPath,
  );
  if (ownerRelationships === undefined) return source.packageGraph;
  let graph = source.packageGraph;
  const reachableTargets = [];
  for (const relationship of ownerRelationships.relationships) {
    if (
      !candidateRelationshipIds.has(String(relationship.id)) ||
      retainedRelationshipIds.has(String(relationship.id))
    ) {
      continue;
    }
    const targetPartPath = resolveInternalRelationshipTarget(handle.partPath, relationship);
    graph = removePartRelationship(graph, handle.partPath, relationship.id);
    if (targetPartPath !== undefined) reachableTargets.push(targetPartPath);
  }
  return removeUnreferencedReachableParts(graph, reachableTargets);
}

interface XmlDrawingMatch {
  readonly node: XmlNode;
  readonly insideAlternateContent: boolean;
}

function findXmlDrawingMatches(root: XmlNode, nodeId: string): XmlDrawingMatch[] {
  const matches: XmlDrawingMatch[] = [];
  const visit = (node: XmlNode, insideAlternateContent: boolean): void => {
    for (const [key, value] of Object.entries(node)) {
      if (key.startsWith("@_")) continue;
      const local = localName(key);
      const nextInside = insideAlternateContent || local === "AlternateContent";
      for (const child of xmlNodes(value)) {
        if (isDrawingElement(local) && xmlDrawingId(child) === nodeId) {
          matches.push({ node: child, insideAlternateContent: nextInside });
        }
        visit(child, nextInside);
      }
    }
  };
  visit(root, false);
  return matches;
}

function findXmlConnectorReferencingIds(
  root: XmlNode,
  skippedSubtrees: ReadonlySet<XmlNode>,
  deletedIds: ReadonlySet<string>,
): string | undefined {
  let result: string | undefined;
  const visit = (node: XmlNode): void => {
    if (skippedSubtrees.has(node) || result !== undefined) return;
    for (const [key, value] of Object.entries(node)) {
      if (key.startsWith("@_")) continue;
      const local = localName(key);
      for (const child of xmlNodes(value)) {
        if (skippedSubtrees.has(child)) continue;
        if (local === "cxnSp") {
          const properties = getChild(getChild(child, "nvCxnSpPr"), "cNvCxnSpPr");
          const startId = getAttr(getChild(properties, "stCxn"), "id");
          const endId = getAttr(getChild(properties, "endCxn"), "id");
          if (
            (startId !== undefined && deletedIds.has(startId)) ||
            (endId !== undefined && deletedIds.has(endId))
          ) {
            result =
              getAttr(getChild(getChild(child, "nvCxnSpPr"), "cNvPr"), "name") ??
              xmlDrawingId(child) ??
              "unknown";
            return;
          }
        }
        visit(child);
      }
    }
  };
  visit(root);
  return result;
}

const RELATIONSHIP_NAMESPACE_URIS = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);

type NamespaceBindings = ReadonlyMap<string, string>;

function collectRelationshipAttributeValues(
  root: XmlNode,
  output: Set<string>,
  skippedSubtrees: ReadonlySet<XmlNode> = new Set(),
  inheritedBindings: NamespaceBindings = new Map(),
): void {
  if (skippedSubtrees.has(root)) return;
  const bindings = namespaceBindingsForElement(root, inheritedBindings);
  for (const [key, value] of Object.entries(root)) {
    if (key.startsWith("@_")) {
      const qualifiedName = key.slice(2);
      const colon = qualifiedName.indexOf(":");
      if (
        colon > 0 &&
        RELATIONSHIP_NAMESPACE_URIS.has(bindings.get(qualifiedName.slice(0, colon)) ?? "")
      ) {
        output.add(String(value));
      }
      continue;
    }
    for (const child of xmlNodes(value)) {
      collectRelationshipAttributeValues(child, output, skippedSubtrees, bindings);
    }
  }
}

function namespaceBindingsForElement(
  node: XmlNode,
  inheritedBindings: NamespaceBindings,
): NamespaceBindings {
  let bindings: Map<string, string> | undefined;
  for (const [key, value] of Object.entries(node)) {
    if (!key.startsWith("@_xmlns:")) continue;
    bindings ??= new Map(inheritedBindings);
    bindings.set(key.slice("@_xmlns:".length), String(value));
  }
  return bindings ?? inheritedBindings;
}

function findNamespaceBindingsForNode(root: XmlNode, target: XmlNode): NamespaceBindings {
  const visit = (
    node: XmlNode,
    inheritedBindings: NamespaceBindings,
  ): NamespaceBindings | undefined => {
    const bindings = namespaceBindingsForElement(node, inheritedBindings);
    if (node === target) return bindings;
    for (const [key, value] of Object.entries(node)) {
      if (key.startsWith("@_")) continue;
      for (const child of xmlNodes(value)) {
        const result = visit(child, bindings);
        if (result !== undefined) return result;
      }
    }
    return undefined;
  };
  return visit(root, new Map()) ?? new Map();
}

function isNodeWithinAnySubtree(node: XmlNode, subtrees: ReadonlySet<XmlNode>): boolean {
  for (const subtree of subtrees) {
    if (xmlSubtreeContains(subtree, node)) return true;
  }
  return false;
}

function xmlSubtreeContains(root: XmlNode, target: XmlNode): boolean {
  if (root === target) return true;
  for (const [key, value] of Object.entries(root)) {
    if (key.startsWith("@_")) continue;
    for (const child of xmlNodes(value)) {
      if (xmlSubtreeContains(child, target)) return true;
    }
  }
  return false;
}

function removeUnreferencedReachableParts(
  initialGraph: PptxSourceModel["packageGraph"],
  initialTargets: readonly SourceHandle["partPath"][],
): PptxSourceModel["packageGraph"] {
  let graph = initialGraph;
  const candidates = new Map<string, SourceHandle["partPath"]>();
  const pending = [...initialTargets];
  while (pending.length > 0) {
    const partPath = pending.shift();
    if (partPath === undefined || candidates.has(partPath)) continue;
    candidates.set(partPath, partPath);
    const outboundTargets =
      initialGraph.relationships
        .find((group) => group.sourcePartPath === partPath)
        ?.relationships.flatMap((relationship) => {
          const target = resolveInternalRelationshipTarget(partPath, relationship);
          return target === undefined ? [] : [target];
        }) ?? [];
    pending.push(...outboundTargets);
  }

  let changed = true;
  while (changed) {
    changed = false;
    for (const partPath of candidates.values()) {
      const exists =
        graph.parts.some((part) => part.partPath === partPath) ||
        graph.media.some((part) => part.partPath === partPath) ||
        graph.rawParts?.some((part) => part.partPath === partPath) === true;
      if (!exists) continue;
      const hasIncoming = graph.relationships.some((group) =>
        group.relationships.some(
          (relationship) =>
            resolveInternalRelationshipTarget(group.sourcePartPath, relationship) === partPath,
        ),
      );
      if (hasIncoming) continue;
      graph = removePackageParts(graph, [partPath]);
      changed = true;
    }
  }
  return graph;
}

function collectSourceShapeIds(shape: SourceShapeNode): Set<string> {
  const ids = new Set<string>();
  for (const node of flattenSourceShapeTree([shape])) {
    if (node.nodeId !== undefined) ids.add(String(node.nodeId));
  }
  return ids;
}

function collectXmlDrawingIds(node: XmlNode, output: Set<string>): void {
  const id = xmlDrawingId(node);
  if (id !== undefined) output.add(id);
  for (const [key, value] of Object.entries(node)) {
    if (key.startsWith("@_")) continue;
    for (const child of xmlNodes(value)) collectXmlDrawingIds(child, output);
  }
}

function collectSourceRelationshipIds(shape: SourceShapeNode): Set<string> {
  const ids = new Set<string>();
  for (const node of flattenSourceShapeTree([shape])) {
    if (node.kind === "image" && node.blipRelationshipId !== undefined) {
      ids.add(String(node.blipRelationshipId));
    }
    if (node.kind === "chart" && node.chartRelationshipId !== undefined) {
      ids.add(String(node.chartRelationshipId));
    }
    if (node.kind === "smartArt" && node.dataRelationshipId !== undefined) {
      ids.add(String(node.dataRelationshipId));
    }
  }
  return ids;
}

function flattenSourceShapeTree(shapes: readonly SourceShapeNode[]): SourceShapeNode[] {
  return shapes.flatMap((shape) => [
    shape,
    ...(shape.kind === "group" ? flattenSourceShapeTree(shape.children) : []),
  ]);
}

function isDirectlyAddedDrawing(
  edits: readonly PptxSourceModelEdit[],
  handle: SourceHandle,
): boolean {
  return edits.some((edit) => {
    if (
      edit.kind !== "addTextBox" &&
      edit.kind !== "addShape" &&
      edit.kind !== "addConnector" &&
      edit.kind !== "addPicture" &&
      edit.kind !== "addChart" &&
      edit.kind !== "addTable"
    ) {
      return false;
    }
    const inserted = editInsertedShape(edit);
    return (
      inserted?.slidePartPath === handle.partPath && inserted.shapeId === String(handle.nodeId)
    );
  });
}

function isPendingGroupCreationForHandle(edit: PptxSourceModelEdit, handle: SourceHandle): boolean {
  return (
    edit.kind === "groupShapes" &&
    edit.targetPartPath === handle.partPath &&
    edit.groupId === String(handle.nodeId)
  );
}

function collectPendingAddedRelationshipIds(
  edits: readonly PptxSourceModelEdit[],
  handle: SourceHandle,
): ReadonlySet<string> {
  const output = new Set<string>();
  for (const edit of edits) {
    if (
      edit.kind !== "addTextBox" &&
      edit.kind !== "addShape" &&
      edit.kind !== "addConnector" &&
      edit.kind !== "addPicture" &&
      edit.kind !== "addChart" &&
      edit.kind !== "addTable"
    ) {
      continue;
    }
    const inserted = editInsertedShape(edit);
    if (inserted?.slidePartPath !== handle.partPath || inserted.shapeId !== String(handle.nodeId)) {
      continue;
    }
    try {
      const fragment = parseXml(edit.xml);
      collectRelationshipAttributeValues(
        fragment,
        output,
        new Set(),
        new Map([["r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"]]),
      );
    } catch (cause) {
      throw new Error("deleteShape: pending drawing XML could not be validated", { cause });
    }
  }
  return output;
}

function xmlDrawingId(node: XmlNode): string | undefined {
  const nonVisual =
    getChild(node, "nvSpPr") ??
    getChild(node, "nvPicPr") ??
    getChild(node, "nvCxnSpPr") ??
    getChild(node, "nvGrpSpPr") ??
    getChild(node, "nvGraphicFramePr") ??
    getChild(node, "nvContentPartPr");
  return getAttr(getChild(nonVisual, "cNvPr"), "id");
}

function isDrawingElement(value: string): boolean {
  return (
    value === "sp" ||
    value === "cxnSp" ||
    value === "pic" ||
    value === "graphicFrame" ||
    value === "grpSp"
  );
}

function xmlNodes(value: unknown): XmlNode[] {
  const values = Array.isArray(value) ? value : [value];
  return values.filter(
    (item): item is XmlNode => typeof item === "object" && item !== null && !Array.isArray(item),
  );
}
