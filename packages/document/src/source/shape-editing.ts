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
  if (target.nested) {
    throw new Error("deleteShape: nested group shape deletion is not supported");
  }
  if (
    target.node.kind !== "shape" &&
    target.node.kind !== "connector" &&
    target.node.kind !== "image" &&
    target.node.kind !== "table" &&
    target.node.kind !== "chart" &&
    target.node.kind !== "group"
  ) {
    throw new Error(
      "deleteShape: only top-level sp, cxnSp, pic, native table/chart graphicFrame, or grpSp drawings can be deleted",
    );
  }
  assertNotAlternateContentTarget(target, "deleteShape");

  const cleanupPlan = planDrawingDeletionCleanup(source, handle, target.node);

  const slides = source.slides.map((slide) => {
    let slideChanged = false;
    const nextShapes = slide.shapes.filter((shape) => {
      if (!sourceHandlesEqual(shape.handle, handle)) return true;
      slideChanged = true;
      return false;
    });
    return slideChanged ? { ...slide, shapes: nextShapes } : slide;
  });

  const referencingConnector = findConnectorReferencingDeletedSubtree(source, handle, target.node);
  if (referencingConnector !== undefined) {
    throw new Error(
      `deleteShape: shape is referenced by connector '${referencingConnector.name ?? referencingConnector.nodeId ?? "unknown"}'`,
    );
  }

  const retainedEdits = (source.edits ?? []).filter(
    (edit) => isPendingGroupCreationForHandle(edit, handle) || !editTargetsShape(edit, handle),
  );
  const deletedInsertedShape = isDirectlyAddedDrawing(source.edits ?? [], handle);

  return {
    ...source,
    slides,
    packageGraph: cleanupPlan,
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
  const xmlTargets = new Set<XmlNode>();
  if (xmlTarget !== undefined) xmlTargets.add(xmlTarget);
  if (xmlTarget === undefined && pendingGroupCreation && target.kind === "group") {
    for (const child of target.children) {
      if (child.nodeId === undefined) continue;
      const childMatches = findXmlDrawingMatches(spTree, String(child.nodeId)).filter(
        (match) => !match.insideAlternateContent,
      );
      if (childMatches.length > 1)
        throw duplicateNodeIdError("deleteShape", child.handle ?? handle);
      if (childMatches[0] !== undefined) xmlTargets.add(childMatches[0].node);
    }
  }
  for (const deletedXmlTarget of xmlTargets) collectXmlDrawingIds(deletedXmlTarget, deletedIds);
  const rawReferencingConnector = findXmlConnectorReferencingIds(spTree, xmlTargets, deletedIds);
  if (rawReferencingConnector !== undefined) {
    throw new Error(`deleteShape: shape is referenced by connector '${rawReferencingConnector}'`);
  }

  const relationshipPrefixes = collectRelationshipPrefixes(root);
  const candidateRelationshipIds = typedRelationshipIds;
  for (const deletedXmlTarget of xmlTargets) {
    collectRelationshipAttributeValues(
      deletedXmlTarget,
      relationshipPrefixes,
      candidateRelationshipIds,
    );
  }
  if (candidateRelationshipIds.size === 0) return source.packageGraph;

  const remainingRelationshipIds = new Set<string>();
  collectRelationshipAttributeValues(
    root,
    relationshipPrefixes,
    remainingRelationshipIds,
    xmlTargets,
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
  collectRelationshipAttributeValues(
    ownerRoot,
    collectRelationshipPrefixes(ownerRoot),
    retainedRelationshipIds,
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

function collectRelationshipPrefixes(root: XmlNode): ReadonlySet<string> {
  // Finalized edit fragments use the conventional `r` prefix and intentionally omit
  // namespace declarations because the owner drawing root supplies them at write time.
  const prefixes = new Set<string>(["r"]);
  const visit = (node: XmlNode): void => {
    for (const [key, value] of Object.entries(node)) {
      if (key.startsWith("@_xmlns:") && String(value).includes("/relationships")) {
        prefixes.add(key.slice("@_xmlns:".length));
      }
      if (!key.startsWith("@_")) for (const child of xmlNodes(value)) visit(child);
    }
  };
  visit(root);
  return prefixes;
}

function collectRelationshipAttributeValues(
  root: XmlNode,
  relationshipPrefixes: ReadonlySet<string>,
  output: Set<string>,
  skippedSubtrees: ReadonlySet<XmlNode> = new Set(),
): void {
  if (skippedSubtrees.has(root)) return;
  for (const [key, value] of Object.entries(root)) {
    if (key.startsWith("@_")) {
      const qualifiedName = key.slice(2);
      const colon = qualifiedName.indexOf(":");
      if (colon > 0 && relationshipPrefixes.has(qualifiedName.slice(0, colon))) {
        output.add(String(value));
      }
      continue;
    }
    for (const child of xmlNodes(value)) {
      collectRelationshipAttributeValues(child, relationshipPrefixes, output, skippedSubtrees);
    }
  }
}

function removeUnreferencedReachableParts(
  initialGraph: PptxSourceModel["packageGraph"],
  initialTargets: readonly SourceHandle["partPath"][],
): PptxSourceModel["packageGraph"] {
  let graph = initialGraph;
  const pending = [...initialTargets];
  const visited = new Set<string>();
  while (pending.length > 0) {
    const partPath = pending.shift();
    if (partPath === undefined || visited.has(partPath)) continue;
    visited.add(partPath);
    const hasIncoming = graph.relationships.some((group) =>
      group.relationships.some(
        (relationship) =>
          resolveInternalRelationshipTarget(group.sourcePartPath, relationship) === partPath,
      ),
    );
    if (hasIncoming) continue;
    const exists =
      graph.parts.some((part) => part.partPath === partPath) ||
      graph.media.some((part) => part.partPath === partPath) ||
      graph.rawParts?.some((part) => part.partPath === partPath) === true;
    if (!exists) continue;
    const outboundTargets =
      graph.relationships
        .find((group) => group.sourcePartPath === partPath)
        ?.relationships.flatMap((relationship) => {
          const target = resolveInternalRelationshipTarget(partPath, relationship);
          return target === undefined ? [] : [target];
        }) ?? [];
    graph = removePackageParts(graph, [partPath]);
    pending.push(...outboundTargets);
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
    if (key.startsWith("@_") || !isDrawingElement(localName(key))) continue;
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
      collectRelationshipAttributeValues(fragment, collectRelationshipPrefixes(fragment), output);
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
    getChild(node, "nvGraphicFramePr");
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
