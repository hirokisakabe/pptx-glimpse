import { getAttr, getChild, localName, parseXml, type XmlNode } from "../reader/xml.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import { sourceHandlesEqual } from "./edit-descriptors.js";
import { reparentSourceTransform } from "./group-transform-matrix.js";
import type { PartPath, SourceHandle, SourceNodeId } from "./handles.js";
import type { PptxSourceModel } from "./pptx-source-model.js";
import type { SourceGroup, SourceShapeNode, SourceTransform } from "./shapes.js";

export interface MoveShapesOptions {
  /** Insert the moved block immediately before this direct child. Omit to move it to the end. */
  readonly beforeShapeHandle?: SourceHandle;
}

type RootTarget = {
  readonly kind: "slide" | "layout" | "master";
  readonly index: number;
  readonly partPath: PartPath;
  readonly shapes: readonly SourceShapeNode[];
};

type Target = RootTarget & {
  readonly parentGroupId?: SourceNodeId;
  readonly directChildren: readonly SourceShapeNode[];
};

interface LocatedShape {
  readonly shape: SourceShapeNode;
  readonly index: number;
}

interface SourceLocation extends LocatedShape {
  readonly parentGroupId?: SourceNodeId;
  readonly directChildren: readonly SourceShapeNode[];
  readonly ancestors: readonly SourceGroup[];
}

/**
 * Moves one consecutive direct-child drawing block within one drawing part. Cross-parent moves
 * re-express moved root transforms in the destination coordinate space when the exact affine
 * mapping is representable as integer OOXML transform values.
 */
export function moveShapes(
  source: PptxSourceModel,
  shapeHandles: readonly SourceHandle[],
  destinationHandle: SourceHandle,
  options: MoveShapesOptions = {},
): PptxSourceModel {
  if (shapeHandles.length === 0) {
    throw new Error("moveShapes: at least one shape handle is required");
  }
  const target = findTarget(source, destinationHandle);
  if (target === undefined) {
    throw new Error("moveShapes: destination must be a slide, layout, master, or native group");
  }
  if (target.shapes.some(hasAlternateContent)) {
    throw new Error("moveShapes: mc:AlternateContent shape trees are not supported");
  }
  assertUniqueDrawingNodeIds(target.shapes);

  const selectedRootIds = new Set<string>();
  const locations: SourceLocation[] = [];
  for (const handle of shapeHandles) {
    if (handle.partPath !== target.partPath) {
      throw new Error("moveShapes: every shape must belong to the destination drawing part");
    }
    if (handle.nodeId === undefined) {
      throw new Error("moveShapes: every shape handle requires a node id");
    }
    const nodeId = String(handle.nodeId);
    if (selectedRootIds.has(nodeId)) {
      throw new Error("moveShapes: shape handles contain a duplicate shape");
    }
    selectedRootIds.add(nodeId);
    const location = locateShape(target.shapes, handle);
    if (location === undefined) {
      throw new Error("moveShapes: shape handle was not found in the destination drawing part");
    }
    const { shape } = location;
    if (shape.nodeId === undefined) {
      throw new Error("moveShapes: every selected shape requires a node id");
    }
    assertSupportedSubtree(shape);
    locations.push(location);
  }

  const sourceParentId = locations[0]?.parentGroupId;
  if (locations.some((location) => location.parentGroupId !== sourceParentId)) {
    throw new Error("moveShapes: selected shapes must share one immediate parent");
  }

  const orderedLocations = [...locations].sort((left, right) => left.index - right.index);
  const firstIndex = orderedLocations[0]?.index ?? -1;
  if (orderedLocations.some((location, index) => location.index !== firstIndex + index)) {
    throw new Error("moveShapes: selected shapes must be consecutive siblings");
  }

  let beforeShapeId: string | undefined;
  if (options.beforeShapeHandle !== undefined) {
    const beforeHandle = options.beforeShapeHandle;
    if (beforeHandle.partPath !== target.partPath) {
      throw new Error("moveShapes: anchor belongs to a different drawing part");
    }
    if (beforeHandle.nodeId === undefined) {
      throw new Error("moveShapes: anchor handle requires a node id");
    }
    beforeShapeId = String(beforeHandle.nodeId);
    if (selectedRootIds.has(beforeShapeId)) {
      throw new Error("moveShapes: anchor must not be inside the moved block");
    }
    const anchor = target.directChildren.find((shape) =>
      sourceHandlesEqual(shape.handle, beforeHandle),
    );
    if (anchor === undefined) {
      if (target.shapes.some((candidate) => containsNestedHandle(candidate, beforeHandle))) {
        throw new Error("moveShapes: anchor must be a direct child of the destination");
      }
      throw new Error("moveShapes: anchor was not found in the destination drawing part");
    }
    if (anchor.nodeId === undefined) {
      throw new Error("moveShapes: anchor shape requires a node id");
    }
    assertSupportedSubtree(anchor);
  }

  let movedShapes = orderedLocations.map((location) => location.shape);
  const crossParent = sourceParentId !== target.parentGroupId;
  let transformedRoots:
    | readonly { readonly shapeId: string; readonly transform: SourceTransform }[]
    | undefined;
  if (crossParent) {
    const sourceAncestors = orderedLocations[0]?.ancestors ?? [];
    const targetAncestors = destinationAncestors(target.shapes, target.parentGroupId);
    if (
      !isIdentityMappedAncestorChain(sourceAncestors) ||
      !isIdentityMappedAncestorChain(targetAncestors)
    ) {
      const transformed = movedShapes.map((shape) => {
        const result = reparentSourceTransform(
          sourceAncestors,
          targetAncestors,
          movableRootTransform(shape),
        );
        if (!result.ok) {
          throw new Error(
            `moveShapes: affine transform is not exactly representable (${result.reason})`,
          );
        }
        return { shapeId: String(shape.nodeId), transform: result.value };
      });
      transformedRoots = transformed;
      const transformById = new Map(transformed.map((entry) => [entry.shapeId, entry.transform]));
      movedShapes = movedShapes.map((shape) =>
        withMovableRootTransform(
          shape,
          transformById.get(String(shape.nodeId)) ?? movableRootTransform(shape),
        ),
      );
    }
    const movedIds = collectSubtreeNodeIds(movedShapes);
    if (target.parentGroupId !== undefined && movedIds.has(String(target.parentGroupId))) {
      throw new Error("moveShapes: destination must not be inside the moved block");
    }
    assertConnectorBoundary(target.shapes, movedIds);
    assertRawXmlConnectorBoundary(source, target, movedIds);
  }

  const remaining = orderedLocations[0]?.directChildren.filter(
    (shape) => !selectedRootIds.has(String(shape.nodeId)),
  );
  if (remaining === undefined) {
    throw new Error("moveShapes: source parent could not be resolved");
  }
  const withoutMoved = withTargetShapes(
    source,
    { ...target, parentGroupId: sourceParentId },
    remaining,
  );
  const updatedTarget = findTarget(withoutMoved, destinationHandle);
  if (updatedTarget === undefined) {
    throw new Error("moveShapes: destination disappeared while applying the move");
  }
  const destinationChildren = crossParent ? updatedTarget.directChildren : remaining;
  const insertionIndex =
    beforeShapeId === undefined
      ? destinationChildren.length
      : destinationChildren.findIndex((shape) => String(shape.nodeId) === beforeShapeId);
  if (insertionIndex < 0) {
    throw new Error("moveShapes: anchor was not found after removing the moved block");
  }
  const nextChildren = [
    ...destinationChildren.slice(0, insertionIndex),
    ...movedShapes,
    ...destinationChildren.slice(insertionIndex),
  ];
  const updated = withTargetShapes(withoutMoved, updatedTarget, nextChildren);

  return {
    ...updated,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "moveShapes",
        targetPartPath: target.partPath,
        ...(sourceParentId !== undefined ? { parentGroupId: String(sourceParentId) } : {}),
        ...(crossParent ? { crossParent: true as const } : {}),
        ...(crossParent && target.parentGroupId !== undefined
          ? { destinationParentGroupId: String(target.parentGroupId) }
          : {}),
        ...(transformedRoots !== undefined ? { transformedRoots } : {}),
        shapeIds: movedShapes.map((shape) => String(shape.nodeId)),
        ...(beforeShapeId !== undefined ? { beforeShapeId } : {}),
      },
    ],
  };
}

function assertSupportedSubtree(shape: SourceShapeNode): void {
  if (shape.kind === "smartArt" || shape.kind === "raw") {
    throw new Error("moveShapes: SmartArt and raw drawing targets are not supported");
  }
  if (hasAlternateContent(shape)) {
    throw new Error("moveShapes: mc:AlternateContent targets are not supported");
  }
  if (shape.kind === "group") {
    for (const child of shape.children) assertSupportedSubtree(child);
  }
}

function movableRootTransform(shape: SourceShapeNode): SourceTransform | undefined {
  if (shape.kind === "raw") return undefined;
  return shape.transform;
}

function withMovableRootTransform(
  shape: SourceShapeNode,
  transform: SourceTransform | undefined,
): SourceShapeNode {
  if (shape.kind === "raw" || shape.kind === "smartArt") {
    throw new Error("moveShapes: SmartArt and raw drawing targets are not supported");
  }
  return { ...shape, transform };
}

function isIdentityMappedAncestorChain(ancestors: readonly SourceGroup[]): boolean {
  for (const group of ancestors) {
    const transform = group.transform;
    const child = group.childTransform;
    if (
      transform === undefined ||
      child === undefined ||
      ![
        transform.offsetX,
        transform.offsetY,
        transform.width,
        transform.height,
        transform.rotation ?? 0,
        child.offsetX,
        child.offsetY,
        child.width,
        child.height,
      ].every((value) => Number.isFinite(Number(value))) ||
      Number(transform.width) <= 0 ||
      Number(transform.height) <= 0 ||
      Number(child.width) <= 0 ||
      Number(child.height) <= 0 ||
      Number(transform.offsetX) !== Number(child.offsetX) ||
      Number(transform.offsetY) !== Number(child.offsetY) ||
      Number(transform.width) !== Number(child.width) ||
      Number(transform.height) !== Number(child.height) ||
      Number(transform.rotation ?? 0) !== 0 ||
      transform.flipHorizontal === true ||
      transform.flipVertical === true
    ) {
      return false;
    }
  }
  return true;
}

function destinationAncestors(
  shapes: readonly SourceShapeNode[],
  parentGroupId: SourceNodeId | undefined,
): readonly SourceGroup[] {
  if (parentGroupId === undefined) return [];
  const ancestors = findGroupAncestors(shapes, String(parentGroupId));
  if (ancestors === undefined) throw new Error("moveShapes: destination group was not found");
  return ancestors;
}

function collectSubtreeNodeIds(nodes: readonly SourceShapeNode[]): Set<string> {
  const ids = new Set<string>();
  visitNodes(nodes, (node) => {
    if (node.nodeId !== undefined) ids.add(String(node.nodeId));
  });
  return ids;
}

function assertConnectorBoundary(
  nodes: readonly SourceShapeNode[],
  movedIds: ReadonlySet<string>,
): void {
  visitNodes(nodes, (node) => {
    if (node.kind !== "connector") return;
    const connectorMoved = node.nodeId !== undefined && movedIds.has(String(node.nodeId));
    for (const endpoint of [node.connection?.start, node.connection?.end]) {
      if (endpoint !== undefined && connectorMoved !== movedIds.has(String(endpoint.shapeId))) {
        throw new Error("moveShapes: connector endpoint crosses the moved block boundary");
      }
    }
  });
}

function assertRawXmlConnectorBoundary(
  source: PptxSourceModel,
  target: Target,
  movedIds: ReadonlySet<string>,
): void {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === target.partPath);
  // From-scratch and intentionally compact source models may not retain raw part XML. Typed
  // connectors were validated above, and the writer validates the effective XML again when the
  // edit journal is replayed. Only scan here when preserved XML can contain untyped connectors.
  if (rawPart?.kind !== "binary") return;
  let root: XmlNode;
  try {
    root = parseXml(new TextDecoder().decode(rawPart.bytes));
  } catch (cause) {
    throw new Error(
      `moveShapes: drawing part '${target.partPath}' could not be parsed for connector validation`,
      { cause },
    );
  }

  const currentIds = collectSubtreeNodeIds(target.shapes);
  const visit = (node: XmlNode): void => {
    for (const [key, value] of Object.entries(node)) {
      if (key.startsWith("@_")) continue;
      for (const child of xmlNodes(value)) {
        if (localName(key) === "cxnSp") {
          const nonVisual = getChild(child, "nvCxnSpPr");
          const connectorId = getAttr(getChild(nonVisual, "cNvPr"), "id");
          // A raw connector removed by an earlier chronological delete is no longer effective.
          if (connectorId !== undefined && !currentIds.has(connectorId)) continue;
          const connectionProperties = getChild(nonVisual, "cNvCxnSpPr");
          const connectorMoved = connectorId !== undefined && movedIds.has(connectorId);
          for (const endpointId of [
            getAttr(getChild(connectionProperties, "stCxn"), "id"),
            getAttr(getChild(connectionProperties, "endCxn"), "id"),
          ]) {
            if (endpointId !== undefined && connectorMoved !== movedIds.has(endpointId)) {
              throw new Error(
                "moveShapes: raw XML connector endpoint crosses the moved block boundary",
              );
            }
          }
        }
        visit(child);
      }
    }
  };
  visit(root);
}

function xmlNodes(value: unknown): XmlNode[] {
  const values = Array.isArray(value) ? value : [value];
  return values.flatMap((candidate) =>
    typeof candidate === "object" && candidate !== null && !Array.isArray(candidate)
      ? [unsafeOoxmlBoundaryAssertion<XmlNode>(candidate)]
      : [],
  );
}

function visitNodes(
  nodes: readonly SourceShapeNode[],
  visitor: (node: SourceShapeNode) => void,
): void {
  for (const node of nodes) {
    visitor(node);
    if (node.kind === "group") visitNodes(node.children, visitor);
  }
}

function assertUniqueDrawingNodeIds(shapes: readonly SourceShapeNode[]): void {
  const seen = new Set<string>();
  const visit = (nodes: readonly SourceShapeNode[]): void => {
    for (const node of nodes) {
      if (node.nodeId !== undefined) {
        const nodeId = String(node.nodeId);
        if (seen.has(nodeId)) {
          throw new Error(
            `moveShapes: duplicate node id '${nodeId}' in the target drawing part is not supported`,
          );
        }
        seen.add(nodeId);
      }
      if (node.kind === "group") visit(node.children);
    }
  };
  visit(shapes);
}

function containsNestedHandle(shape: SourceShapeNode, handle: SourceHandle): boolean {
  if (sourceHandlesEqual(shape.handle, handle)) return true;
  return (
    shape.kind === "group" && shape.children.some((child) => containsNestedHandle(child, handle))
  );
}

function locateShape(
  nodes: readonly SourceShapeNode[],
  handle: SourceHandle,
  ancestors: readonly SourceGroup[] = [],
): SourceLocation | undefined {
  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    if (node === undefined) continue;
    if (sourceHandlesEqual(node.handle, handle)) {
      return {
        shape: node,
        index,
        directChildren: nodes,
        ancestors,
        ...(ancestors.at(-1)?.nodeId !== undefined
          ? { parentGroupId: ancestors.at(-1)?.nodeId }
          : {}),
      };
    }
    if (node.kind === "group") {
      const nested = locateShape(node.children, handle, [...ancestors, node]);
      if (nested !== undefined) return nested;
    }
  }
  return undefined;
}

function findGroupAncestors(
  nodes: readonly SourceShapeNode[],
  groupId: string,
  ancestors: readonly SourceGroup[] = [],
): readonly SourceGroup[] | undefined {
  for (const node of nodes) {
    if (node.kind !== "group") continue;
    const next = [...ancestors, node];
    if (String(node.nodeId) === groupId) return next;
    const nested = findGroupAncestors(node.children, groupId, next);
    if (nested !== undefined) return nested;
  }
  return undefined;
}

function hasAlternateContent(shape: SourceShapeNode): boolean {
  if (shape.kind === "raw") return localName(shape.raw.node.name) === "AlternateContent";
  if (
    shape.rawSidecars?.some((sidecar) => localName(sidecar.node.name) === "AlternateContent") ??
    false
  ) {
    return true;
  }
  return shape.kind === "group" && shape.children.some(hasAlternateContent);
}

function findTarget(source: PptxSourceModel, handle: SourceHandle): Target | undefined {
  const collections = [
    ["slide", source.slides],
    ["layout", source.slideLayouts],
    ["master", source.slideMasters],
  ] as const;
  for (const [kind, targets] of collections) {
    const index = targets.findIndex((candidate) => sourceHandlesEqual(candidate.handle, handle));
    const root = targets[index];
    if (root !== undefined) {
      return {
        kind,
        index,
        partPath: root.partPath,
        shapes: root.shapes,
        directChildren: root.shapes,
      };
    }
  }
  for (const [kind, targets] of collections) {
    for (let index = 0; index < targets.length; index += 1) {
      const root = targets[index];
      if (root === undefined || root.partPath !== handle.partPath) continue;
      const group = findGroup(root.shapes, handle);
      if (group !== undefined && group.nodeId !== undefined) {
        return {
          kind,
          index,
          partPath: root.partPath,
          shapes: root.shapes,
          parentGroupId: group.nodeId,
          directChildren: group.children,
        };
      }
    }
  }
  return undefined;
}

function findGroup(
  nodes: readonly SourceShapeNode[],
  handle: SourceHandle,
): Extract<SourceShapeNode, { kind: "group" }> | undefined {
  for (const node of nodes) {
    if (node.kind !== "group") continue;
    if (sourceHandlesEqual(node.handle, handle)) return node;
    const nested = findGroup(node.children, handle);
    if (nested !== undefined) return nested;
  }
  return undefined;
}

function withTargetShapes(
  source: PptxSourceModel,
  target: Target,
  shapes: readonly SourceShapeNode[],
): PptxSourceModel {
  const nextRootShapes =
    target.parentGroupId === undefined
      ? shapes
      : replaceGroupChildren(target.shapes, String(target.parentGroupId), shapes);
  switch (target.kind) {
    case "slide":
      return {
        ...source,
        slides: source.slides.map((slide, index) =>
          index === target.index ? { ...slide, shapes: nextRootShapes } : slide,
        ),
      };
    case "layout":
      return {
        ...source,
        slideLayouts: source.slideLayouts.map((layout, index) =>
          index === target.index ? { ...layout, shapes: nextRootShapes } : layout,
        ),
      };
    case "master":
      return {
        ...source,
        slideMasters: source.slideMasters.map((master, index) =>
          index === target.index ? { ...master, shapes: nextRootShapes } : master,
        ),
      };
  }
}

function replaceGroupChildren(
  nodes: readonly SourceShapeNode[],
  groupId: string,
  children: readonly SourceShapeNode[],
): readonly SourceShapeNode[] {
  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    if (node?.kind !== "group") continue;
    if (String(node.nodeId) === groupId) {
      return [...nodes.slice(0, index), { ...node, children }, ...nodes.slice(index + 1)];
    }
    const nested = replaceGroupChildren(node.children, groupId, children);
    if (nested !== node.children) {
      return [...nodes.slice(0, index), { ...node, children: nested }, ...nodes.slice(index + 1)];
    }
  }
  return nodes;
}
