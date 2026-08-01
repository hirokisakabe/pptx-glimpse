/**
 * Lossless topology edits for existing DrawingML groups.
 *
 * The supported slice moves existing sibling nodes without rewriting their transforms or ids.
 * Grouping creates an identity child-coordinate mapping around the selected bounds; ungrouping
 * accepts only groups whose children already use that identity mapping.
 */

import { nextDrawingShapeId } from "./drawing-authoring-allocation.js";
import type { PartPath, SourceHandle, SourceNodeId } from "./handles.js";
import type { PptxSourceModel } from "./pptx-source-model.js";
import type { SourceGroup, SourceShapeNode, SourceTransform } from "./shapes.js";
import { asEmu } from "./units.js";

type DrawingTargetKind = "slide" | "layout" | "master";

interface DrawingTarget {
  readonly kind: DrawingTargetKind;
  readonly index: number;
  readonly partPath: PartPath;
  readonly shapes: readonly SourceShapeNode[];
}

interface NodeLocation {
  readonly node: SourceShapeNode;
  readonly parent: readonly SourceShapeNode[];
  readonly parentGroupId?: SourceNodeId;
  readonly index: number;
}

/** Groups two or more consecutive sibling nodes and preserves every child node id and handle. */
export function groupShapes(
  source: PptxSourceModel,
  shapeHandles: readonly SourceHandle[],
): PptxSourceModel {
  if (shapeHandles.length < 2) {
    throw new Error("groupShapes: at least two shape handles are required");
  }
  const partPath = shapeHandles[0]?.partPath;
  if (partPath === undefined || shapeHandles.some((handle) => handle.partPath !== partPath)) {
    throw new Error("groupShapes: every shape handle must belong to the same drawing part");
  }
  const target = findDrawingTarget(source, partPath);
  if (target === undefined) {
    throw new Error("groupShapes: target drawing part was not found");
  }
  assertUniqueNodeIds(target.shapes, "groupShapes");

  const selectedIds = new Set<string>();
  const locations: NodeLocation[] = [];
  for (const handle of shapeHandles) {
    if (handle.nodeId === undefined) {
      throw new Error("groupShapes: every shape handle requires a node id");
    }
    const nodeId = String(handle.nodeId);
    if (selectedIds.has(nodeId)) {
      throw new Error("groupShapes: shape handles contain a duplicate shape");
    }
    selectedIds.add(nodeId);
    const location = findNodeLocation(target.shapes, nodeId);
    if (location === undefined) {
      throw new Error(`groupShapes: shape '${nodeId}' was not found in the target drawing part`);
    }
    if (location.node.nodeId === undefined) {
      throw new Error("groupShapes: every selected shape requires a node id");
    }
    if (isAlternateContentNode(location.node)) {
      throw new Error("groupShapes: mc:AlternateContent nodes are not supported");
    }
    locations.push(location);
  }

  const parent = locations[0]?.parent;
  const parentGroupId = locations[0]?.parentGroupId;
  if (
    parent === undefined ||
    locations.some(
      (location) =>
        location.parent !== parent || String(location.parentGroupId) !== String(parentGroupId),
    )
  ) {
    throw new Error("groupShapes: selected shapes must have the same immediate parent");
  }
  const orderedLocations = [...locations].sort((left, right) => left.index - right.index);
  const firstIndex = orderedLocations[0]?.index ?? -1;
  if (orderedLocations.some((location, index) => location.index !== firstIndex + index)) {
    throw new Error("groupShapes: selected shapes must be consecutive siblings");
  }

  assertConnectorBoundary(target.shapes, selectedIds);
  const children = orderedLocations.map((location) => location.node);
  const bounds = unionTransformBounds(children);
  const groupId = nextDrawingShapeId(source, target.shapes, target.partPath);
  const transform: SourceTransform = {
    offsetX: asEmu(Math.floor(bounds.left)),
    offsetY: asEmu(Math.floor(bounds.top)),
    width: asEmu(Math.ceil(bounds.right) - Math.floor(bounds.left)),
    height: asEmu(Math.ceil(bounds.bottom) - Math.floor(bounds.top)),
  };
  const group: SourceGroup = {
    kind: "group",
    nodeId: groupId,
    name: `Group ${groupId}`,
    transform,
    childTransform: transform,
    children,
    handle: {
      partPath: target.partPath,
      nodeId: groupId,
      ...(children[0]?.handle?.orderingSlot !== undefined
        ? { orderingSlot: children[0].handle.orderingSlot }
        : {}),
    },
  };
  const nextParent = [
    ...parent.slice(0, firstIndex),
    group,
    ...parent.slice(firstIndex + children.length),
  ];
  const updated = withReplacedParent(source, target, parentGroupId, nextParent, "groupShapes");

  return {
    ...updated,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "groupShapes",
        targetPartPath: target.partPath,
        ...(parentGroupId !== undefined ? { parentGroupId: String(parentGroupId) } : {}),
        shapeIds: children.map((child) => String(child.nodeId)),
        groupId: String(groupId),
        groupName: group.name ?? `Group ${groupId}`,
        offsetX: transform.offsetX,
        offsetY: transform.offsetY,
        width: transform.width,
        height: transform.height,
      },
    ],
  };
}

/** Expands one identity-mapped group into its immediate parent without changing child nodes. */
export function ungroupShape(source: PptxSourceModel, groupHandle: SourceHandle): PptxSourceModel {
  if (groupHandle.nodeId === undefined) {
    throw new Error("ungroupShape: group handle requires a node id");
  }
  const target = findDrawingTarget(source, groupHandle.partPath);
  if (target === undefined) {
    throw new Error("ungroupShape: target drawing part was not found");
  }
  assertUniqueNodeIds(target.shapes, "ungroupShape");
  const groupId = String(groupHandle.nodeId);
  const location = findNodeLocation(target.shapes, groupId);
  if (location === undefined || location.node.kind !== "group") {
    throw new Error("ungroupShape: handle does not reference a group shape");
  }
  const group = location.node;
  if (!hasIdentityChildMapping(group)) {
    throw new Error("ungroupShape: group transform is not a lossless identity child mapping");
  }
  if (
    group.fill !== undefined ||
    group.effects !== undefined ||
    hasUnsupportedGroupSidecars(group)
  ) {
    throw new Error("ungroupShape: group appearance or unknown XML cannot be losslessly expanded");
  }
  const referencingConnector = findConnectorReferencingId(target.shapes, groupId);
  if (referencingConnector !== undefined) {
    throw new Error(
      `ungroupShape: group is referenced by connector '${referencingConnector.name ?? referencingConnector.nodeId ?? "unknown"}'`,
    );
  }

  const nextParent = [
    ...location.parent.slice(0, location.index),
    ...group.children,
    ...location.parent.slice(location.index + 1),
  ];
  const updated = withReplacedParent(
    source,
    target,
    location.parentGroupId,
    nextParent,
    "ungroupShape",
  );
  return {
    ...updated,
    edits: [
      ...(source.edits ?? []),
      { kind: "ungroupShape", targetPartPath: target.partPath, groupId },
    ],
  };
}

function hasUnsupportedGroupSidecars(group: SourceGroup): boolean {
  const childNodeNames = new Set(["sp", "pic", "cxnSp", "graphicFrame", "grpSp"]);
  return (
    group.rawSidecars?.some(
      (sidecar) => !childNodeNames.has(sidecar.node.name.split(":").at(-1) ?? ""),
    ) ?? false
  );
}

function findDrawingTarget(source: PptxSourceModel, partPath: PartPath): DrawingTarget | undefined {
  const collections = [
    ["slide", source.slides],
    ["layout", source.slideLayouts],
    ["master", source.slideMasters],
  ] as const;
  for (const [kind, targets] of collections) {
    const index = targets.findIndex((candidate) => candidate.partPath === partPath);
    const target = targets[index];
    if (target !== undefined) return { kind, index, partPath, shapes: target.shapes };
  }
  return undefined;
}

function findNodeLocation(
  nodes: readonly SourceShapeNode[],
  nodeId: string,
  parentGroupId?: SourceNodeId,
): NodeLocation | undefined {
  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    if (node === undefined) continue;
    if (String(node.nodeId) === nodeId) return { node, parent: nodes, parentGroupId, index };
    if (node.kind === "group") {
      const nested = findNodeLocation(node.children, nodeId, node.nodeId);
      if (nested !== undefined) return nested;
    }
  }
  return undefined;
}

function withReplacedParent(
  source: PptxSourceModel,
  target: DrawingTarget,
  parentGroupId: SourceNodeId | undefined,
  shapes: readonly SourceShapeNode[],
  operationName: string,
): PptxSourceModel {
  const nextRootShapes =
    parentGroupId === undefined
      ? shapes
      : replaceGroupChildren(target.shapes, String(parentGroupId), shapes, operationName);
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
  operationName: string,
): readonly SourceShapeNode[] {
  const replaced = tryReplaceGroupChildren(nodes, groupId, children);
  if (replaced !== undefined) return replaced;
  throw new Error(`${operationName}: immediate parent group was not found`);
}

function tryReplaceGroupChildren(
  nodes: readonly SourceShapeNode[],
  groupId: string,
  children: readonly SourceShapeNode[],
): readonly SourceShapeNode[] | undefined {
  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    if (node?.kind !== "group") continue;
    if (String(node.nodeId) === groupId) {
      return [...nodes.slice(0, index), { ...node, children }, ...nodes.slice(index + 1)];
    }
    const nested = tryReplaceGroupChildren(node.children, groupId, children);
    if (nested !== undefined) {
      return [...nodes.slice(0, index), { ...node, children: nested }, ...nodes.slice(index + 1)];
    }
  }
  return undefined;
}

function assertUniqueNodeIds(nodes: readonly SourceShapeNode[], operationName: string): void {
  const seen = new Set<string>();
  const visit = (children: readonly SourceShapeNode[]): void => {
    for (const child of children) {
      if (child.nodeId !== undefined) {
        const nodeId = String(child.nodeId);
        if (seen.has(nodeId)) {
          throw new Error(
            `${operationName}: duplicate node id '${nodeId}' in the target drawing part is not supported`,
          );
        }
        seen.add(nodeId);
      }
      if (child.kind === "group") visit(child.children);
    }
  };
  visit(nodes);
}

function assertConnectorBoundary(
  nodes: readonly SourceShapeNode[],
  selectedIds: ReadonlySet<string>,
): void {
  visitNodes(nodes, (node) => {
    if (node.kind !== "connector" || node.nodeId === undefined) return;
    const connectorSelected = selectedIds.has(String(node.nodeId));
    for (const endpoint of [node.connection?.start, node.connection?.end]) {
      if (endpoint === undefined) continue;
      if (connectorSelected !== selectedIds.has(String(endpoint.shapeId))) {
        throw new Error("groupShapes: connector endpoint crosses the selection boundary");
      }
    }
  });
}

function findConnectorReferencingId(
  nodes: readonly SourceShapeNode[],
  nodeId: string,
): Extract<SourceShapeNode, { readonly kind: "connector" }> | undefined {
  let match: Extract<SourceShapeNode, { readonly kind: "connector" }> | undefined;
  visitNodes(nodes, (node) => {
    if (
      match === undefined &&
      node.kind === "connector" &&
      [node.connection?.start, node.connection?.end].some(
        (endpoint) => String(endpoint?.shapeId) === nodeId,
      )
    ) {
      match = node;
    }
  });
  return match;
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

function unionTransformBounds(nodes: readonly SourceShapeNode[]): {
  left: number;
  top: number;
  right: number;
  bottom: number;
} {
  const bounds = nodes.map((node) => {
    if (node.kind === "raw" || node.transform === undefined) {
      throw new Error("groupShapes: every selected shape requires a complete transform");
    }
    return rotatedTransformBounds(node.transform);
  });
  return {
    left: Math.min(...bounds.map((bound) => bound.left)),
    top: Math.min(...bounds.map((bound) => bound.top)),
    right: Math.max(...bounds.map((bound) => bound.right)),
    bottom: Math.max(...bounds.map((bound) => bound.bottom)),
  };
}

function rotatedTransformBounds(transform: SourceTransform): {
  left: number;
  top: number;
  right: number;
  bottom: number;
} {
  const left = Number(transform.offsetX);
  const top = Number(transform.offsetY);
  const width = Number(transform.width);
  const height = Number(transform.height);
  if (![left, top, width, height].every(Number.isFinite)) {
    throw new Error("groupShapes: every selected shape transform must contain finite values");
  }
  const centerX = left + width / 2;
  const centerY = top + height / 2;
  const radians = (Number(transform.rotation ?? 0) / 60000 / 180) * Math.PI;
  const cosine = Math.cos(radians);
  const sine = Math.sin(radians);
  const corners = [
    [left, top],
    [left + width, top],
    [left + width, top + height],
    [left, top + height],
  ].map(([x, y]) => ({
    x: snapNearInteger(centerX + (x - centerX) * cosine - (y - centerY) * sine),
    y: snapNearInteger(centerY + (x - centerX) * sine + (y - centerY) * cosine),
  }));
  return {
    left: Math.min(...corners.map((corner) => corner.x)),
    top: Math.min(...corners.map((corner) => corner.y)),
    right: Math.max(...corners.map((corner) => corner.x)),
    bottom: Math.max(...corners.map((corner) => corner.y)),
  };
}

function snapNearInteger(value: number): number {
  const rounded = Math.round(value);
  return Math.abs(value - rounded) < 1e-7 ? rounded : value;
}

function hasIdentityChildMapping(group: SourceGroup): boolean {
  const transform = group.transform;
  const child = group.childTransform;
  return (
    transform !== undefined &&
    child !== undefined &&
    Number(transform.rotation ?? 0) === 0 &&
    transform.flipHorizontal !== true &&
    transform.flipVertical !== true &&
    transform.offsetX === child.offsetX &&
    transform.offsetY === child.offsetY &&
    transform.width === child.width &&
    transform.height === child.height
  );
}

function isAlternateContentNode(node: SourceShapeNode): boolean {
  if (node.kind === "raw") {
    return node.raw.node.name.split(":").at(-1) === "AlternateContent";
  }
  return (
    node.rawSidecars?.some(
      (sidecar) => sidecar.node.name.split(":").at(-1) === "AlternateContent",
    ) ?? false
  );
}
