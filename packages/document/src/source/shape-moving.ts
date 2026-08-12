import { sourceHandlesEqual } from "./edit-descriptors.js";
import type { PartPath, SourceHandle, SourceNodeId } from "./handles.js";
import type { PptxSourceModel } from "./pptx-source-model.js";
import type { SourceShapeNode } from "./shapes.js";

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

/** Moves one consecutive direct-child drawing block within its current root or native group. */
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

  const selectedIds = new Set<string>();
  const locations: LocatedShape[] = [];
  for (const handle of shapeHandles) {
    if (handle.partPath !== target.partPath) {
      throw new Error("moveShapes: every shape must belong to the destination drawing part");
    }
    if (handle.nodeId === undefined) {
      throw new Error("moveShapes: every shape handle requires a node id");
    }
    const nodeId = String(handle.nodeId);
    if (selectedIds.has(nodeId)) {
      throw new Error("moveShapes: shape handles contain a duplicate shape");
    }
    selectedIds.add(nodeId);
    const index = target.directChildren.findIndex((shape) =>
      sourceHandlesEqual(shape.handle, handle),
    );
    const shape = target.directChildren[index];
    if (shape === undefined) {
      if (target.shapes.some((candidate) => containsNestedHandle(candidate, handle))) {
        throw new Error("moveShapes: every shape must be a direct child of the destination");
      }
      throw new Error("moveShapes: shape handle was not found in the destination drawing part");
    }
    if (shape.nodeId === undefined) {
      throw new Error("moveShapes: every selected shape requires a node id");
    }
    assertSupportedTarget(shape);
    locations.push({ shape, index });
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
    if (selectedIds.has(beforeShapeId)) {
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
    assertSupportedTarget(anchor);
  }

  const movedShapes = orderedLocations.map((location) => location.shape);
  const remaining = target.directChildren.filter((shape) => !selectedIds.has(String(shape.nodeId)));
  const insertionIndex =
    beforeShapeId === undefined
      ? remaining.length
      : remaining.findIndex((shape) => String(shape.nodeId) === beforeShapeId);
  if (insertionIndex < 0) {
    throw new Error("moveShapes: anchor was not found after removing the moved block");
  }
  const nextChildren = [
    ...remaining.slice(0, insertionIndex),
    ...movedShapes,
    ...remaining.slice(insertionIndex),
  ];
  const updated = withTargetShapes(source, target, nextChildren);

  return {
    ...updated,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "moveShapes",
        targetPartPath: target.partPath,
        ...(target.parentGroupId !== undefined
          ? { parentGroupId: String(target.parentGroupId) }
          : {}),
        shapeIds: movedShapes.map((shape) => String(shape.nodeId)),
        ...(beforeShapeId !== undefined ? { beforeShapeId } : {}),
      },
    ],
  };
}

function assertSupportedTarget(shape: SourceShapeNode): void {
  if (shape.kind === "smartArt" || shape.kind === "raw") {
    throw new Error("moveShapes: SmartArt and raw drawing targets are not supported");
  }
  if (hasAlternateContent(shape)) {
    throw new Error("moveShapes: mc:AlternateContent targets are not supported");
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

function localName(name: string): string {
  return name.split(":").at(-1) ?? name;
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
