import { sourceHandlesEqual } from "./edit-descriptors.js";
import type { PartPath, SourceHandle, SourceNodeId } from "./handles.js";
import type { PptxSourceModel } from "./pptx-source-model.js";
import type { SourceShapeNode } from "./shapes.js";

type RootTarget = {
  kind: "slide" | "layout" | "master";
  index: number;
  partPath: PartPath;
  shapes: readonly SourceShapeNode[];
};

type Target = RootTarget & {
  readonly parentGroupId?: SourceNodeId;
  readonly directChildren: readonly SourceShapeNode[];
};

/** Reorders every direct child of one slide, layout, master, or native group. */
export function reorderShapes(
  source: PptxSourceModel,
  targetHandle: SourceHandle,
  orderedShapeHandles: readonly SourceHandle[],
): PptxSourceModel {
  const target = findTarget(source, targetHandle);
  if (target === undefined) {
    throw new Error("reorderShapes: target handle was not found");
  }
  if (target.shapes.some(hasAlternateContent)) {
    throw new Error("reorderShapes: mc:AlternateContent shape trees are not supported");
  }
  assertUniqueDrawingNodeIds(target.shapes);
  if (orderedShapeHandles.length !== target.directChildren.length) {
    throw new Error("reorderShapes: ordered handles must contain every target shape exactly once");
  }

  const orderedShapes: SourceShapeNode[] = [];
  const seen = new Set<string>();
  for (const handle of orderedShapeHandles) {
    if (handle.partPath !== target.partPath) {
      throw new Error("reorderShapes: shape handle belongs to a different drawing part");
    }
    if (handle.nodeId === undefined) {
      throw new Error("reorderShapes: every shape handle requires a node id");
    }
    const nodeId = String(handle.nodeId);
    if (seen.has(nodeId)) {
      throw new Error("reorderShapes: ordered handles contain a duplicate shape");
    }
    seen.add(nodeId);
    const shape = target.directChildren.find((candidate) =>
      sourceHandlesEqual(candidate.handle, handle),
    );
    if (shape === undefined) {
      if (target.directChildren.some((candidate) => containsNestedHandle(candidate, handle))) {
        throw new Error("reorderShapes: every shape must be a direct child of the target");
      }
      throw new Error("reorderShapes: shape handle was not found in the target drawing part");
    }
    orderedShapes.push(shape);
  }
  if (target.directChildren.some((shape) => shape.nodeId === undefined)) {
    throw new Error("reorderShapes: every target shape requires a node id");
  }

  const updated = withTargetShapes(source, target, orderedShapes);
  return {
    ...updated,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "reorderShapes",
        targetPartPath: target.partPath,
        ...(target.parentGroupId !== undefined
          ? { parentGroupId: String(target.parentGroupId) }
          : {}),
        shapeIds: orderedShapes.map((shape) => String(shape.nodeId)),
      },
    ],
  };
}

function assertUniqueDrawingNodeIds(shapes: readonly SourceShapeNode[]): void {
  const seen = new Set<string>();
  const visit = (nodes: readonly SourceShapeNode[]): void => {
    for (const node of nodes) {
      if (node.nodeId !== undefined) {
        const nodeId = String(node.nodeId);
        if (seen.has(nodeId)) {
          throw new Error(
            `reorderShapes: duplicate node id '${nodeId}' in the target drawing part is not supported`,
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
  if (shape.kind !== "group") return false;
  for (const child of shape.children) {
    if (sourceHandlesEqual(child.handle, handle)) return true;
    if (containsNestedHandle(child, handle)) return true;
  }
  return false;
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
    if (index >= 0) {
      const target = targets[index];
      return {
        kind,
        index,
        partPath: target.partPath,
        shapes: target.shapes,
        directChildren: target.shapes,
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
