import { sourceHandlesEqual } from "./edit-descriptors.js";
import type { SourceHandle } from "./handles.js";
import type { PptxSourceModel } from "./pptx-source-model.js";
import type { SourceShapeNode } from "./shapes.js";

type Target = {
  kind: "slide" | "layout" | "master";
  index: number;
  partPath: SourceHandle["partPath"];
  shapes: readonly SourceShapeNode[];
};

/** Reorders every top-level drawing in one slide, layout, or master shape tree. */
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
  if (orderedShapeHandles.length !== target.shapes.length) {
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
    const shape = target.shapes.find((candidate) => sourceHandlesEqual(candidate.handle, handle));
    if (shape === undefined) {
      throw new Error("reorderShapes: shape handle was not found in the target drawing part");
    }
    orderedShapes.push(shape);
  }
  if (target.shapes.some((shape) => shape.nodeId === undefined)) {
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
        shapeIds: orderedShapes.map((shape) => String(shape.nodeId)),
      },
    ],
  };
}

function hasAlternateContent(shape: SourceShapeNode): boolean {
  if (shape.kind === "raw") return shape.raw.node.name === "mc:AlternateContent";
  return shape.rawSidecars?.some((sidecar) => sidecar.node.name === "mc:AlternateContent") ?? false;
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
      return { kind, index, partPath: target.partPath, shapes: target.shapes };
    }
  }
  return undefined;
}

function withTargetShapes(
  source: PptxSourceModel,
  target: Target,
  shapes: readonly SourceShapeNode[],
): PptxSourceModel {
  switch (target.kind) {
    case "slide":
      return {
        ...source,
        slides: source.slides.map((slide, index) =>
          index === target.index ? { ...slide, shapes } : slide,
        ),
      };
    case "layout":
      return {
        ...source,
        slideLayouts: source.slideLayouts.map((layout, index) =>
          index === target.index ? { ...layout, shapes } : layout,
        ),
      };
    case "master":
      return {
        ...source,
        slideMasters: source.slideMasters.map((master, index) =>
          index === target.index ? { ...master, shapes } : master,
        ),
      };
  }
}
