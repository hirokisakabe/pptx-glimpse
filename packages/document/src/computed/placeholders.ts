import type { SourcePlaceholder, SourceShapeNode } from "../source/index.js";
import type { ComputedPlaceholderMatch } from "./pptx-computed-view.js";

interface PlaceholderMatchContext {
  readonly layoutShapes: readonly SourceShapeNode[];
  readonly masterShapes: readonly SourceShapeNode[];
}

/** Effective schema default for `p:ph@type`. */
export function effectivePlaceholderType(placeholder: SourcePlaceholder): string {
  return placeholder.type ?? "obj";
}

/** Effective schema default for `p:ph@idx`. */
export function effectivePlaceholderIndex(placeholder: SourcePlaceholder): number {
  return placeholder.index ?? 0;
}

export function findPlaceholderMatch(
  context: PlaceholderMatchContext,
  shape: SourceShapeNode,
): ComputedPlaceholderMatch | undefined {
  const placeholder = sourceNodePlaceholder(shape);
  if (placeholder === undefined) return undefined;

  const layoutCandidates = placeholderNodes(context.layoutShapes).filter(
    (candidate) =>
      effectivePlaceholderIndex(candidate.placeholder) === effectivePlaceholderIndex(placeholder),
  );
  if (layoutCandidates.length !== 1) return undefined;

  const layout = layoutCandidates[0];
  const masterCandidates = placeholderNodes(context.masterShapes).filter((candidate) =>
    masterTypeMatchesLayout(
      effectivePlaceholderType(candidate.placeholder),
      effectivePlaceholderType(layout.placeholder),
    ),
  );
  const master = masterCandidates.length === 1 ? masterCandidates[0] : undefined;
  return {
    layoutNode: layout,
    ...(layout.kind === "shape" ? { layout } : {}),
    ...(master !== undefined ? { masterNode: master } : {}),
    ...(master?.kind === "shape" ? { master } : {}),
  };
}

function placeholderNodes(
  shapes: readonly SourceShapeNode[],
): (SourceShapeNode & { readonly placeholder: SourcePlaceholder })[] {
  return shapes.filter(hasSourceNodePlaceholder);
}

function hasSourceNodePlaceholder(
  shape: SourceShapeNode,
): shape is SourceShapeNode & { readonly placeholder: SourcePlaceholder } {
  return sourceNodePlaceholder(shape) !== undefined;
}

export function sourceNodePlaceholder(shape: SourceShapeNode): SourcePlaceholder | undefined {
  switch (shape.kind) {
    case "shape":
    case "image":
    case "table":
    case "chart":
    case "smartArt":
    case "raw":
      return shape.placeholder;
    case "connector":
    case "group":
      return undefined;
  }
}

function masterTypeMatchesLayout(masterType: string, layoutType: string): boolean {
  if (layoutType === "title" || layoutType === "ctrTitle") return masterType === "title";
  if (layoutType === "body" || layoutType === "subTitle" || layoutType === "obj") {
    return masterType === "body";
  }
  return masterType === layoutType;
}
