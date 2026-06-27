import type { SourceShape, SourceShapeNode } from "../source/index.js";
import type { ComputedPlaceholderMatch } from "./pptx-computed-view.js";

interface PlaceholderMatchContext {
  readonly layoutShapes: readonly SourceShapeNode[];
  readonly masterShapes: readonly SourceShapeNode[];
}

export function findPlaceholderMatch(
  context: PlaceholderMatchContext,
  shape: SourceShape,
): ComputedPlaceholderMatch | undefined {
  if (shape.placeholder === undefined) return undefined;
  const type = shape.placeholder.type ?? "body";
  const index = shape.placeholder.index;
  const layout = findMatchingPlaceholder(type, index, context.layoutShapes);
  const master = findMatchingPlaceholder(type, index, context.masterShapes);
  if (layout === undefined && master === undefined) return undefined;
  return {
    ...(layout !== undefined ? { layout } : {}),
    ...(master !== undefined ? { master } : {}),
  };
}

function findMatchingPlaceholder(
  type: string,
  index: number | undefined,
  shapes: readonly SourceShapeNode[],
): SourceShape | undefined {
  const placeholders = shapes.filter(
    (shape): shape is SourceShape => shape.kind === "shape" && shape.placeholder !== undefined,
  );

  if (index !== undefined) {
    const byIndexWithTransform = placeholders.find(
      (shape) => shape.placeholder?.index === index && shape.transform !== undefined,
    );
    if (byIndexWithTransform !== undefined) return byIndexWithTransform;
    const byIndex = placeholders.find((shape) => shape.placeholder?.index === index);
    if (byIndex !== undefined) return byIndex;
  }

  const byTypeWithTransform = placeholders.find(
    (shape) => (shape.placeholder?.type ?? "body") === type && shape.transform !== undefined,
  );
  if (byTypeWithTransform !== undefined) return byTypeWithTransform;

  const fallbackType = getPlaceholderFallbackType(type);
  if (fallbackType !== undefined) {
    const byFallbackWithTransform = placeholders.find(
      (shape) =>
        (shape.placeholder?.type ?? "body") === fallbackType && shape.transform !== undefined,
    );
    if (byFallbackWithTransform !== undefined) return byFallbackWithTransform;
  }

  const byType = placeholders.find((shape) => (shape.placeholder?.type ?? "body") === type);
  if (byType !== undefined) return byType;

  if (fallbackType !== undefined) {
    return placeholders.find((shape) => (shape.placeholder?.type ?? "body") === fallbackType);
  }

  return undefined;
}

function getPlaceholderFallbackType(type: string): string | undefined {
  if (type === "ctrTitle") return "title";
  if (type === "subTitle") return "body";
  return undefined;
}
