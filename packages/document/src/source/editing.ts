import type {
  Emu,
  PptxSourceModel,
  SourceHandle,
  SourceParagraph,
  SourceShape,
  SourceShapeNode,
  SourceTextRun,
} from "./index.js";

type TransformableShapeNode = Exclude<SourceShapeNode, { readonly kind: "raw" }>;

export function findTextRunBySourceHandle(
  source: PptxSourceModel,
  handle: SourceHandle,
): SourceTextRun | undefined {
  for (const slide of source.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind !== "shape") continue;
      const run = findTextRunInShape(shape, handle);
      if (run !== undefined) return run;
    }
  }
  return undefined;
}

export function replaceTextRunPlainText(
  source: PptxSourceModel,
  handle: SourceHandle,
  text: string,
): PptxSourceModel {
  let replaced = false;

  const slides = source.slides.map((slide) => ({
    ...slide,
    shapes: slide.shapes.map((shape) => {
      if (shape.kind !== "shape" || shape.textBody === undefined) return shape;

      let shapeChanged = false;
      const paragraphs = shape.textBody.paragraphs.map((paragraph) => {
        let paragraphChanged = false;
        const runs = paragraph.runs.map((run) => {
          if (!sourceHandlesEqual(run.handle, handle)) return run;
          replaced = true;
          paragraphChanged = true;
          shapeChanged = true;
          return { ...run, text };
        });
        return !paragraphChanged ? paragraph : ({ ...paragraph, runs } satisfies SourceParagraph);
      });

      if (!shapeChanged) return shape;
      return {
        ...shape,
        textBody: {
          ...shape.textBody,
          paragraphs,
        },
      } satisfies SourceShape;
    }),
  }));

  if (!replaced) {
    throw new Error(
      "replaceTextRunPlainText: text run handle was not found in PptxSourceModel source",
    );
  }

  return {
    ...source,
    slides,
    edits: [...(source.edits ?? []), { kind: "replaceTextRunPlainText", handle, text }],
  };
}

export interface UpdateShapeTransformInput {
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
}

export function findShapeNodeBySourceHandle(
  source: PptxSourceModel,
  handle: SourceHandle,
): SourceShapeNode | undefined {
  for (const slide of source.slides) {
    const shape = findShapeNodeInTree(slide.shapes, handle);
    if (shape !== undefined) return shape;
  }
  return undefined;
}

export function updateShapeTransform(
  source: PptxSourceModel,
  handle: SourceHandle,
  transform: UpdateShapeTransformInput,
): PptxSourceModel {
  if (handle.nodeId === undefined) {
    throw new Error("updateShapeTransform: shape transform edit requires a node id");
  }

  let updated = false;

  const slides = source.slides.map((slide) => ({
    ...slide,
    shapes: slide.shapes.map((shape) => {
      if (!sourceHandlesEqual(shape.handle, handle)) return shape;
      if (hasAlternateContentSidecar(shape)) {
        throw new Error("updateShapeTransform: shapes inside AlternateContent are not supported");
      }
      if (!hasEditableTransform(shape)) {
        throw new Error("updateShapeTransform: shape handle does not reference a shape with xfrm");
      }
      updated = true;
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
    }),
  }));

  if (!updated) {
    if (source.slides.some((slide) => hasNestedShapeNodeWithHandle(slide.shapes, handle))) {
      throw new Error("updateShapeTransform: nested group shape editing is not supported");
    }
    throw new Error("updateShapeTransform: shape handle was not found in PptxSourceModel source");
  }

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

function findTextRunInShape(shape: SourceShape, handle: SourceHandle): SourceTextRun | undefined {
  for (const paragraph of shape.textBody?.paragraphs ?? []) {
    for (const run of paragraph.runs) {
      if (sourceHandlesEqual(run.handle, handle)) return run;
    }
  }
  return undefined;
}

function hasEditableTransform(shape: SourceShapeNode): shape is TransformableShapeNode & {
  readonly transform: NonNullable<TransformableShapeNode["transform"]>;
} {
  return shape.kind !== "raw" && shape.transform !== undefined;
}

function findShapeNodeInTree(
  shapes: readonly SourceShapeNode[],
  handle: SourceHandle,
): SourceShapeNode | undefined {
  for (const shape of shapes) {
    if (sourceHandlesEqual(shape.handle, handle)) return shape;
    if (shape.kind === "group") {
      const child = findShapeNodeInTree(shape.children, handle);
      if (child !== undefined) return child;
    }
  }
  return undefined;
}

function hasNestedShapeNodeWithHandle(
  shapes: readonly SourceShapeNode[],
  handle: SourceHandle,
): boolean {
  return shapes.some(
    (shape) =>
      shape.kind === "group" &&
      (findShapeNodeInTree(shape.children, handle) !== undefined ||
        hasNestedShapeNodeWithHandle(shape.children, handle)),
  );
}

function hasAlternateContentSidecar(shape: SourceShapeNode): boolean {
  if (shape.kind === "raw") return false;
  return shape.rawSidecars?.some((sidecar) => sidecar.node.name === "mc:AlternateContent") ?? false;
}

function sourceHandlesEqual(left: SourceHandle | undefined, right: SourceHandle): boolean {
  if (left === undefined) return false;
  return (
    left.partPath === right.partPath &&
    left.nodeId === right.nodeId &&
    left.relationshipId === right.relationshipId &&
    left.orderingSlot === right.orderingSlot
  );
}
