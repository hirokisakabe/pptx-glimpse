import { asSourceNodeId } from "./handles.js";
import type {
  EditableTextRunProperties,
  EditableTextRunProperty,
  Emu,
  PptxSourceModel,
  SourceHandle,
  SourceParagraph,
  SourceRunProperties,
  SourceShape,
  SourceShapeNode,
  SourceTextRun,
} from "./index.js";

type TransformableShapeNode = Exclude<SourceShapeNode, { readonly kind: "raw" }>;
type MutableRunProperties = {
  -readonly [K in keyof SourceRunProperties]?: SourceRunProperties[K];
};

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

export function findParagraphBySourceHandle(
  source: PptxSourceModel,
  handle: SourceHandle,
): SourceParagraph | undefined {
  for (const slide of source.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind !== "shape") continue;
      const paragraph = findParagraphInShape(shape, handle);
      if (paragraph !== undefined) return paragraph;
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

export function setTextRunProperties(
  source: PptxSourceModel,
  handle: SourceHandle,
  properties: EditableTextRunProperties,
): PptxSourceModel {
  return updateTextRunProperties(source, handle, {
    set: properties,
    clear: [],
  });
}

export function clearTextRunProperties(
  source: PptxSourceModel,
  handle: SourceHandle,
  properties: readonly EditableTextRunProperty[],
): PptxSourceModel {
  return updateTextRunProperties(source, handle, {
    set: {},
    clear: properties,
  });
}

export function replaceParagraphPlainText(
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
        if (!sourceHandlesEqual(paragraph.handle, handle)) return paragraph;
        const replacementHandle = createReplacementRunHandle(paragraph);
        replaced = true;
        shapeChanged = true;
        return {
          ...paragraph,
          runs: [
            {
              kind: "textRun",
              text,
              ...(paragraph.runs[0]?.properties !== undefined
                ? { properties: paragraph.runs[0].properties }
                : {}),
              ...(replacementHandle !== undefined ? { handle: replacementHandle } : {}),
            },
          ],
        } satisfies SourceParagraph;
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
      "replaceParagraphPlainText: paragraph handle was not found in PptxSourceModel source",
    );
  }

  return {
    ...source,
    slides,
    edits: [...(source.edits ?? []), { kind: "replaceParagraphPlainText", handle, text }],
  };
}

interface UpdateTextRunPropertiesPatch {
  readonly set: EditableTextRunProperties;
  readonly clear: readonly EditableTextRunProperty[];
}

function updateTextRunProperties(
  source: PptxSourceModel,
  handle: SourceHandle,
  patch: UpdateTextRunPropertiesPatch,
): PptxSourceModel {
  assertEditableTextRunProperties(patch.set);
  if (Object.values(patch.set).every((value) => value === undefined) && patch.clear.length === 0) {
    throw new Error("updateTextRunProperties: patch must set or clear at least one property");
  }

  let updated = false;

  const slides = source.slides.map((slide) => ({
    ...slide,
    shapes: slide.shapes.map((shape) => {
      if (shape.kind !== "shape" || shape.textBody === undefined) return shape;

      let shapeChanged = false;
      const paragraphs = shape.textBody.paragraphs.map((paragraph) => {
        let paragraphChanged = false;
        const runs = paragraph.runs.map((run) => {
          if (!sourceHandlesEqual(run.handle, handle)) return run;
          updated = true;
          paragraphChanged = true;
          shapeChanged = true;
          const properties = patchTextRunProperties(run.properties, patch);
          return {
            kind: run.kind,
            text: run.text,
            ...(run.handle !== undefined ? { handle: run.handle } : {}),
            ...(run.rawSidecars !== undefined ? { rawSidecars: run.rawSidecars } : {}),
            ...(properties !== undefined ? { properties } : {}),
          } satisfies SourceTextRun;
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

  if (!updated) {
    throw new Error(
      "updateTextRunProperties: text run handle was not found in PptxSourceModel source",
    );
  }

  return {
    ...source,
    slides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "updateTextRunProperties",
        handle,
        ...(Object.keys(patch.set).length > 0 ? { set: patch.set } : {}),
        ...(patch.clear.length > 0 ? { clear: patch.clear } : {}),
      },
    ],
  };
}

function patchTextRunProperties(
  current: SourceRunProperties | undefined,
  patch: UpdateTextRunPropertiesPatch,
): SourceRunProperties | undefined {
  const next: MutableRunProperties = { ...(current ?? {}) };
  for (const property of patch.clear) {
    delete next[property];
  }
  if (patch.set.bold !== undefined) next.bold = patch.set.bold;
  if (patch.set.italic !== undefined) next.italic = patch.set.italic;
  if (patch.set.underline !== undefined) next.underline = patch.set.underline;
  if (patch.set.fontSize !== undefined) next.fontSize = patch.set.fontSize;
  if (patch.set.color !== undefined) next.color = patch.set.color;
  if (patch.set.typeface !== undefined) next.typeface = patch.set.typeface;
  return Object.keys(next).length > 0 ? next : undefined;
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

function findParagraphInShape(
  shape: SourceShape,
  handle: SourceHandle,
): SourceParagraph | undefined {
  return shape.textBody?.paragraphs.find((paragraph) =>
    sourceHandlesEqual(paragraph.handle, handle),
  );
}

function assertEditableTextRunProperties(properties: EditableTextRunProperties): void {
  if (properties.color !== undefined) {
    if (properties.color.kind !== "srgb") {
      throw new Error("updateTextRunProperties: only srgb text run color is supported");
    }
    if (!/^[0-9A-Fa-f]{6}$/.test(properties.color.hex)) {
      throw new Error("updateTextRunProperties: srgb text run color must be a 6-digit hex value");
    }
  }
}

function createReplacementRunHandle(paragraph: SourceParagraph): SourceHandle | undefined {
  if (paragraph.runs[0]?.handle !== undefined) return paragraph.runs[0].handle;
  if (paragraph.handle?.nodeId === undefined) return undefined;
  return {
    ...paragraph.handle,
    nodeId: asSourceNodeId(`${paragraph.handle.nodeId}:r:0`),
    orderingSlot: 0,
  };
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
