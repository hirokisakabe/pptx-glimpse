import { asSourceNodeId } from "./handles.js";
import type {
  EditableTextRunProperties,
  EditableTextRunProperty,
  Emu,
  PartPath,
  PartRelationships,
  PptxSourceModel,
  PptxSourceModelEdit,
  RawPackagePart,
  Relationship,
  RelationshipId,
  SourceHandle,
  SourceParagraph,
  SourceRunProperties,
  SourceShape,
  SourceShapeNode,
  SourceSlide,
  SourceTextRun,
} from "./index.js";
import { asPartPath, asRelationshipId } from "./index.js";
import { relationshipsPartPath, resolveInternalRelationshipTarget } from "./package-paths.js";

type TransformableShapeNode = Exclude<SourceShapeNode, { readonly kind: "raw" }>;
type MutableRunProperties = {
  -readonly [K in keyof SourceRunProperties]?: SourceRunProperties[K];
};
type MutableEditableTextRunProperties = {
  -readonly [K in keyof EditableTextRunProperties]?: EditableTextRunProperties[K];
};

const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const NOTES_SLIDE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const NOTES_SLIDE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";

const EDITABLE_TEXT_RUN_PROPERTIES = [
  "bold",
  "italic",
  "underline",
  "fontSize",
  "color",
  "typeface",
] as const satisfies readonly EditableTextRunProperty[];
const EDITABLE_TEXT_RUN_PROPERTY_SET: ReadonlySet<string> = new Set(EDITABLE_TEXT_RUN_PROPERTIES);

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
  assertEditableTextRunPropertyNames(patch.clear);
  const set = definedEditableTextRunProperties(patch.set);
  if (Object.values(set).every((value) => value === undefined) && patch.clear.length === 0) {
    throw new Error("updateTextRunProperties: patch must set or clear at least one property");
  }

  let found = false;
  let changed = false;

  const slides = source.slides.map((slide) => ({
    ...slide,
    shapes: slide.shapes.map((shape) => {
      if (shape.kind !== "shape" || shape.textBody === undefined) return shape;

      let shapeChanged = false;
      const paragraphs = shape.textBody.paragraphs.map((paragraph) => {
        let paragraphChanged = false;
        const runs = paragraph.runs.map((run) => {
          if (!sourceHandlesEqual(run.handle, handle)) return run;
          found = true;
          const properties = patchTextRunProperties(run.properties, { set, clear: patch.clear });
          if (textRunPropertiesEqual(run.properties, properties)) return run;
          changed = true;
          paragraphChanged = true;
          shapeChanged = true;
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

  if (!found) {
    throw new Error(
      "updateTextRunProperties: text run handle was not found in PptxSourceModel source",
    );
  }
  if (!changed) return source;

  return {
    ...source,
    slides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "updateTextRunProperties",
        handle,
        ...(Object.keys(set).length > 0 ? { set } : {}),
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

export function duplicateSlide(
  source: PptxSourceModel,
  slideHandle: SourceHandle,
): PptxSourceModel {
  const slideIndex = source.slides.findIndex((slide) =>
    sourceHandlesEqual(slide.handle, slideHandle),
  );
  if (slideIndex === -1) {
    throw new Error("duplicateSlide: slide handle was not found in PptxSourceModel source");
  }

  const sourceSlide = source.slides[slideIndex];
  if (hasDirtyEditForPart(source.edits ?? [], sourceSlide.partPath)) {
    throw new Error(
      "duplicateSlide: duplicating a slide with pending dirty part edits is unsupported",
    );
  }

  const presentationRels = requirePartRelationships(
    source,
    source.presentation.partPath,
    "duplicateSlide",
  );
  const sourcePresentationRelationship = requireSlideRelationship(
    source,
    presentationRels,
    sourceSlide.partPath,
    "duplicateSlide",
  );
  const sourceRawSlide = requireRawBinaryPart(source, sourceSlide.partPath, "duplicateSlide");
  const sourceSlideRelationships = source.packageGraph.relationships.find(
    (relationships) => relationships.sourcePartPath === sourceSlide.partPath,
  );
  const newSlidePartPath = nextNumberedPartPath(source, "ppt/slides/slide", ".xml");
  const newPresentationRelationshipId = nextRelationshipId(presentationRels.relationships);
  const notesCopy = createNotesSlideCopy(
    source,
    sourceSlide,
    sourceSlideRelationships,
    newSlidePartPath,
  );
  const newSlideRelationships =
    sourceSlideRelationships === undefined
      ? undefined
      : {
          sourcePartPath: newSlidePartPath,
          relationships: sourceSlideRelationships.relationships.map((relationship) =>
            notesCopy !== undefined && relationship.id === notesCopy.slideRelationshipId
              ? { ...relationship, target: relativeTarget(newSlidePartPath, notesCopy.newPartPath) }
              : relationship,
          ),
        };

  const slideContentType =
    source.packageGraph.parts.find((part) => part.partPath === sourceSlide.partPath)?.contentType ??
    SLIDE_CONTENT_TYPE;
  const insertAt = slideIndex + 1;
  const newSlide = withPartPath(cloneJson(sourceSlide), newSlidePartPath);

  return {
    ...source,
    presentation: {
      ...source.presentation,
      slidePartPaths: insertAtReadonly(
        source.presentation.slidePartPaths,
        insertAt,
        newSlidePartPath,
      ),
    },
    slides: insertAtReadonly(source.slides, insertAt, newSlide),
    packageGraph: {
      ...source.packageGraph,
      parts: [
        ...source.packageGraph.parts,
        { partPath: newSlidePartPath, contentType: slideContentType },
        ...(newSlideRelationships === undefined
          ? []
          : [
              {
                partPath: asPartPath(relationshipsPartPath(newSlidePartPath)),
                contentType: RELS_CONTENT_TYPE,
              },
            ]),
        ...(notesCopy === undefined
          ? []
          : [
              { partPath: notesCopy.newPartPath, contentType: notesCopy.contentType },
              ...(notesCopy.relationships === undefined
                ? []
                : [
                    {
                      partPath: asPartPath(relationshipsPartPath(notesCopy.newPartPath)),
                      contentType: RELS_CONTENT_TYPE,
                    },
                  ]),
            ]),
      ],
      contentTypes: {
        ...source.packageGraph.contentTypes,
        overrides: [
          ...source.packageGraph.contentTypes.overrides,
          { partName: newSlidePartPath, contentType: slideContentType },
          ...(notesCopy === undefined
            ? []
            : [{ partName: notesCopy.newPartPath, contentType: notesCopy.contentType }]),
        ],
      },
      relationships: [
        ...source.packageGraph.relationships.map((relationships) =>
          relationships.sourcePartPath !== source.presentation.partPath
            ? relationships
            : {
                ...relationships,
                relationships: [
                  ...relationships.relationships,
                  {
                    id: newPresentationRelationshipId,
                    type: SLIDE_REL_TYPE,
                    target: relativeTarget(source.presentation.partPath, newSlidePartPath),
                  },
                ],
              },
        ),
        ...(newSlideRelationships === undefined ? [] : [newSlideRelationships]),
        ...(notesCopy?.relationships === undefined ? [] : [notesCopy.relationships]),
      ],
      rawParts: [
        ...(source.packageGraph.rawParts ?? []),
        {
          kind: "binary",
          partPath: newSlidePartPath,
          contentType: slideContentType,
          bytes: copyBytes(sourceRawSlide.bytes),
        },
        ...(notesCopy === undefined
          ? []
          : [
              {
                kind: "binary" as const,
                partPath: notesCopy.newPartPath,
                contentType: notesCopy.contentType,
                bytes: copyBytes(notesCopy.raw.bytes),
              },
            ]),
      ],
    },
    edits: [
      ...(source.edits ?? []),
      {
        kind: "duplicateSlide",
        sourceSlidePartPath: sourceSlide.partPath,
        sourceRelationshipId: sourcePresentationRelationship.id,
        newSlidePartPath,
        newRelationshipId: newPresentationRelationshipId,
      },
    ],
  };
}

export function deleteSlide(source: PptxSourceModel, slideHandle: SourceHandle): PptxSourceModel {
  if (source.slides.length <= 1) {
    throw new Error("deleteSlide: cannot delete the last slide");
  }

  const slideIndex = source.slides.findIndex((slide) =>
    sourceHandlesEqual(slide.handle, slideHandle),
  );
  if (slideIndex === -1) {
    throw new Error("deleteSlide: slide handle was not found in PptxSourceModel source");
  }

  const slide = source.slides[slideIndex];
  const presentationRels = requirePartRelationships(
    source,
    source.presentation.partPath,
    "deleteSlide",
  );
  const presentationRelationship = requireSlideRelationship(
    source,
    presentationRels,
    slide.partPath,
    "deleteSlide",
  );
  const slideRelationships = source.packageGraph.relationships.find(
    (relationships) => relationships.sourcePartPath === slide.partPath,
  );
  const notesPartPaths =
    slideRelationships?.relationships.flatMap((relationship) => {
      if (relationship.type !== NOTES_SLIDE_REL_TYPE) return [];
      const target = resolveInternalRelationshipTarget(slide.partPath, relationship);
      return target === undefined ? [] : [target];
    }) ?? [];
  const removedPartPaths = new Set<string>([slide.partPath, ...notesPartPaths]);
  const removedRelationshipPartPaths = new Set<string>(
    [slide.partPath, ...notesPartPaths].map((partPath) => relationshipsPartPath(partPath)),
  );
  const retainedEdits = (source.edits ?? []).filter(
    (edit) => !editIsInvalidatedByDeletedParts(edit, removedPartPaths),
  );
  const deletedInsertedSlide = (source.edits ?? []).some(
    (edit) => edit.kind === "duplicateSlide" && edit.newSlidePartPath === slide.partPath,
  );

  return {
    ...source,
    presentation: {
      ...source.presentation,
      slidePartPaths: source.presentation.slidePartPaths.filter(
        (partPath) => partPath !== slide.partPath,
      ),
    },
    slides: source.slides.filter((candidate) => candidate.partPath !== slide.partPath),
    packageGraph: {
      ...source.packageGraph,
      parts: source.packageGraph.parts.filter(
        (part) =>
          !removedPartPaths.has(part.partPath) && !removedRelationshipPartPaths.has(part.partPath),
      ),
      contentTypes: {
        ...source.packageGraph.contentTypes,
        overrides: source.packageGraph.contentTypes.overrides.filter(
          (override) => !removedPartPaths.has(override.partName),
        ),
      },
      relationships: source.packageGraph.relationships
        .filter((relationships) => !removedPartPaths.has(relationships.sourcePartPath))
        .map((relationships) =>
          relationships.sourcePartPath !== source.presentation.partPath
            ? relationships
            : {
                ...relationships,
                relationships: relationships.relationships.filter(
                  (relationship) => relationship.id !== presentationRelationship.id,
                ),
              },
        ),
      rawParts: source.packageGraph.rawParts?.filter(
        (part) => !removedPartPaths.has(part.partPath),
      ),
    },
    edits: deletedInsertedSlide
      ? retainedEdits
      : [
          ...retainedEdits,
          {
            kind: "deleteSlide",
            slidePartPath: slide.partPath,
            relationshipId: presentationRelationship.id,
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
  for (const property of Object.keys(properties)) {
    if (!EDITABLE_TEXT_RUN_PROPERTY_SET.has(property)) {
      throw new Error(`updateTextRunProperties: unsupported text run property '${property}'`);
    }
  }
  requireBooleanOrUndefined(properties.bold, "bold");
  requireBooleanOrUndefined(properties.italic, "italic");
  requireBooleanOrUndefined(properties.underline, "underline");
  if (
    properties.fontSize !== undefined &&
    (!Number.isFinite(properties.fontSize) || properties.fontSize <= 0)
  ) {
    throw new Error("updateTextRunProperties: fontSize must be a finite positive pt value");
  }
  if (properties.typeface !== undefined && properties.typeface.trim() === "") {
    throw new Error("updateTextRunProperties: typeface must be a non-empty string");
  }
  if (properties.color !== undefined) {
    if (properties.color.kind !== "srgb") {
      throw new Error("updateTextRunProperties: only srgb text run color is supported");
    }
    if (!/^[0-9A-Fa-f]{6}$/.test(properties.color.hex)) {
      throw new Error("updateTextRunProperties: srgb text run color must be a 6-digit hex value");
    }
  }
}

function requireBooleanOrUndefined(
  value: boolean | undefined,
  fieldName: "bold" | "italic" | "underline",
): void {
  if (value !== undefined && typeof value !== "boolean") {
    throw new Error(`updateTextRunProperties: ${fieldName} must be a boolean value`);
  }
}

function assertEditableTextRunPropertyNames(properties: readonly EditableTextRunProperty[]): void {
  for (const property of properties) {
    if (!EDITABLE_TEXT_RUN_PROPERTY_SET.has(property)) {
      throw new Error(`updateTextRunProperties: unsupported text run property '${property}'`);
    }
  }
}

function definedEditableTextRunProperties(
  properties: EditableTextRunProperties,
): EditableTextRunProperties {
  const defined: MutableEditableTextRunProperties = {};
  if (properties.bold !== undefined) defined.bold = properties.bold;
  if (properties.italic !== undefined) defined.italic = properties.italic;
  if (properties.underline !== undefined) defined.underline = properties.underline;
  if (properties.fontSize !== undefined) defined.fontSize = properties.fontSize;
  if (properties.color !== undefined) defined.color = properties.color;
  if (properties.typeface !== undefined) defined.typeface = properties.typeface;
  return defined;
}

function textRunPropertiesEqual(
  left: SourceRunProperties | undefined,
  right: SourceRunProperties | undefined,
): boolean {
  return JSON.stringify(left ?? {}) === JSON.stringify(right ?? {});
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

interface NotesSlideCopy {
  readonly slideRelationshipId: RelationshipId;
  readonly newPartPath: PartPath;
  readonly contentType: string;
  readonly raw: Extract<RawPackagePart, { readonly kind: "binary" }>;
  readonly relationships?: PartRelationships;
}

function createNotesSlideCopy(
  source: PptxSourceModel,
  sourceSlide: SourceSlide,
  slideRelationships: PartRelationships | undefined,
  newSlidePartPath: PartPath,
): NotesSlideCopy | undefined {
  const notesRelationship = slideRelationships?.relationships.find(
    (relationship) => relationship.type === NOTES_SLIDE_REL_TYPE,
  );
  if (notesRelationship === undefined) return undefined;
  const notesPartPath = resolveInternalRelationshipTarget(sourceSlide.partPath, notesRelationship);
  if (notesPartPath === undefined) return undefined;

  const raw = requireRawBinaryPart(source, notesPartPath, "duplicateSlide");
  const contentType =
    source.packageGraph.parts.find((part) => part.partPath === notesPartPath)?.contentType ??
    NOTES_SLIDE_CONTENT_TYPE;
  const newPartPath = nextNumberedPartPath(source, "ppt/notesSlides/notesSlide", ".xml");
  const notesRelationships = source.packageGraph.relationships.find(
    (relationships) => relationships.sourcePartPath === notesPartPath,
  );

  // Notes slides are 1:1 with a slide, so their slide back-reference must move
  // to the duplicated slide. Other notes relationships remain shared.
  return {
    slideRelationshipId: notesRelationship.id,
    newPartPath,
    contentType,
    raw,
    ...(notesRelationships === undefined
      ? {}
      : {
          relationships: {
            sourcePartPath: newPartPath,
            relationships: notesRelationships.relationships.map((relationship) =>
              relationship.type === SLIDE_REL_TYPE && relationship.targetMode !== "External"
                ? { ...relationship, target: relativeTarget(newPartPath, newSlidePartPath) }
                : relationship,
            ),
          },
        }),
  };
}

function requirePartRelationships(
  source: PptxSourceModel,
  partPath: PartPath,
  operationName: "duplicateSlide" | "deleteSlide",
): PartRelationships {
  const relationships = source.packageGraph.relationships.find(
    (candidate) => candidate.sourcePartPath === partPath,
  );
  if (relationships === undefined) {
    throw new Error(`${operationName}: presentation relationships were not found`);
  }
  return relationships;
}

function requireSlideRelationship(
  source: PptxSourceModel,
  relationships: PartRelationships,
  slidePartPath: PartPath,
  operationName: "duplicateSlide" | "deleteSlide",
): Relationship {
  const relationship = relationships.relationships.find(
    (candidate) =>
      candidate.type === SLIDE_REL_TYPE &&
      candidate.targetMode !== "External" &&
      resolveInternalRelationshipTarget(source.presentation.partPath, candidate) === slidePartPath,
  );
  if (relationship === undefined) {
    throw new Error(`${operationName}: slide relationship was not found in presentation.xml.rels`);
  }
  return relationship;
}

function requireRawBinaryPart(
  source: PptxSourceModel,
  partPath: PartPath,
  operationName: "duplicateSlide",
): Extract<RawPackagePart, { readonly kind: "binary" }> {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === partPath);
  if (rawPart === undefined) {
    throw new Error(`${operationName}: part '${partPath}' has no preserved raw package material`);
  }
  if (rawPart.kind !== "binary") {
    throw new Error(
      `${operationName}: part '${partPath}' is not backed by binary package material`,
    );
  }
  return rawPart;
}

function nextRelationshipId(relationships: readonly Relationship[]): RelationshipId {
  const used = new Set(relationships.map((relationship) => relationship.id));
  const max = relationships.reduce((current, relationship) => {
    const match = /(\d+)$/.exec(relationship.id);
    return match === null ? current : Math.max(current, Number(match[1]));
  }, 0);
  for (let index = max + 1; ; index += 1) {
    const candidate = asRelationshipId(`rId${index}`);
    if (!used.has(candidate)) return candidate;
  }
}

function nextNumberedPartPath(source: PptxSourceModel, prefix: string, suffix: string): PartPath {
  const used = new Set<string>([
    ...source.packageGraph.parts.map((part) => part.partPath),
    ...source.packageGraph.contentTypes.overrides.map((override) => override.partName),
    ...(source.packageGraph.rawParts ?? []).map((part) => part.partPath),
    ...(source.edits?.flatMap((edit) => topologyEditPartPaths(edit)) ?? []),
  ]);
  const pattern = new RegExp(`^${escapeRegExp(prefix)}(\\d+)${escapeRegExp(suffix)}$`);
  const max = [...used].reduce((current, path) => {
    const match = pattern.exec(path);
    return match === null ? current : Math.max(current, Number(match[1]));
  }, 0);
  for (let index = max + 1; ; index += 1) {
    const candidate = asPartPath(`${prefix}${index}${suffix}`);
    if (!used.has(candidate)) return candidate;
  }
}

function topologyEditPartPaths(edit: PptxSourceModelEdit): readonly PartPath[] {
  switch (edit.kind) {
    case "duplicateSlide":
      return [edit.sourceSlidePartPath, edit.newSlidePartPath];
    case "deleteSlide":
      return [edit.slidePartPath];
    case "replaceTextRunPlainText":
    case "updateTextRunProperties":
    case "replaceParagraphPlainText":
    case "updateShapeTransform":
      return [];
  }
}

function relativeTarget(sourcePartPath: PartPath, targetPartPath: PartPath): string {
  const sourceDir = sourcePartPath.split("/").slice(0, -1);
  const targetSegments = targetPartPath.split("/");
  while (sourceDir.length > 0 && targetSegments.length > 0 && sourceDir[0] === targetSegments[0]) {
    sourceDir.shift();
    targetSegments.shift();
  }
  return [...sourceDir.map(() => ".."), ...targetSegments].join("/");
}

function withPartPath(slide: SourceSlide, partPath: PartPath): SourceSlide {
  return {
    ...slide,
    partPath,
    handle: { partPath },
    shapes: slide.shapes.map((shape) => withShapePartPath(shape, partPath)),
  };
}

function withShapePartPath(shape: SourceShapeNode, partPath: PartPath): SourceShapeNode {
  if (shape.kind === "raw") return shape;
  const handle = shape.handle === undefined ? undefined : { ...shape.handle, partPath };
  if (shape.kind === "group") {
    return {
      ...shape,
      ...(handle !== undefined ? { handle } : {}),
      children: shape.children.map((child) => withShapePartPath(child, partPath)),
    };
  }
  if (shape.kind !== "shape" || shape.textBody === undefined) {
    return { ...shape, ...(handle !== undefined ? { handle } : {}) };
  }
  return {
    ...shape,
    ...(handle !== undefined ? { handle } : {}),
    textBody: {
      ...shape.textBody,
      paragraphs: shape.textBody.paragraphs.map((paragraph) => ({
        ...paragraph,
        ...(paragraph.handle !== undefined ? { handle: { ...paragraph.handle, partPath } } : {}),
        runs: paragraph.runs.map((run) => ({
          ...run,
          ...(run.handle !== undefined ? { handle: { ...run.handle, partPath } } : {}),
        })),
      })),
    },
  };
}

function cloneJson<T>(value: T): T {
  return structuredClone(value);
}

function copyBytes(bytes: Uint8Array): Uint8Array {
  return new Uint8Array(bytes);
}

function insertAtReadonly<T>(items: readonly T[], index: number, item: T): readonly T[] {
  return [...items.slice(0, index), item, ...items.slice(index)];
}

function hasDirtyEditForPart(edits: readonly PptxSourceModelEdit[], partPath: PartPath): boolean {
  return edits.some((edit) => {
    switch (edit.kind) {
      case "replaceTextRunPlainText":
      case "updateTextRunProperties":
      case "replaceParagraphPlainText":
      case "updateShapeTransform":
        return edit.handle.partPath === partPath;
      case "duplicateSlide":
      case "deleteSlide":
        return false;
    }
  });
}

function editIsInvalidatedByDeletedParts(
  edit: PptxSourceModelEdit,
  partPaths: ReadonlySet<string>,
): boolean {
  switch (edit.kind) {
    case "replaceTextRunPlainText":
    case "updateTextRunProperties":
    case "replaceParagraphPlainText":
    case "updateShapeTransform":
      return partPaths.has(edit.handle.partPath);
    case "duplicateSlide":
      return partPaths.has(edit.newSlidePartPath);
    case "deleteSlide":
      return partPaths.has(edit.slidePartPath);
  }
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
