import { asSourceNodeId } from "./handles.js";
import type {
  ConnectorPresetGeometry,
  EditableTextRunProperties,
  EditableTextRunProperty,
  Emu,
  PartPath,
  PartRelationships,
  PptxSourceModel,
  PptxSourceModelAddConnectorEdit,
  PptxSourceModelEdit,
  RawPackagePart,
  Relationship,
  RelationshipId,
  SourceArrowEndpoint,
  SourceConnector,
  SourceHandle,
  SourceNodeId,
  SourceOutline,
  SourceParagraph,
  SourceRunProperties,
  SourceShape,
  SourceShapeNode,
  SourceSlide,
  SourceTextRun,
  SourceTransform,
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
const SLIDE_LAYOUT_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const NOTES_SLIDE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const EMPTY_SLIDE_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
  `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
  `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
  `<p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
  `<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
  `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

const EDITABLE_TEXT_RUN_PROPERTIES = [
  "bold",
  "italic",
  "underline",
  "fontSize",
  "color",
  "typeface",
] as const satisfies readonly EditableTextRunProperty[];
const EDITABLE_TEXT_RUN_PROPERTY_SET: ReadonlySet<string> = new Set(EDITABLE_TEXT_RUN_PROPERTIES);
const CONNECTOR_PRESETS: ReadonlySet<ConnectorPresetGeometry> = new Set([
  "straightConnector1",
  "bentConnector3",
  "curvedConnector3",
]);
const ARROW_TYPES = new Set(["triangle", "stealth", "diamond", "oval", "arrow"]);
const ARROW_SIZES = new Set(["sm", "med", "lg"]);
const DEFAULT_CONNECTOR_OUTLINE: SourceOutline = {
  fill: { kind: "solid", color: { kind: "srgb", hex: "000000" } },
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

export interface AddTextBoxInput extends UpdateShapeTransformInput {
  readonly text: string;
  readonly name?: string;
}

export interface AddConnectorConnectionEndpointInput {
  readonly shapeHandle: SourceHandle;
  readonly connectionSiteIndex: number;
}

export interface AddConnectorOutlineInput {
  readonly headEnd?: SourceArrowEndpoint;
  readonly tailEnd?: SourceArrowEndpoint;
}

export interface AddConnectorInput extends UpdateShapeTransformInput {
  readonly preset: ConnectorPresetGeometry;
  readonly start: AddConnectorConnectionEndpointInput;
  readonly end: AddConnectorConnectionEndpointInput;
  readonly name?: string;
  readonly outline?: AddConnectorOutlineInput;
}

export interface AddEmptySlideFromLayoutInput {
  readonly layoutPartPath: PartPath;
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

export function addTextBox(
  source: PptxSourceModel,
  slideHandle: SourceHandle,
  input: AddTextBoxInput,
): PptxSourceModel {
  assertTextBoxInput(input);
  const slideIndex = source.slides.findIndex((slide) =>
    sourceHandlesEqual(slide.handle, slideHandle),
  );
  if (slideIndex === -1) {
    throw new Error("addTextBox: slide handle was not found in PptxSourceModel source");
  }

  const slide = source.slides[slideIndex];
  const shapeId = nextShapeId(slide.shapes, source.edits ?? [], slide.partPath);
  const shapeIdValue = String(shapeId);
  const name = input.name?.trim() || `TextBox ${shapeIdValue}`;
  const orderingSlot = nextOrderingSlot(slide.shapes);
  const shape = createTextBoxShape(slide.partPath, shapeId, name, orderingSlot, input);
  const slides = source.slides.map((candidate, index) =>
    index === slideIndex
      ? {
          ...candidate,
          shapes: [...candidate.shapes, shape],
        }
      : candidate,
  );

  return {
    ...source,
    slides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "addTextBox",
        slidePartPath: slide.partPath,
        shapeId: shapeIdValue,
        name,
        offsetX: input.offsetX,
        offsetY: input.offsetY,
        width: input.width,
        height: input.height,
        text: input.text,
      },
    ],
  };
}

export function addConnector(
  source: PptxSourceModel,
  slideHandle: SourceHandle,
  input: AddConnectorInput,
): PptxSourceModel {
  assertConnectorInput(input);
  const slideIndex = source.slides.findIndex((slide) =>
    sourceHandlesEqual(slide.handle, slideHandle),
  );
  if (slideIndex === -1) {
    throw new Error("addConnector: slide handle was not found in PptxSourceModel source");
  }

  const slide = source.slides[slideIndex];
  const startShape = requireConnectorTargetShape(slide, input.start.shapeHandle, "start");
  const endShape = requireConnectorTargetShape(slide, input.end.shapeHandle, "end");
  const shapeId = nextShapeId(slide.shapes, source.edits ?? [], slide.partPath);
  const shapeIdValue = String(shapeId);
  const name = input.name?.trim() || `Connector ${shapeIdValue}`;
  const orderingSlot = nextOrderingSlot(slide.shapes);
  const connector = createConnectorShape(
    slide.partPath,
    shapeId,
    name,
    orderingSlot,
    startShape,
    endShape,
    input,
  );
  const slides = source.slides.map((candidate, index) =>
    index === slideIndex
      ? {
          ...candidate,
          shapes: [...candidate.shapes, connector],
        }
      : candidate,
  );

  return {
    ...source,
    slides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "addConnector",
        slidePartPath: slide.partPath,
        shapeId: shapeIdValue,
        name,
        preset: input.preset,
        offsetX: input.offsetX,
        offsetY: input.offsetY,
        width: input.width,
        height: input.height,
        startShapeId: String(startShape.nodeId),
        startConnectionSiteIndex: input.start.connectionSiteIndex,
        endShapeId: String(endShape.nodeId),
        endConnectionSiteIndex: input.end.connectionSiteIndex,
        ...(input.outline !== undefined ? { outline: input.outline } : {}),
      } satisfies PptxSourceModelAddConnectorEdit,
    ],
  };
}

export function addEmptySlideFromLayout(
  source: PptxSourceModel,
  input: AddEmptySlideFromLayoutInput,
): PptxSourceModel {
  const layout = source.slideLayouts.find(
    (candidate) => candidate.partPath === input.layoutPartPath,
  );
  if (layout === undefined) {
    throw new Error("addEmptySlideFromLayout: slide layout part path was not found");
  }

  const presentationRels = requirePartRelationships(
    source,
    source.presentation.partPath,
    "addEmptySlideFromLayout",
  );
  const newSlidePartPath = nextNumberedPartPath(source, "ppt/slides/slide", ".xml");
  const newPresentationRelationshipId = nextRelationshipId(presentationRels.relationships);
  const newSlideRelationshipsPartPath = asPartPath(relationshipsPartPath(newSlidePartPath));
  const newSlide: SourceSlide = {
    partPath: newSlidePartPath,
    layoutPartPath: layout.partPath,
    shapes: [],
    handle: { partPath: newSlidePartPath },
  };
  const newSlideRelationships: PartRelationships = {
    sourcePartPath: newSlidePartPath,
    relationships: [
      {
        id: asRelationshipId("rId1"),
        type: SLIDE_LAYOUT_REL_TYPE,
        target: relativeTarget(newSlidePartPath, layout.partPath),
      },
    ],
  };

  return {
    ...source,
    presentation: {
      ...source.presentation,
      slidePartPaths: [...source.presentation.slidePartPaths, newSlidePartPath],
    },
    slides: [...source.slides, newSlide],
    packageGraph: {
      ...source.packageGraph,
      parts: [
        ...source.packageGraph.parts,
        { partPath: newSlidePartPath, contentType: SLIDE_CONTENT_TYPE },
        { partPath: newSlideRelationshipsPartPath, contentType: RELS_CONTENT_TYPE },
      ],
      contentTypes: {
        ...source.packageGraph.contentTypes,
        overrides: [
          ...source.packageGraph.contentTypes.overrides,
          { partName: newSlidePartPath, contentType: SLIDE_CONTENT_TYPE },
          ...relationshipPartOverrides(source, [newSlideRelationshipsPartPath]),
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
        newSlideRelationships,
      ],
      rawParts: [
        ...(source.packageGraph.rawParts ?? []),
        {
          kind: "binary",
          partPath: newSlidePartPath,
          contentType: SLIDE_CONTENT_TYPE,
          bytes: new TextEncoder().encode(EMPTY_SLIDE_XML),
        },
      ],
    },
    edits: [
      ...(source.edits ?? []),
      {
        kind: "addEmptySlideFromLayout",
        layoutPartPath: layout.partPath,
        newSlidePartPath,
        newRelationshipId: newPresentationRelationshipId,
      },
    ],
  };
}

export function deleteShape(source: PptxSourceModel, handle: SourceHandle): PptxSourceModel {
  if (handle.nodeId === undefined) {
    throw new Error("deleteShape: shape delete requires a node id");
  }

  let found: SourceShapeNode | undefined;
  let deleted = false;
  const slides = source.slides.map((slide) => {
    let slideChanged = false;
    const nextShapes = slide.shapes.filter((shape) => {
      if (!sourceHandlesEqual(shape.handle, handle)) return true;
      found = shape;
      if (shape.kind !== "shape") {
        throw new Error("deleteShape: only top-level sp shapes can be deleted");
      }
      if (hasAlternateContentSidecar(shape)) {
        throw new Error("deleteShape: shapes inside AlternateContent are not supported");
      }
      deleted = true;
      slideChanged = true;
      return false;
    });
    return slideChanged ? { ...slide, shapes: nextShapes } : slide;
  });

  if (!deleted) {
    if (
      found === undefined &&
      source.slides.some((slide) => hasNestedShapeNodeWithHandle(slide.shapes, handle))
    ) {
      throw new Error("deleteShape: nested group shape deletion is not supported");
    }
    throw new Error("deleteShape: shape handle was not found in PptxSourceModel source");
  }

  const referencingConnector = findConnectorReferencingShape(source, handle);
  if (referencingConnector !== undefined) {
    throw new Error(
      `deleteShape: shape is referenced by connector '${referencingConnector.name ?? referencingConnector.nodeId ?? "unknown"}'`,
    );
  }

  const retainedEdits = (source.edits ?? []).filter((edit) => !editTargetsShape(edit, handle));
  const deletedInsertedShape = (source.edits ?? []).some(
    (edit) =>
      edit.kind === "addTextBox" &&
      edit.slidePartPath === handle.partPath &&
      edit.shapeId === String(handle.nodeId),
  );

  return {
    ...source,
    slides,
    edits: deletedInsertedShape
      ? retainedEdits
      : [
          ...retainedEdits,
          {
            kind: "deleteShape",
            handle,
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
  const newRelationshipPartPaths = [
    ...(newSlideRelationships === undefined
      ? []
      : [asPartPath(relationshipsPartPath(newSlidePartPath))]),
    ...(notesCopy?.relationships === undefined
      ? []
      : [asPartPath(relationshipsPartPath(notesCopy.newPartPath))]),
  ];

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
          ...relationshipPartOverrides(source, newRelationshipPartPaths),
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
    (edit) =>
      (edit.kind === "addEmptySlideFromLayout" || edit.kind === "duplicateSlide") &&
      edit.newSlidePartPath === slide.partPath,
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
          (override) =>
            !removedPartPaths.has(override.partName) &&
            !removedRelationshipPartPaths.has(override.partName),
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

function assertTextBoxInput(input: AddTextBoxInput): void {
  assertFiniteEmu(input.offsetX, "addTextBox", "offsetX");
  assertFiniteEmu(input.offsetY, "addTextBox", "offsetY");
  assertPositiveFiniteEmu(input.width, "addTextBox", "width");
  assertPositiveFiniteEmu(input.height, "addTextBox", "height");
  if (typeof input.text !== "string") {
    throw new Error("addTextBox: text must be a string");
  }
  if (input.name !== undefined && input.name.trim() === "") {
    throw new Error("addTextBox: name must be a non-empty string when provided");
  }
}

function assertConnectorInput(input: AddConnectorInput): void {
  assertFiniteEmu(input.offsetX, "addConnector", "offsetX");
  assertFiniteEmu(input.offsetY, "addConnector", "offsetY");
  assertPositiveFiniteEmu(input.width, "addConnector", "width");
  assertPositiveFiniteEmu(input.height, "addConnector", "height");
  if (!CONNECTOR_PRESETS.has(input.preset)) {
    throw new Error(
      "addConnector: preset must be straightConnector1, bentConnector3, or curvedConnector3",
    );
  }
  assertConnectionSiteIndex(input.start.connectionSiteIndex, "start");
  assertConnectionSiteIndex(input.end.connectionSiteIndex, "end");
  if (input.name !== undefined && input.name.trim() === "") {
    throw new Error("addConnector: name must be a non-empty string when provided");
  }
  assertArrowEndpoint(input.outline?.headEnd, "headEnd");
  assertArrowEndpoint(input.outline?.tailEnd, "tailEnd");
}

function assertConnectionSiteIndex(value: number, fieldName: "start" | "end"): void {
  if (!Number.isInteger(value) || value < 0) {
    throw new Error(
      `addConnector: ${fieldName}.connectionSiteIndex must be a non-negative integer`,
    );
  }
}

function assertArrowEndpoint(endpoint: SourceArrowEndpoint | undefined, fieldName: string): void {
  if (endpoint === undefined) return;
  if (!ARROW_TYPES.has(endpoint.type)) {
    throw new Error(`addConnector: outline.${fieldName}.type is not supported`);
  }
  if (!ARROW_SIZES.has(endpoint.width)) {
    throw new Error(`addConnector: outline.${fieldName}.width is not supported`);
  }
  if (!ARROW_SIZES.has(endpoint.length)) {
    throw new Error(`addConnector: outline.${fieldName}.length is not supported`);
  }
}

function assertFiniteEmu(
  value: Emu,
  operationName: "addTextBox" | "addConnector",
  fieldName: string,
): void {
  if (!Number.isFinite(value)) {
    throw new Error(`${operationName}: ${fieldName} must be a finite EMU value`);
  }
}

function assertPositiveFiniteEmu(
  value: Emu,
  operationName: "addTextBox" | "addConnector",
  fieldName: string,
): void {
  if (!Number.isFinite(value) || value <= 0) {
    throw new Error(`${operationName}: ${fieldName} must be a finite positive EMU value`);
  }
}

function createTextBoxShape(
  partPath: PartPath,
  shapeId: SourceNodeId,
  name: string,
  orderingSlot: number,
  input: AddTextBoxInput,
): SourceShape {
  const transform: SourceTransform = {
    offsetX: input.offsetX,
    offsetY: input.offsetY,
    width: input.width,
    height: input.height,
  };
  return {
    kind: "shape",
    nodeId: shapeId,
    name,
    transform,
    geometry: { preset: "rect" },
    textBody: {
      handle: { partPath, nodeId: asSourceNodeId(`text:shape:${shapeId}`), orderingSlot },
      paragraphs: [
        {
          handle: {
            partPath,
            nodeId: asSourceNodeId(`text:shape:${shapeId}:p:0`),
            orderingSlot: 0,
          },
          runs: [
            {
              kind: "textRun",
              text: input.text,
              handle: {
                partPath,
                nodeId: asSourceNodeId(`text:shape:${shapeId}:p:0:r:0`),
                orderingSlot: 0,
              },
            },
          ],
        },
      ],
    },
    handle: { partPath, nodeId: shapeId, orderingSlot },
  };
}

function createConnectorShape(
  partPath: PartPath,
  shapeId: SourceNodeId,
  name: string,
  orderingSlot: number,
  startShape: SourceShape & { readonly nodeId: SourceNodeId },
  endShape: SourceShape & { readonly nodeId: SourceNodeId },
  input: AddConnectorInput,
): SourceConnector {
  const transform: SourceTransform = {
    offsetX: input.offsetX,
    offsetY: input.offsetY,
    width: input.width,
    height: input.height,
  };
  return {
    kind: "connector",
    nodeId: shapeId,
    name,
    connection: {
      start: {
        shapeId: startShape.nodeId,
        connectionSiteIndex: input.start.connectionSiteIndex,
      },
      end: {
        shapeId: endShape.nodeId,
        connectionSiteIndex: input.end.connectionSiteIndex,
      },
    },
    transform,
    geometry: { preset: input.preset },
    outline: createConnectorOutline(input.outline),
    handle: { partPath, nodeId: shapeId, orderingSlot },
  };
}

function createConnectorOutline(input: AddConnectorOutlineInput | undefined): SourceOutline {
  return {
    ...DEFAULT_CONNECTOR_OUTLINE,
    ...(input?.headEnd !== undefined ? { headEnd: input.headEnd } : {}),
    ...(input?.tailEnd !== undefined ? { tailEnd: input.tailEnd } : {}),
  };
}

function requireConnectorTargetShape(
  slide: SourceSlide,
  handle: SourceHandle,
  endpointName: "start" | "end",
): SourceShape & { readonly nodeId: SourceNodeId } {
  const shape = slide.shapes.find((candidate) => sourceHandlesEqual(candidate.handle, handle));
  if (shape === undefined) {
    throw new Error(`addConnector: ${endpointName} shape handle was not found on the target slide`);
  }
  if (shape.kind !== "shape") {
    throw new Error(`addConnector: ${endpointName} target must be a top-level sp shape`);
  }
  const nodeId = shape.nodeId;
  if (nodeId === undefined) {
    throw new Error(`addConnector: ${endpointName} target shape requires a node id`);
  }
  return { ...shape, nodeId };
}

function nextShapeId(
  shapes: readonly SourceShapeNode[],
  edits: readonly PptxSourceModelEdit[],
  slidePartPath: PartPath,
): SourceNodeId {
  const used = new Set<number>();
  collectNumericShapeIds(shapes, used);
  collectNumericShapeEditIds(edits, slidePartPath, used);
  const max = Math.max(0, ...used);
  for (let candidate = max + 1; ; candidate += 1) {
    if (!used.has(candidate)) return asSourceNodeId(String(candidate));
  }
}

function collectNumericShapeIds(shapes: readonly SourceShapeNode[], used: Set<number>): void {
  for (const shape of shapes) {
    const numericId = Number(shape.nodeId);
    if (Number.isInteger(numericId) && numericId > 0) used.add(numericId);
    if (shape.kind === "group") collectNumericShapeIds(shape.children, used);
  }
}

function collectNumericShapeEditIds(
  edits: readonly PptxSourceModelEdit[],
  slidePartPath: PartPath,
  used: Set<number>,
): void {
  for (const edit of edits) {
    const id =
      edit.kind === "addTextBox" && edit.slidePartPath === slidePartPath
        ? edit.shapeId
        : edit.kind === "addConnector" && edit.slidePartPath === slidePartPath
          ? edit.shapeId
          : edit.kind === "deleteShape" && edit.handle.partPath === slidePartPath
            ? edit.handle.nodeId
            : undefined;
    const numericId = Number(id);
    if (Number.isInteger(numericId) && numericId > 0) used.add(numericId);
  }
}

function nextOrderingSlot(shapes: readonly SourceShapeNode[]): number {
  return (
    shapes.reduce((current, shape) => {
      const slot = shape.handle?.orderingSlot ?? -1;
      return Math.max(current, slot);
    }, -1) + 1
  );
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

function findConnectorReferencingShape(
  source: PptxSourceModel,
  handle: SourceHandle,
): SourceConnector | undefined {
  if (handle.nodeId === undefined) return undefined;
  for (const slide of source.slides) {
    if (slide.partPath !== handle.partPath) continue;
    for (const shape of slide.shapes) {
      if (shape.kind !== "connector") continue;
      if (
        shape.connection?.start?.shapeId === handle.nodeId ||
        shape.connection?.end?.shapeId === handle.nodeId
      ) {
        return shape;
      }
    }
  }
  return undefined;
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
  operationName: "addEmptySlideFromLayout" | "duplicateSlide" | "deleteSlide",
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
  operationName: "addEmptySlideFromLayout" | "duplicateSlide" | "deleteSlide",
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

function relationshipPartOverrides(
  source: PptxSourceModel,
  partPaths: readonly PartPath[],
): readonly { readonly partName: PartPath; readonly contentType: string }[] {
  if (
    source.packageGraph.contentTypes.defaults.some(
      (entry) => entry.extension === "rels" && entry.contentType === RELS_CONTENT_TYPE,
    )
  ) {
    return [];
  }

  const existingOverrides = new Set(
    source.packageGraph.contentTypes.overrides.map((override) => override.partName),
  );
  return partPaths
    .filter((partPath) => !existingOverrides.has(partPath))
    .map((partName) => ({ partName, contentType: RELS_CONTENT_TYPE }));
}

function topologyEditPartPaths(edit: PptxSourceModelEdit): readonly PartPath[] {
  switch (edit.kind) {
    case "addEmptySlideFromLayout":
      return [edit.layoutPartPath, edit.newSlidePartPath];
    case "duplicateSlide":
      return [edit.sourceSlidePartPath, edit.newSlidePartPath];
    case "deleteSlide":
      return [edit.slidePartPath];
    case "addTextBox":
    case "addConnector":
      return [edit.slidePartPath];
    case "replaceTextRunPlainText":
    case "updateTextRunProperties":
    case "replaceParagraphPlainText":
    case "updateShapeTransform":
    case "deleteShape":
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
      case "deleteShape":
        return edit.handle.partPath === partPath;
      case "addTextBox":
      case "addConnector":
        return edit.slidePartPath === partPath;
      case "addEmptySlideFromLayout":
      case "duplicateSlide":
      case "deleteSlide":
        return false;
    }
  });
}

function editTargetsShape(edit: PptxSourceModelEdit, handle: SourceHandle): boolean {
  switch (edit.kind) {
    case "replaceTextRunPlainText":
    case "updateTextRunProperties":
      return (
        edit.handle.partPath === handle.partPath && textRunShapeId(edit.handle) === handle.nodeId
      );
    case "replaceParagraphPlainText":
      return (
        edit.handle.partPath === handle.partPath && paragraphShapeId(edit.handle) === handle.nodeId
      );
    case "updateShapeTransform":
    case "deleteShape":
      return sourceHandlesEqual(edit.handle, handle);
    case "addTextBox":
    case "addConnector":
      return edit.slidePartPath === handle.partPath && edit.shapeId === String(handle.nodeId);
    case "addEmptySlideFromLayout":
    case "duplicateSlide":
    case "deleteSlide":
      return false;
  }
}

function textRunShapeId(handle: SourceHandle): string | undefined {
  return /^text:shape:([^:]+):p:\d+:r:\d+$/.exec(String(handle.nodeId ?? ""))?.[1];
}

function paragraphShapeId(handle: SourceHandle): string | undefined {
  return /^text:shape:([^:]+):p:\d+$/.exec(String(handle.nodeId ?? ""))?.[1];
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
    case "addTextBox":
    case "addConnector":
      return partPaths.has(edit.slidePartPath);
    case "addEmptySlideFromLayout":
      return partPaths.has(edit.newSlidePartPath);
    case "deleteShape":
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
