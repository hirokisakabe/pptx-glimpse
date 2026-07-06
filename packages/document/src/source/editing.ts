import { getAttr, getChild, getChildArray, parseXml } from "../reader/xml.js";
import {
  editDirtyPartPath,
  editInsertedShape,
  editInsertedSlidePartPath,
  editInvalidatingPartPaths,
  editReservedPartPaths,
  editReservedShapeId,
  editTargetsShape,
  sourceHandlesEqual,
} from "./edit-descriptors.js";
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
  PptxSourceModelReplaceImageEdit,
  RawPackagePart,
  Relationship,
  RelationshipId,
  SourceArrowEndpoint,
  SourceConnector,
  SourceHandle,
  SourceImage,
  SourceNodeId,
  SourceParagraph,
  SourceRunProperties,
  SourceShape,
  SourceShapeNode,
  SourceSlide,
  SourceTextRun,
} from "./index.js";
import { asRelationshipId } from "./index.js";
import {
  addPackagePart,
  addPartRelationship,
  nextNumberedName,
  nextNumberedPartPath,
  nextRelationshipId,
  removePackageParts,
  removePartRelationship,
} from "./package-graph-mutations.js";
import { resolveInternalRelationshipTarget } from "./package-paths.js";
import { buildConnectorXml, buildTextBoxXml, parseShapeNodeXml } from "./shape-xml.js";

type TransformableShapeNode = Exclude<SourceShapeNode, { readonly kind: "raw" }>;
type MutableRunProperties = {
  -readonly [K in keyof SourceRunProperties]?: SourceRunProperties[K];
};
type MutableEditableTextRunProperties = {
  -readonly [K in keyof EditableTextRunProperties]?: EditableTextRunProperties[K];
};

const textDecoder = new TextDecoder();

const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const NOTES_SLIDE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_LAYOUT_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const NOTES_SLIDE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
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
  const xml = buildTextBoxXml({
    shapeId: shapeIdValue,
    name,
    offsetX: input.offsetX,
    offsetY: input.offsetY,
    width: input.width,
    height: input.height,
    text: input.text,
  });
  const shape = parseShapeNodeXml(xml, slide.partPath, orderingSlot);
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
        xml,
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
  const startShapeId = String(startShape.nodeId);
  const endShapeId = String(endShape.nodeId);
  const xml = buildConnectorXml({
    shapeId: shapeIdValue,
    name,
    preset: input.preset,
    offsetX: input.offsetX,
    offsetY: input.offsetY,
    width: input.width,
    height: input.height,
    startShapeId,
    startConnectionSiteIndex: input.start.connectionSiteIndex,
    endShapeId,
    endConnectionSiteIndex: input.end.connectionSiteIndex,
    ...(input.outline !== undefined ? { outline: input.outline } : {}),
  });
  const connector = parseShapeNodeXml(xml, slide.partPath, orderingSlot);
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
        startShapeId,
        endShapeId,
        xml,
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
  const newSlidePartPath = nextNumberedPartPath(
    source.packageGraph,
    source.edits?.flatMap((edit) => editReservedPartPaths(edit)) ?? [],
    "ppt/slides/slide",
    ".xml",
  );
  const newPresentationRelationshipId = nextRelationshipId(presentationRels.relationships);
  const newSlideNumericId = nextSlideNumericId(source, "addEmptySlideFromLayout");
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
  const packageGraph = addPartRelationship(
    addPackagePart(source.packageGraph, {
      partPath: newSlidePartPath,
      contentType: SLIDE_CONTENT_TYPE,
      bytes: new TextEncoder().encode(EMPTY_SLIDE_XML),
      relationships: newSlideRelationships,
    }),
    source.presentation.partPath,
    presentationSlideRelationship(source, newPresentationRelationshipId, newSlidePartPath),
  );

  return {
    ...source,
    presentation: {
      ...source.presentation,
      slidePartPaths: [...source.presentation.slidePartPaths, newSlidePartPath],
    },
    slides: [...source.slides, newSlide],
    packageGraph,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "addEmptySlideFromLayout",
        layoutPartPath: layout.partPath,
        newSlidePartPath,
        newRelationshipId: newPresentationRelationshipId,
        newSlideNumericId,
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
  const deletedInsertedShape = (source.edits ?? []).some((edit) => {
    const inserted = editInsertedShape(edit);
    return (
      inserted !== undefined &&
      inserted.slidePartPath === handle.partPath &&
      inserted.shapeId === String(handle.nodeId)
    );
  });

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

export function replaceImageBytes(
  source: PptxSourceModel,
  handle: SourceHandle,
  bytes: Uint8Array,
): PptxSourceModel {
  const image = requireImageBySourceHandle(source, handle, "replaceImageBytes");
  const media = requireMediaForImage(source, image, "replaceImageBytes");
  const detectedContentType = detectImageContentType(bytes);
  if (detectedContentType === undefined) {
    throw new Error("replaceImageBytes: unsupported or unknown replacement image format");
  }
  if (detectedContentType !== media.contentType) {
    throw new Error(
      `replaceImageBytes: replacement image content type '${detectedContentType}' does not match existing media content type '${media.contentType}'`,
    );
  }

  const sharedReferenceCount = countImageReferencesToMedia(source, media.partPath);
  const edit = {
    kind: "replaceImage",
    handle,
    mediaPartPath: media.partPath,
    contentType: media.contentType,
    sharedReferenceCount,
  } satisfies PptxSourceModelReplaceImageEdit;

  return {
    ...source,
    packageGraph: {
      ...source.packageGraph,
      media: source.packageGraph.media.map((part) =>
        part.partPath === media.partPath ? { ...part, bytes: copyBytes(bytes) } : part,
      ),
    },
    edits: [...(source.edits ?? []), edit],
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
  const newSlidePartPath = nextNumberedPartPath(
    source.packageGraph,
    source.edits?.flatMap((edit) => editReservedPartPaths(edit)) ?? [],
    "ppt/slides/slide",
    ".xml",
  );
  const newPresentationRelationshipId = nextRelationshipId(presentationRels.relationships);
  const newSlideNumericId = nextSlideNumericId(source, "duplicateSlide");
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

  let packageGraph = addPackagePart(source.packageGraph, {
    partPath: newSlidePartPath,
    contentType: slideContentType,
    bytes: copyBytes(sourceRawSlide.bytes),
    ...(newSlideRelationships === undefined ? {} : { relationships: newSlideRelationships }),
  });
  if (notesCopy !== undefined) {
    packageGraph = addPackagePart(packageGraph, {
      partPath: notesCopy.newPartPath,
      contentType: notesCopy.contentType,
      bytes: copyBytes(notesCopy.raw.bytes),
      ...(notesCopy.relationships === undefined ? {} : { relationships: notesCopy.relationships }),
    });
  }
  packageGraph = addPartRelationship(
    packageGraph,
    source.presentation.partPath,
    presentationSlideRelationship(source, newPresentationRelationshipId, newSlidePartPath),
  );

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
    packageGraph,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "duplicateSlide",
        sourceSlidePartPath: sourceSlide.partPath,
        sourceRelationshipId: sourcePresentationRelationship.id,
        newSlidePartPath,
        newRelationshipId: newPresentationRelationshipId,
        newSlideNumericId,
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
  const removedPartPaths = [slide.partPath, ...notesPartPaths];
  const removedPartPathSet = new Set<string>(removedPartPaths);
  const retainedEdits = (source.edits ?? []).filter(
    (edit) => !editIsInvalidatedByDeletedParts(edit, removedPartPathSet),
  );
  const deletedInsertedSlide = (source.edits ?? []).some(
    (edit) => editInsertedSlidePartPath(edit) === slide.partPath,
  );
  const packageGraph = removePartRelationship(
    removePackageParts(source.packageGraph, removedPartPaths),
    source.presentation.partPath,
    presentationRelationship.id,
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
    packageGraph,
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

function requireImageBySourceHandle(
  source: PptxSourceModel,
  handle: SourceHandle,
  operation: "replaceImageBytes",
): SourceImage {
  const shape = findShapeNodeBySourceHandle(source, handle);
  if (shape === undefined) {
    throw new Error(`${operation}: image handle was not found in PptxSourceModel source`);
  }
  if (shape.kind !== "image") {
    throw new Error(`${operation}: shape handle does not reference a pic image shape`);
  }
  return shape;
}

function requireMediaForImage(
  source: PptxSourceModel,
  image: SourceImage,
  operation: "replaceImageBytes",
) {
  if (image.blipRelationshipId === undefined) {
    throw new Error(`${operation}: image shape has no embedded blip relationship`);
  }
  const partPath = image.handle?.partPath;
  if (partPath === undefined) {
    throw new Error(`${operation}: image handle has no source part path`);
  }
  const relationships = requirePartRelationships(source, partPath, operation);
  const relationship = relationships.relationships.find(
    (candidate) => candidate.id === image.blipRelationshipId && candidate.type === IMAGE_REL_TYPE,
  );
  if (relationship === undefined) {
    throw new Error(`${operation}: image relationship was not found`);
  }
  const mediaPartPath = resolveInternalRelationshipTarget(
    relationships.sourcePartPath,
    relationship,
  );
  if (mediaPartPath === undefined) {
    throw new Error(`${operation}: image relationship does not target an internal media part`);
  }
  const media = source.packageGraph.media.find((part) => part.partPath === mediaPartPath);
  if (media === undefined) {
    throw new Error(`${operation}: image media part was not found`);
  }
  return media;
}

function countImageReferencesToMedia(source: PptxSourceModel, mediaPartPath: PartPath): number {
  const parsedImageRelationshipKeys = new Set<string>();
  let count = 0;

  const countParsedImages = (partPath: PartPath, shapes: readonly SourceShapeNode[]) => {
    for (const image of findImagesInTree(shapes)) {
      if (image.blipRelationshipId === undefined) continue;
      const relationships = source.packageGraph.relationships.find(
        (candidate) => candidate.sourcePartPath === partPath,
      );
      const relationship = relationships?.relationships.find(
        (candidate) =>
          candidate.id === image.blipRelationshipId && candidate.type === IMAGE_REL_TYPE,
      );
      if (relationship === undefined) continue;
      if (resolveInternalRelationshipTarget(partPath, relationship) === mediaPartPath) {
        count += 1;
        parsedImageRelationshipKeys.add(imageRelationshipKey(partPath, relationship.id));
      }
    }
  };

  for (const slide of source.slides) countParsedImages(slide.partPath, slide.shapes);
  for (const layout of source.slideLayouts) countParsedImages(layout.partPath, layout.shapes);
  for (const master of source.slideMasters) countParsedImages(master.partPath, master.shapes);

  for (const relationships of source.packageGraph.relationships) {
    for (const relationship of relationships.relationships) {
      if (relationship.type !== IMAGE_REL_TYPE) continue;
      if (
        parsedImageRelationshipKeys.has(
          imageRelationshipKey(relationships.sourcePartPath, relationship.id),
        )
      ) {
        continue;
      }
      if (
        resolveInternalRelationshipTarget(relationships.sourcePartPath, relationship) ===
        mediaPartPath
      ) {
        count += 1;
      }
    }
  }
  return count;
}

function imageRelationshipKey(partPath: PartPath, relationshipId: RelationshipId): string {
  return `${partPath}\0${relationshipId}`;
}

function findImagesInTree(shapes: readonly SourceShapeNode[]): SourceImage[] {
  return shapes.flatMap((shape): SourceImage[] => {
    if (shape.kind === "image") return [shape];
    if (shape.kind === "group") return findImagesInTree(shape.children);
    return [];
  });
}

function detectImageContentType(bytes: Uint8Array): string | undefined {
  if (startsWithBytes(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
    return "image/png";
  }
  if (startsWithBytes(bytes, [0xff, 0xd8, 0xff])) return "image/jpeg";
  if (
    startsWithBytes(bytes, [0x47, 0x49, 0x46, 0x38, 0x37, 0x61]) ||
    startsWithBytes(bytes, [0x47, 0x49, 0x46, 0x38, 0x39, 0x61])
  ) {
    return "image/gif";
  }
  if (startsWithBytes(bytes, [0x42, 0x4d])) return "image/bmp";
  if (
    startsWithBytes(bytes, [0x49, 0x49, 0x2a, 0x00]) ||
    startsWithBytes(bytes, [0x4d, 0x4d, 0x00, 0x2a])
  ) {
    return "image/tiff";
  }
  if (
    startsWithBytes(bytes, [0x52, 0x49, 0x46, 0x46]) &&
    bytes[8] === 0x57 &&
    bytes[9] === 0x45 &&
    bytes[10] === 0x42 &&
    bytes[11] === 0x50
  ) {
    return "image/webp";
  }
  return undefined;
}

function startsWithBytes(bytes: Uint8Array, prefix: readonly number[]): boolean {
  if (bytes.length < prefix.length) return false;
  return prefix.every((value, index) => bytes[index] === value);
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
  const usedNames = new Set([...used].map(String));
  return asSourceNodeId(nextNumberedName(usedNames, /^(\d+)$/, String));
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
    const numericId = Number(editReservedShapeId(edit, slidePartPath));
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
  const newPartPath = nextNumberedPartPath(
    source.packageGraph,
    source.edits?.flatMap((edit) => editReservedPartPaths(edit)) ?? [],
    "ppt/notesSlides/notesSlide",
    ".xml",
  );
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
  operationName: "addEmptySlideFromLayout" | "duplicateSlide" | "deleteSlide" | "replaceImageBytes",
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
  operationName: "addEmptySlideFromLayout" | "duplicateSlide",
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

/**
 * Assigns the numeric `p:sldId@id` for a new slide at edit time by reading the
 * preserved presentation XML and the ids already claimed by pending slide edits.
 * Ids freed by pending deletes are intentionally never reused.
 */
function nextSlideNumericId(
  source: PptxSourceModel,
  operationName: "addEmptySlideFromLayout" | "duplicateSlide",
): number {
  const rawPart = requireRawBinaryPart(source, source.presentation.partPath, operationName);
  const root = parseXml(textDecoder.decode(rawPart.bytes));
  const presentation = getChild(root, "presentation");
  if (presentation === undefined) {
    throw new Error(`${operationName}: presentation part does not contain p:presentation root`);
  }
  const used = new Set<number>();
  for (const item of getChildArray(getChild(presentation, "sldIdLst"), "sldId")) {
    const id = Number(getAttr(item, "id"));
    if (Number.isFinite(id)) used.add(id);
  }
  for (const edit of source.edits ?? []) {
    if (edit.kind === "addEmptySlideFromLayout" || edit.kind === "duplicateSlide") {
      used.add(edit.newSlideNumericId);
    }
  }
  const max = Math.max(255, ...used);
  for (let candidate = max + 1; ; candidate += 1) {
    if (!used.has(candidate)) return candidate;
  }
}

function presentationSlideRelationship(
  source: PptxSourceModel,
  relationshipId: RelationshipId,
  slidePartPath: PartPath,
): Relationship {
  return {
    id: relationshipId,
    type: SLIDE_REL_TYPE,
    target: relativeTarget(source.presentation.partPath, slidePartPath),
  };
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
  return edits.some((edit) => editDirtyPartPath(edit) === partPath);
}

function editIsInvalidatedByDeletedParts(
  edit: PptxSourceModelEdit,
  partPaths: ReadonlySet<string>,
): boolean {
  return editInvalidatingPartPaths(edit).some((partPath) => partPaths.has(partPath));
}
