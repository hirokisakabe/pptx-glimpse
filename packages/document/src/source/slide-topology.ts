import {
  editInsertedSlidePartPath,
  editReservedPartPaths,
  sourceHandlesEqual,
} from "./edit-descriptors.js";
import {
  cloneJson,
  copyBytes,
  editIsInvalidatedByDeletedParts,
  EMPTY_SLIDE_XML,
  hasDirtyEditForPart,
  insertAtReadonly,
  nextSlideNumericId,
  NOTES_SLIDE_CONTENT_TYPE,
  NOTES_SLIDE_REL_TYPE,
  presentationSlideRelationship,
  relativeTarget,
  requirePartRelationships,
  requireRawBinaryPart,
  requireSlideRelationship,
  SLIDE_CONTENT_TYPE,
  SLIDE_LAYOUT_REL_TYPE,
  SLIDE_REL_TYPE,
} from "./editing-shared.js";
import type {
  PartPath,
  PartRelationships,
  PptxSourceModel,
  RawPackagePart,
  RelationshipId,
  SourceHandle,
  SourceShapeNode,
  SourceSlide,
} from "./index.js";
import { asRelationshipId } from "./index.js";
import {
  addPackagePart,
  addPartRelationship,
  nextNumberedPartPath,
  nextRelationshipId,
  removePackageParts,
  removePartRelationship,
} from "./package-graph-mutations.js";
import { resolveInternalRelationshipTarget } from "./package-paths.js";

export interface AddEmptySlideFromLayoutInput {
  readonly layoutPartPath: PartPath;
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
  // This topology operation removes the slide and its directly attached notes slide
  // only. It intentionally does not sweep now-orphaned media parts because raw or
  // unsupported parts may still reference them; media reference counting remains local
  // to replaceImageBytes.
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
