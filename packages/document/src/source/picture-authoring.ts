import { editReservedShapeId, sourceHandlesEqual } from "./edit-descriptors.js";
import { copyBytes, IMAGE_REL_TYPE, relativeTarget } from "./editing-shared.js";
import { asPartPath, type PartPath } from "./handles.js";
import type {
  ContentTypeDefault,
  MediaPart,
  PackageGraph,
  PartRelationships,
  PptxSourceModel,
  PptxSourceModelAddPictureEdit,
  Relationship,
  SourceHandle,
  SourceShapeNode,
} from "./index.js";
import {
  nextNumberedName,
  nextNumberedPartPath,
  nextRelationshipId,
} from "./package-graph-mutations.js";
import { relationshipsPartPath } from "./package-paths.js";
import { buildPictureXml, parseShapeNodeXml } from "./shape-xml.js";
import type { SourceImageCrop, SourceTransform } from "./shapes.js";
import type { Emu, OoxmlAngle, OoxmlPercent } from "./units.js";

export interface AddPictureCropInput {
  readonly left?: OoxmlPercent;
  readonly top?: OoxmlPercent;
  readonly right?: OoxmlPercent;
  readonly bottom?: OoxmlPercent;
}

export interface AddPictureInput {
  readonly bytes: Uint8Array;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly rotation?: OoxmlAngle;
  readonly crop?: AddPictureCropInput;
  readonly name?: string;
}

interface DetectedImageType {
  readonly contentType: "image/png" | "image/jpeg";
  readonly extension: "png" | "jpeg";
}

const IMAGE_MEDIA_PREFIX = "ppt/media/image";
const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";

export function addPicture(
  source: PptxSourceModel,
  slideHandle: SourceHandle,
  input: AddPictureInput,
): PptxSourceModel {
  assertPictureInput(input);
  const imageType = detectSupportedImageType(input.bytes);
  if (imageType === undefined) {
    throw new Error("addPicture: unsupported or unknown image format");
  }

  const slideIndex = source.slides.findIndex((slide) =>
    sourceHandlesEqual(slide.handle, slideHandle),
  );
  if (slideIndex === -1) {
    throw new Error("addPicture: slide handle was not found in PptxSourceModel source");
  }

  const slide = source.slides[slideIndex];
  const shapeId = nextPictureShapeId(source, slide.partPath);
  const shapeIdValue = String(shapeId);
  const relationshipGroup = relationshipGroupForSlide(source.packageGraph, slide.partPath);
  const relationshipId = nextRelationshipId(relationshipGroup.relationships);
  const mediaPartPath = nextMediaPartPath(source.packageGraph, source.edits ?? [], imageType);
  const media: MediaPart = {
    partPath: mediaPartPath,
    contentType: imageType.contentType,
    bytes: copyBytes(input.bytes),
  };
  const relationship: Relationship = {
    id: relationshipId,
    type: IMAGE_REL_TYPE,
    target: relativeTarget(slide.partPath, mediaPartPath),
  };
  const name = input.name?.trim() || `Picture ${shapeIdValue}`;
  const orderingSlot = nextOrderingSlot(slide.shapes);
  const xml = buildPictureXml({
    shapeId: shapeIdValue,
    name,
    relationshipId,
    offsetX: input.offsetX,
    offsetY: input.offsetY,
    width: input.width,
    height: input.height,
    ...(input.rotation !== undefined ? { rotation: input.rotation } : {}),
    ...(input.crop !== undefined ? { crop: input.crop } : {}),
  });
  const picture = parseShapeNodeXml(xml, slide.partPath, orderingSlot);
  if (picture.kind !== "image") {
    throw new Error("addPicture: finalized picture XML did not parse as a p:pic image");
  }

  const edit = {
    kind: "addPicture",
    slidePartPath: slide.partPath,
    shapeId: shapeIdValue,
    relationshipId,
    mediaPartPath,
    contentType: imageType.contentType,
    xml,
  } satisfies PptxSourceModelAddPictureEdit;

  return {
    ...source,
    packageGraph: addPicturePackageGraphEntries(
      source.packageGraph,
      slide.partPath,
      media,
      imageType.extension,
      relationship,
    ),
    slides: source.slides.map((candidate, index) =>
      index === slideIndex
        ? {
            ...candidate,
            shapes: [...candidate.shapes, picture],
          }
        : candidate,
    ),
    edits: [...(source.edits ?? []), edit],
  };
}

function addPicturePackageGraphEntries(
  graph: PackageGraph,
  slidePartPath: PartPath,
  media: MediaPart,
  extension: string,
  relationship: Relationship,
): PackageGraph {
  const relationshipGroup = relationshipGroupForSlide(graph, slidePartPath);
  const relationshipPartPath = asPartPath(relationshipsPartPath(slidePartPath));
  const hasRelationshipGroup = graph.relationships.some(
    (candidate) => candidate.sourcePartPath === slidePartPath,
  );
  const hasRelationshipPart = graph.parts.some((part) => part.partPath === relationshipPartPath);
  const needsRelationshipOverride =
    !hasRelationshipPart &&
    !graph.contentTypes.defaults.some(
      (entry) => entry.extension === "rels" && entry.contentType === RELS_CONTENT_TYPE,
    ) &&
    !graph.contentTypes.overrides.some((entry) => entry.partName === relationshipPartPath);

  return {
    ...graph,
    contentTypes: {
      ...graph.contentTypes,
      defaults: withImageContentTypeDefault(
        graph.contentTypes.defaults,
        extension,
        media.contentType,
      ),
      overrides: [
        ...graph.contentTypes.overrides,
        ...(needsRelationshipOverride
          ? [{ partName: relationshipPartPath, contentType: RELS_CONTENT_TYPE }]
          : []),
      ],
    },
    parts: [
      ...graph.parts,
      { partPath: media.partPath, contentType: media.contentType },
      ...(hasRelationshipPart
        ? []
        : [
            {
              partPath: relationshipPartPath,
              contentType: RELS_CONTENT_TYPE,
            },
          ]),
    ],
    relationships: hasRelationshipGroup
      ? graph.relationships.map((candidate) =>
          candidate.sourcePartPath === slidePartPath
            ? { ...candidate, relationships: [...candidate.relationships, relationship] }
            : candidate,
        )
      : [...graph.relationships, { ...relationshipGroup, relationships: [relationship] }],
    media: [...graph.media, media],
  };
}

function relationshipGroupForSlide(
  graph: PackageGraph,
  slidePartPath: PartPath,
): PartRelationships {
  return (
    graph.relationships.find((candidate) => candidate.sourcePartPath === slidePartPath) ?? {
      sourcePartPath: slidePartPath,
      relationships: [],
    }
  );
}

function withImageContentTypeDefault(
  defaults: readonly ContentTypeDefault[],
  extension: string,
  contentType: string,
): readonly ContentTypeDefault[] {
  const existing = defaults.find((entry) => entry.extension === extension);
  if (existing === undefined) return [...defaults, { extension, contentType }];
  if (existing.contentType === contentType) return defaults;
  throw new Error(
    `addPicture: content type default for extension '${extension}' already maps to '${existing.contentType}'`,
  );
}

function nextMediaPartPath(
  graph: PackageGraph,
  edits: readonly { readonly kind: string; readonly mediaPartPath?: PartPath }[],
  imageType: DetectedImageType,
): PartPath {
  const reserved = edits.flatMap((edit) => {
    if (edit.kind !== "addPicture" || edit.mediaPartPath === undefined) return [];
    return [edit.mediaPartPath];
  });
  return nextNumberedPartPath(graph, reserved, IMAGE_MEDIA_PREFIX, `.${imageType.extension}`);
}

function nextPictureShapeId(source: PptxSourceModel, slidePartPath: PartPath): string {
  const slide = source.slides.find((candidate) => candidate.partPath === slidePartPath);
  const used = new Set<number>();
  if (slide !== undefined) collectShapeIds(slide.shapes, used);
  for (const edit of source.edits ?? []) {
    const numericId = Number(editReservedShapeId(edit, slidePartPath));
    if (Number.isInteger(numericId) && numericId > 0) used.add(numericId);
  }
  const usedNames = new Set([...used].map(String));
  return nextNumberedName(usedNames, /^(\d+)$/, String);
}

function collectShapeIds(shapes: readonly SourceShapeNode[], used: Set<number>): void {
  for (const shape of shapes) {
    const numericId = Number(shape.nodeId);
    if (Number.isInteger(numericId) && numericId > 0) used.add(numericId);
    if (shape.kind === "group") collectShapeIds(shape.children, used);
  }
}

function nextOrderingSlot(shapes: readonly { readonly handle?: SourceHandle }[]): number {
  return (
    shapes.reduce((current, shape) => {
      const slot = shape.handle?.orderingSlot ?? -1;
      return Math.max(current, slot);
    }, -1) + 1
  );
}

function assertPictureInput(input: AddPictureInput): void {
  assertFiniteEmu(input.offsetX, "offsetX");
  assertFiniteEmu(input.offsetY, "offsetY");
  assertPositiveFiniteEmu(input.width, "width");
  assertPositiveFiniteEmu(input.height, "height");
  if (!(input.bytes instanceof Uint8Array)) {
    throw new Error("addPicture: bytes must be a Uint8Array");
  }
  if (input.rotation !== undefined && !Number.isInteger(input.rotation)) {
    throw new Error("addPicture: rotation must be a finite integer");
  }
  if (input.name !== undefined && input.name.trim() === "") {
    throw new Error("addPicture: name must be a non-empty string when provided");
  }
  if (input.crop !== undefined) assertCrop(input.crop);
}

function assertCrop(crop: AddPictureCropInput): asserts crop is SourceImageCrop {
  assertCropInset(crop.left, "crop.left");
  assertCropInset(crop.top, "crop.top");
  assertCropInset(crop.right, "crop.right");
  assertCropInset(crop.bottom, "crop.bottom");
}

function assertCropInset(value: OoxmlPercent | undefined, fieldName: string): void {
  if (value === undefined) return;
  if (typeof value !== "number" || !Number.isInteger(value)) {
    throw new Error(`addPicture: ${fieldName} must be an integer OOXML percentage`);
  }
}

function assertFiniteEmu(value: unknown, fieldName: keyof SourceTransform): void {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`addPicture: ${fieldName} must be a finite EMU value`);
  }
}

function assertPositiveFiniteEmu(value: unknown, fieldName: keyof SourceTransform): void {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw new Error(`addPicture: ${fieldName} must be a finite positive EMU value`);
  }
}

function detectSupportedImageType(bytes: Uint8Array): DetectedImageType | undefined {
  if (startsWithBytes(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
    return { contentType: "image/png", extension: "png" };
  }
  if (startsWithBytes(bytes, [0xff, 0xd8, 0xff])) {
    return { contentType: "image/jpeg", extension: "jpeg" };
  }
  return undefined;
}

function startsWithBytes(bytes: Uint8Array, prefix: readonly number[]): boolean {
  if (bytes.length < prefix.length) return false;
  return prefix.every((value, index) => bytes[index] === value);
}
