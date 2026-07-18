import { editReservedShapeId, sourceHandlesEqual } from "./edit-descriptors.js";
import { copyBytes, IMAGE_REL_TYPE, relativeTarget } from "./editing-shared.js";
import { assertShadowEffectsInput, type ShadowEffectsInput } from "./effect-authoring.js";
import { type PartPath } from "./handles.js";
import type {
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
  addMediaPartRelationship,
  nextNumberedName,
  nextNumberedPartPath,
  nextRelationshipId,
} from "./package-graph-mutations.js";
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
  readonly effects?: AddPictureEffectsInput;
  readonly name?: string;
}

export type AddPictureEffectsInput = ShadowEffectsInput;

interface DetectedImageType {
  readonly contentType: "image/png" | "image/jpeg";
  readonly extension: "png" | "jpeg";
}

interface PictureAuthoringTarget {
  readonly kind: "slide" | "layout" | "master";
  readonly index: number;
  readonly partPath: PartPath;
  readonly shapes: readonly SourceShapeNode[];
}

const IMAGE_MEDIA_PREFIX = "ppt/media/image";

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

  const target = findPictureAuthoringTarget(source, slideHandle);
  if (target === undefined) {
    throw new Error(
      "addPicture: slide, layout, or master handle was not found in PptxSourceModel source",
    );
  }

  const shapeId = nextPictureShapeId(source, target.partPath, target.shapes);
  const shapeIdValue = String(shapeId);
  const relationshipGroup = relationshipGroupForPart(source.packageGraph, target.partPath);
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
    target: relativeTarget(target.partPath, mediaPartPath),
  };
  const name = input.name?.trim() || `Picture ${shapeIdValue}`;
  const orderingSlot = nextOrderingSlot(target.shapes);
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
    ...(input.effects !== undefined ? { effects: input.effects } : {}),
  });
  const picture = parseShapeNodeXml(xml, target.partPath, orderingSlot);
  if (picture.kind !== "image") {
    throw new Error("addPicture: finalized picture XML did not parse as a p:pic image");
  }

  const edit = {
    kind: "addPicture",
    slidePartPath: target.partPath,
    shapeId: shapeIdValue,
    relationshipId,
    mediaPartPath,
    contentType: imageType.contentType,
    xml,
  } satisfies PptxSourceModelAddPictureEdit;

  return {
    ...withPictureAuthoringTargetShapes(source, target, [...target.shapes, picture]),
    packageGraph: addMediaPartRelationship(source.packageGraph, {
      ownerPartPath: target.partPath,
      media,
      extension: imageType.extension,
      relationship,
      contentTypeDefaultConflictError: (existingContentType) =>
        new Error(
          `addPicture: content type default for extension '${imageType.extension}' already maps to '${existingContentType}'`,
        ),
    }),
    edits: [...(source.edits ?? []), edit],
  };
}

function relationshipGroupForPart(graph: PackageGraph, slidePartPath: PartPath): PartRelationships {
  return (
    graph.relationships.find((candidate) => candidate.sourcePartPath === slidePartPath) ?? {
      sourcePartPath: slidePartPath,
      relationships: [],
    }
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

function nextPictureShapeId(
  source: PptxSourceModel,
  slidePartPath: PartPath,
  shapes: readonly SourceShapeNode[],
): string {
  const used = new Set<number>();
  collectShapeIds(shapes, used);
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
  if (input.effects !== undefined) {
    assertShadowEffectsInput(input.effects, "addPicture");
    if (input.effects.outerShadow === undefined && input.effects.innerShadow === undefined) {
      throw new Error("addPicture: effects must set outerShadow or innerShadow");
    }
  }
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

function findPictureAuthoringTarget(
  source: PptxSourceModel,
  handle: SourceHandle,
): PictureAuthoringTarget | undefined {
  const groups = [
    { kind: "slide" as const, values: source.slides },
    { kind: "layout" as const, values: source.slideLayouts },
    { kind: "master" as const, values: source.slideMasters },
  ];
  for (const group of groups) {
    const index = group.values.findIndex((candidate) =>
      sourceHandlesEqual(candidate.handle, handle),
    );
    if (index >= 0) {
      const candidate = group.values[index];
      return {
        kind: group.kind,
        index,
        partPath: candidate.partPath,
        shapes: candidate.shapes,
      };
    }
  }
  return undefined;
}

function withPictureAuthoringTargetShapes(
  source: PptxSourceModel,
  target: PictureAuthoringTarget,
  shapes: readonly SourceShapeNode[],
): PptxSourceModel {
  switch (target.kind) {
    case "slide":
      return {
        ...source,
        slides: source.slides.map((candidate, index) =>
          index === target.index ? { ...candidate, shapes } : candidate,
        ),
      };
    case "layout":
      return {
        ...source,
        slideLayouts: source.slideLayouts.map((candidate, index) =>
          index === target.index ? { ...candidate, shapes } : candidate,
        ),
      };
    case "master":
      return {
        ...source,
        slideMasters: source.slideMasters.map((candidate, index) =>
          index === target.index ? { ...candidate, shapes } : candidate,
        ),
      };
  }
}
