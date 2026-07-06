import {
  assertNeverShapeNode,
  copyBytes,
  IMAGE_REL_TYPE,
  requirePartRelationships,
} from "./editing-shared.js";
import type {
  PartPath,
  PptxSourceModel,
  PptxSourceModelReplaceImageEdit,
  RelationshipId,
  SourceHandle,
  SourceImage,
  SourceShapeNode,
} from "./index.js";
import { resolveInternalRelationshipTarget } from "./package-paths.js";
import { findShapeNodeBySourceHandle } from "./shape-editing.js";

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
    switch (shape.kind) {
      case "image":
        return [shape];
      case "group":
        return findImagesInTree(shape.children);
      case "shape":
      case "connector":
      case "table":
      case "chart":
      case "smartArt":
      case "raw":
        // The denominator is for replaceImageBytes' p:pic targets only. Other typed
        // nodes, image fills, and raw/unsupported relationship users are preserved but
        // are not replaceImageBytes targets in this editing slice.
        return [];
    }
    return assertNeverShapeNode(shape);
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
