import { parseXml } from "../reader/xml.js";
import { sourceHandlesEqual } from "./edit-descriptors.js";
import {
  assertNeverShapeNode,
  copyBytes,
  IMAGE_REL_TYPE,
  relativeTarget,
  requirePartRelationships,
} from "./editing-shared.js";
import { detectSupportedImageType, startsWithBytes } from "./image-type.js";
import type {
  MediaPart,
  PartPath,
  PptxSourceModel,
  PptxSourceModelReplaceImageEdit,
  Relationship,
  RelationshipId,
  SourceHandle,
  SourceImage,
  SourceShapeNode,
} from "./index.js";
import {
  addMediaPartRelationship,
  nextNumberedPartPath,
  nextRelationshipId,
} from "./package-graph-mutations.js";
import { resolveInternalRelationshipTarget } from "./package-paths.js";
import { findShapeNodeBySourceHandle } from "./shape-editing.js";

const IMAGE_MEDIA_PREFIX = "ppt/media/image";

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
  if (sharedReferenceCount > 1) {
    return replaceSharedImageBytes(
      source,
      handle,
      media,
      bytes,
      detectedContentType,
      sharedReferenceCount,
    );
  }
  const edit = {
    kind: "replaceImage",
    handle,
    mode: "inPlace",
    sourceMediaPartPath: media.partPath,
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

function replaceSharedImageBytes(
  source: PptxSourceModel,
  handle: SourceHandle,
  media: MediaPart,
  bytes: Uint8Array,
  contentType: string,
  sharedReferenceCount: number,
): PptxSourceModel {
  if (handle.nodeId === undefined) {
    throw new Error("replaceImageBytes: shared image replacement requires a picture node id");
  }
  const ownerPartPath = handle.partPath;
  const relationships = requirePartRelationships(source, ownerPartPath, "replaceImageBytes");
  const replacementRelationshipId = nextRelationshipId(relationships.relationships);
  const extension = imageExtension(contentType);
  const reservedMediaPartPaths = (source.edits ?? []).flatMap((edit) =>
    edit.kind === "replaceImage" && edit.mode === "copyOnWrite" ? [edit.mediaPartPath] : [],
  );
  const mediaPartPath = nextNumberedPartPath(
    source.packageGraph,
    reservedMediaPartPaths,
    IMAGE_MEDIA_PREFIX,
    `.${extension}`,
  );
  const replacementMedia: MediaPart = {
    partPath: mediaPartPath,
    contentType,
    bytes: copyBytes(bytes),
  };
  const relationship: Relationship = {
    id: replacementRelationshipId,
    type: IMAGE_REL_TYPE,
    target: relativeTarget(ownerPartPath, mediaPartPath),
  };
  const edit = {
    kind: "replaceImage",
    handle,
    mode: "copyOnWrite",
    sourceMediaPartPath: media.partPath,
    mediaPartPath,
    contentType,
    sharedReferenceCount,
    replacementRelationshipId,
  } satisfies PptxSourceModelReplaceImageEdit;

  return {
    ...source,
    ...replaceImageRelationship(source, handle, replacementRelationshipId),
    packageGraph: addMediaPartRelationship(source.packageGraph, {
      ownerPartPath,
      media: replacementMedia,
      extension,
      relationship,
      useOverrideOnContentTypeDefaultConflict: true,
      contentTypeDefaultConflictError: (existingContentType) =>
        new Error(
          `replaceImageBytes: content type default for extension '${extension}' already maps to '${existingContentType}'`,
        ),
    }),
    edits: [...(source.edits ?? []), edit],
  };
}

function replaceImageRelationship(
  source: PptxSourceModel,
  handle: SourceHandle,
  relationshipId: RelationshipId,
): Pick<PptxSourceModel, "slides" | "slideLayouts" | "slideMasters"> {
  const replaceShapes = (shapes: readonly SourceShapeNode[]): readonly SourceShapeNode[] => {
    let changed = false;
    const next = shapes.map((shape): SourceShapeNode => {
      if (sourceHandlesEqual(shape.handle, handle)) {
        if (shape.kind !== "image") return shape;
        changed = true;
        return {
          ...shape,
          blipRelationshipId: relationshipId,
          ...(shape.handle === undefined ? {} : { handle: { ...shape.handle, relationshipId } }),
        };
      }
      if (shape.kind !== "group") return shape;
      const children = replaceShapes(shape.children);
      if (children === shape.children) return shape;
      changed = true;
      return { ...shape, children };
    });
    return changed ? next : shapes;
  };
  const replaceTargets = <
    T extends { readonly partPath: PartPath; readonly shapes: readonly SourceShapeNode[] },
  >(
    targets: readonly T[],
  ): readonly T[] =>
    targets.map((target) => {
      if (target.partPath !== handle.partPath) return target;
      const shapes = replaceShapes(target.shapes);
      return shapes === target.shapes ? target : { ...target, shapes };
    });

  return {
    slides: replaceTargets(source.slides),
    slideLayouts: replaceTargets(source.slideLayouts),
    slideMasters: replaceTargets(source.slideMasters),
  };
}

function imageExtension(contentType: string): string {
  switch (contentType) {
    case "image/png":
      return "png";
    case "image/jpeg":
      return "jpeg";
    case "image/gif":
      return "gif";
    case "image/bmp":
      return "bmp";
    case "image/tiff":
      return "tiff";
    case "image/webp":
      return "webp";
    default:
      throw new Error(`replaceImageBytes: unsupported image content type '${contentType}'`);
  }
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

/**
 * Counts typed picture references and preserved OOXML uses of an internal image media part.
 * Unknown or unavailable owner XML is treated conservatively so callers do not mutate a
 * potentially shared part in place.
 */
export function countImageReferencesToMedia(
  source: PptxSourceModel,
  mediaPartPath: PartPath,
): number {
  const parsedImageCounts = new Map<string, number>();

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
        const key = imageRelationshipKey(partPath, relationship.id);
        parsedImageCounts.set(key, (parsedImageCounts.get(key) ?? 0) + 1);
      }
    }
  };

  for (const slide of source.slides) countParsedImages(slide.partPath, slide.shapes);
  for (const layout of source.slideLayouts) countParsedImages(layout.partPath, layout.shapes);
  for (const master of source.slideMasters) countParsedImages(master.partPath, master.shapes);

  let count = 0;
  for (const relationships of source.packageGraph.relationships) {
    for (const relationship of relationships.relationships) {
      if (relationship.type !== IMAGE_REL_TYPE) continue;
      if (
        resolveInternalRelationshipTarget(relationships.sourcePartPath, relationship) !==
        mediaPartPath
      )
        continue;
      const parsedCount =
        parsedImageCounts.get(
          imageRelationshipKey(relationships.sourcePartPath, relationship.id),
        ) ?? 0;
      const preservedXmlCount = countPreservedRelationshipUses(
        source,
        relationships.sourcePartPath,
        relationship.id,
      );
      const pendingReplacementCount = countPendingRelationshipReplacements(
        source,
        relationships.sourcePartPath,
        relationship.id,
        mediaPartPath,
      );
      // If preserved XML is unavailable, a parsed target cannot prove exclusive use.
      const conservativeCount =
        preservedXmlCount === undefined
          ? parsedCount > 0
            ? 2
            : 1
          : Math.max(0, preservedXmlCount - pendingReplacementCount);
      count += Math.max(1, parsedCount, conservativeCount);
    }
  }
  return count;
}

function countPendingRelationshipReplacements(
  source: PptxSourceModel,
  ownerPartPath: PartPath,
  relationshipId: RelationshipId,
  mediaPartPath: PartPath,
): number {
  return (source.edits ?? []).filter(
    (edit) =>
      edit.kind === "replaceImage" &&
      edit.mode === "copyOnWrite" &&
      edit.handle.partPath === ownerPartPath &&
      edit.handle.relationshipId === relationshipId &&
      edit.sourceMediaPartPath === mediaPartPath,
  ).length;
}

const textDecoder = new TextDecoder();

function countPreservedRelationshipUses(
  source: PptxSourceModel,
  partPath: PartPath,
  relationshipId: RelationshipId,
): number | undefined {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === partPath);
  if (rawPart?.kind !== "binary") return undefined;
  try {
    return countRelationshipAttributes(parseXml(textDecoder.decode(rawPart.bytes)), relationshipId);
  } catch {
    return undefined;
  }
}

function countRelationshipAttributes(value: unknown, relationshipId: RelationshipId): number {
  if (isUnknownArray(value)) {
    let count = 0;
    for (const item of value) count += countRelationshipAttributes(item, relationshipId);
    return count;
  }
  if (!isUnknownRecord(value)) return 0;
  return Object.entries(value).reduce((count, [key, item]) => {
    const isRelationshipAttribute =
      key.startsWith("@_") &&
      (key.endsWith(":embed") || key.endsWith(":link") || key.endsWith(":id"));
    return (
      count +
      (isRelationshipAttribute && item === relationshipId ? 1 : 0) +
      countRelationshipAttributes(item, relationshipId)
    );
  }, 0);
}

function isUnknownArray(value: unknown): value is unknown[] {
  return Array.isArray(value);
}

function isUnknownRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
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
  const supportedImageType = detectSupportedImageType(bytes);
  if (supportedImageType !== undefined) return supportedImageType.contentType;
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
