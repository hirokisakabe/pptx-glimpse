import type { MediaPart, PartPath, PptxSourceModel } from "../source/index.js";
import { asPartPath } from "../source/index.js";
import { resolveRelationshipTarget } from "../source/package-paths.js";
import type { ComputedRelationship } from "./pptx-computed-view.js";

export function resolveComputedRelationships(
  source: PptxSourceModel,
  sourcePartPath: PartPath,
): ComputedRelationship[] {
  const rels =
    source.packageGraph.relationships.find((entry) => entry.sourcePartPath === sourcePartPath)
      ?.relationships ?? [];

  return rels.map((relationship) => {
    const external = relationship.targetMode === "External";
    const target = external
      ? relationship.target
      : resolveRelationshipTarget(sourcePartPath, relationship.target);
    const targetPartPath = external ? undefined : asPartPath(target);
    const media =
      targetPartPath !== undefined
        ? findMedia(source.packageGraph.media, targetPartPath)
        : undefined;

    return {
      id: relationship.id,
      type: relationship.type,
      source: relationship,
      target,
      ...(relationship.targetMode !== undefined ? { targetMode: relationship.targetMode } : {}),
      ...(targetPartPath !== undefined ? { targetPartPath } : {}),
      ...(media !== undefined ? { media } : {}),
    };
  });
}

function findMedia(media: readonly MediaPart[], partPath: PartPath): MediaPart | undefined {
  return media.find((part) => part.partPath === partPath);
}
