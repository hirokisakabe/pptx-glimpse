/**
 * Pure mutation helpers for PackageGraph.
 *
 * Adding or removing a single package part must update its related PackageGraph
 * lists consistently. These helpers own that invariant so each editing operation
 * does not have to re-implement it by hand. All functions return a new PackageGraph
 * and never mutate their input. Media registration deliberately uses `graph.media`
 * instead of `rawParts`, matching the writer's source-of-truth invariant.
 */

import { asPartPath, asRelationshipId, type PartPath, type RelationshipId } from "./handles.js";
import type {
  ContentTypeOverride,
  MediaPart,
  PackageGraph,
  PartRelationships,
  Relationship,
} from "./package-graph.js";
import { relationshipsPartPath } from "./package-paths.js";

const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";

interface AddPackagePartInput {
  readonly partPath: PartPath;
  readonly contentType: string;
  /** Raw binary payload preserved for writing the new part back. */
  readonly bytes: Uint8Array;
  /** Relationships owned by the new part. When present, the `.rels` part is registered too. */
  readonly relationships?: PartRelationships;
}

interface AddMediaPartRelationshipInput {
  readonly ownerPartPath: PartPath;
  readonly media: MediaPart;
  readonly extension: string;
  readonly relationship: Relationship;
  readonly contentTypeDefaultConflictError: (existingContentType: string) => Error;
}

/** Adds one part (and optionally its `.rels` part) to all four PackageGraph lists. */
export function addPackagePart(graph: PackageGraph, input: AddPackagePartInput): PackageGraph {
  const relationshipsPart =
    input.relationships === undefined
      ? undefined
      : asPartPath(relationshipsPartPath(input.partPath));
  return {
    ...graph,
    parts: [
      ...graph.parts,
      { partPath: input.partPath, contentType: input.contentType },
      ...(relationshipsPart === undefined
        ? []
        : [{ partPath: relationshipsPart, contentType: RELS_CONTENT_TYPE }]),
    ],
    contentTypes: {
      ...graph.contentTypes,
      overrides: [
        ...graph.contentTypes.overrides,
        { partName: input.partPath, contentType: input.contentType },
        ...relationshipPartOverrides(
          graph,
          relationshipsPart === undefined ? [] : [relationshipsPart],
        ),
      ],
    },
    relationships: [
      ...graph.relationships,
      ...(input.relationships === undefined ? [] : [input.relationships]),
    ],
    rawParts: [
      ...(graph.rawParts ?? []),
      {
        kind: "binary",
        partPath: input.partPath,
        contentType: input.contentType,
        bytes: input.bytes,
      },
    ],
  };
}

/**
 * Registers a media part and its relationship from an existing owner part.
 *
 * The media payload is added only to `media`, never to `rawParts`. The owner's
 * relationship group and `.rels` package part are created when absent, including
 * a content type override when the package has no standard `rels` default.
 */
export function addMediaPartRelationship(
  graph: PackageGraph,
  input: AddMediaPartRelationshipInput,
): PackageGraph {
  const relationshipPart = asPartPath(relationshipsPartPath(input.ownerPartPath));
  const hasRelationshipGroup = graph.relationships.some(
    (candidate) => candidate.sourcePartPath === input.ownerPartPath,
  );
  const hasRelationshipPart = graph.parts.some((part) => part.partPath === relationshipPart);

  return {
    ...graph,
    parts: [
      ...graph.parts,
      { partPath: input.media.partPath, contentType: input.media.contentType },
      ...(hasRelationshipPart
        ? []
        : [{ partPath: relationshipPart, contentType: RELS_CONTENT_TYPE }]),
    ],
    contentTypes: {
      ...graph.contentTypes,
      defaults: withContentTypeDefault(
        graph,
        input.extension,
        input.media.contentType,
        input.contentTypeDefaultConflictError,
      ),
      overrides: [
        ...graph.contentTypes.overrides,
        ...relationshipPartOverrides(graph, hasRelationshipPart ? [] : [relationshipPart]),
      ],
    },
    relationships: hasRelationshipGroup
      ? graph.relationships.map((candidate) =>
          candidate.sourcePartPath === input.ownerPartPath
            ? {
                ...candidate,
                relationships: [...candidate.relationships, input.relationship],
              }
            : candidate,
        )
      : [
          ...graph.relationships,
          { sourcePartPath: input.ownerPartPath, relationships: [input.relationship] },
        ],
    media: [...graph.media, input.media],
  };
}

/** Removes the given parts and their `.rels` parts from all four PackageGraph lists. */
export function removePackageParts(
  graph: PackageGraph,
  partPaths: readonly PartPath[],
): PackageGraph {
  const removedPartPaths = new Set<string>(partPaths);
  const removedRelationshipPartPaths = new Set<string>(
    partPaths.map((partPath) => relationshipsPartPath(partPath)),
  );
  return {
    ...graph,
    parts: graph.parts.filter(
      (part) =>
        !removedPartPaths.has(part.partPath) && !removedRelationshipPartPaths.has(part.partPath),
    ),
    contentTypes: {
      ...graph.contentTypes,
      overrides: graph.contentTypes.overrides.filter(
        (override) =>
          !removedPartPaths.has(override.partName) &&
          !removedRelationshipPartPaths.has(override.partName),
      ),
    },
    relationships: graph.relationships.filter(
      (relationships) => !removedPartPaths.has(relationships.sourcePartPath),
    ),
    rawParts: graph.rawParts?.filter(
      (part) =>
        !removedPartPaths.has(part.partPath) && !removedRelationshipPartPaths.has(part.partPath),
    ),
  };
}

/**
 * Appends one relationship to the `.rels` owned by sourcePartPath. No-op when
 * the graph has no relationships entry for that part; callers must ensure the
 * entry exists beforehand.
 */
export function addPartRelationship(
  graph: PackageGraph,
  sourcePartPath: PartPath,
  relationship: Relationship,
): PackageGraph {
  return {
    ...graph,
    relationships: graph.relationships.map((relationships) =>
      relationships.sourcePartPath !== sourcePartPath
        ? relationships
        : {
            ...relationships,
            relationships: [...relationships.relationships, relationship],
          },
    ),
  };
}

/**
 * Removes one relationship by id from the `.rels` owned by sourcePartPath.
 * No-op when the graph has no relationships entry for that part.
 */
export function removePartRelationship(
  graph: PackageGraph,
  sourcePartPath: PartPath,
  relationshipId: RelationshipId,
): PackageGraph {
  return {
    ...graph,
    relationships: graph.relationships.map((relationships) =>
      relationships.sourcePartPath !== sourcePartPath
        ? relationships
        : {
            ...relationships,
            relationships: relationships.relationships.filter(
              (relationship) => relationship.id !== relationshipId,
            ),
          },
    ),
  };
}

/**
 * Shared "max trailing number + 1" allocator: scan used names with the pattern
 * (capture group 1 must be the number; the pattern must not carry the g / y
 * flags, which would make exec stateful), continue after the maximum, and skip
 * candidates that are still taken.
 */
export function nextNumberedName(
  used: ReadonlySet<string>,
  pattern: RegExp,
  format: (index: number) => string,
): string {
  const max = [...used].reduce((current, value) => {
    const match = pattern.exec(value);
    return match === null ? current : Math.max(current, Number(match[1]));
  }, 0);
  for (let index = max + 1; ; index += 1) {
    const candidate = format(index);
    if (!used.has(candidate)) return candidate;
  }
}

export function nextRelationshipId(relationships: readonly Relationship[]): RelationshipId {
  const used = new Set<string>(relationships.map((relationship) => relationship.id));
  return asRelationshipId(nextNumberedName(used, /(\d+)$/, (index) => `rId${index}`));
}

/**
 * Allocates the next `<prefix><number><suffix>` part path that is unused by the
 * package graph. reservedPartPaths lets callers exclude paths that are no longer
 * in the graph but must not be reused (e.g. parts referenced by the edit journal).
 */
export function nextNumberedPartPath(
  graph: PackageGraph,
  reservedPartPaths: readonly string[],
  prefix: string,
  suffix: string,
): PartPath {
  const used = new Set<string>([
    ...graph.parts.map((part) => part.partPath),
    ...graph.contentTypes.overrides.map((override) => override.partName),
    ...(graph.rawParts ?? []).map((part) => part.partPath),
    ...reservedPartPaths,
  ]);
  const pattern = new RegExp(`^${escapeRegExp(prefix)}(\\d+)${escapeRegExp(suffix)}$`);
  return asPartPath(nextNumberedName(used, pattern, (index) => `${prefix}${index}${suffix}`));
}

/**
 * Content type overrides required for newly added `.rels` parts. Packages that
 * declare the standard `rels` extension default need no override; without the
 * default, each new `.rels` part needs an explicit override entry.
 */
function relationshipPartOverrides(
  graph: PackageGraph,
  partPaths: readonly PartPath[],
): readonly ContentTypeOverride[] {
  if (
    graph.contentTypes.defaults.some(
      (entry) => entry.extension === "rels" && entry.contentType === RELS_CONTENT_TYPE,
    )
  ) {
    return [];
  }

  const existingOverrides = new Set(
    graph.contentTypes.overrides.map((override) => override.partName),
  );
  return partPaths
    .filter((partPath) => !existingOverrides.has(partPath))
    .map((partName) => ({ partName, contentType: RELS_CONTENT_TYPE }));
}

function withContentTypeDefault(
  graph: PackageGraph,
  extension: string,
  contentType: string,
  conflictError: (existingContentType: string) => Error,
): PackageGraph["contentTypes"]["defaults"] {
  const existing = graph.contentTypes.defaults.find((entry) => entry.extension === extension);
  if (existing === undefined) {
    return [...graph.contentTypes.defaults, { extension, contentType }];
  }
  if (existing.contentType === contentType) return graph.contentTypes.defaults;
  throw conflictError(existing.contentType);
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
