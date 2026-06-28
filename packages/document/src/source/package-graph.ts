/**
 * Internal note.
 * Internal note.
 */

import type { PartPath, RelationshipId } from "./handles.js";
import type { RawPackagePart } from "./raw.js";

/** Internal note. */
export type RelationshipTargetMode = "Internal" | "External";

/** Internal note. */
export interface Relationship {
  readonly id: RelationshipId;
  /** relationship type URI。 */
  readonly type: string;
  /** Target relative to sourcePartPath, or an external URL. */
  readonly target: string;
  readonly targetMode?: RelationshipTargetMode;
}

/** Relationships owned by a source part (`<part>/_rels/<part>.rels`)。 */
export interface PartRelationships {
  readonly sourcePartPath: PartPath;
  readonly relationships: readonly Relationship[];
}

/** Internal note. */
export interface ContentTypeDefault {
  readonly extension: string;
  readonly contentType: string;
}

/** Internal note. */
export interface ContentTypeOverride {
  readonly partName: PartPath;
  readonly contentType: string;
}

export interface ContentTypes {
  readonly defaults: readonly ContentTypeDefault[];
  readonly overrides: readonly ContentTypeOverride[];
}

/** Package part reference (path + content type)。 */
export interface PackagePartRef {
  readonly partPath: PartPath;
  readonly contentType: string;
}

/** Internal note. */
export interface MediaPart {
  readonly partPath: PartPath;
  readonly contentType: string;
  readonly bytes: Uint8Array;
}

/**
 * Internal note.
 * Internal note.
 */
export interface PackageGraph {
  readonly contentTypes: ContentTypes;
  readonly parts: readonly PackagePartRef[];
  readonly relationships: readonly PartRelationships[];
  readonly media: readonly MediaPart[];
  /** Raw fallback for writing back parts that are not typed / were not edited. */
  readonly rawParts?: readonly RawPackagePart[];
}
