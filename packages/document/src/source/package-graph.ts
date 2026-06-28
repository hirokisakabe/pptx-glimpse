/**
 * Package graph type. package part / relationship / content type / media
 * Represented as source-model data.
 */

import type { PartPath, RelationshipId } from "./handles.js";
import type { RawPackagePart } from "./raw.js";

/** target mode of relationship. */
export type RelationshipTargetMode = "Internal" | "External";

/** 1 entry in `_rels/*.rels`. */
export interface Relationship {
  readonly id: RelationshipId;
  /** relationship type URI. */
  readonly type: string;
  /** Target relative to sourcePartPath, or an external URL. */
  readonly target: string;
  readonly targetMode?: RelationshipTargetMode;
}

/** Relationships owned by a source part (`<part>/_rels/<part>.rels`). */
export interface PartRelationships {
  readonly sourcePartPath: PartPath;
  readonly relationships: readonly Relationship[];
}

/** Default entry in `[Content_Types].xml` (extension -> content type). */
export interface ContentTypeDefault {
  readonly extension: string;
  readonly contentType: string;
}

/** Override entry (part name -> content type) in `[Content_Types].xml`. */
export interface ContentTypeOverride {
  readonly partName: PartPath;
  readonly contentType: string;
}

export interface ContentTypes {
  readonly defaults: readonly ContentTypeDefault[];
  readonly overrides: readonly ContentTypeOverride[];
}

/** Package part reference (path + content type). */
export interface PackagePartRef {
  readonly partPath: PartPath;
  readonly contentType: string;
}

/** Part of a media asset (image, audio, video, etc.). Keep bytes as is. */
export interface MediaPart {
  readonly partPath: PartPath;
  readonly contentType: string;
  readonly bytes: Uint8Array;
}

/**
 * The overall structure of the package. content types / part list / relationship by part /
 * media, and retain the raw fallback of the unedited part.
 */
export interface PackageGraph {
  readonly contentTypes: ContentTypes;
  readonly parts: readonly PackagePartRef[];
  readonly relationships: readonly PartRelationships[];
  readonly media: readonly MediaPart[];
  /** Raw fallback for writing back parts that are not typed / were not edited. */
  readonly rawParts?: readonly RawPackagePart[];
}
