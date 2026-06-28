/**
 * Package graph の型。package part / relationship / content type / media を
 * source-model data として表す。
 */

import type { PartPath, RelationshipId } from "./handles.js";
import type { RawPackagePart } from "./raw.js";

/** relationship の target mode。 */
export type RelationshipTargetMode = "Internal" | "External";

/** `_rels/*.rels` の 1 エントリ。 */
export interface Relationship {
  readonly id: RelationshipId;
  /** relationship type URI。 */
  readonly type: string;
  /** sourcePartPath からの相対 target、もしくは external URL。 */
  readonly target: string;
  readonly targetMode?: RelationshipTargetMode;
}

/** ある source part が持つ relationship 群 (`<part>/_rels/<part>.rels`)。 */
export interface PartRelationships {
  readonly sourcePartPath: PartPath;
  readonly relationships: readonly Relationship[];
}

/** `[Content_Types].xml` の Default エントリ (拡張子 → content type)。 */
export interface ContentTypeDefault {
  readonly extension: string;
  readonly contentType: string;
}

/** `[Content_Types].xml` の Override エントリ (part 名 → content type)。 */
export interface ContentTypeOverride {
  readonly partName: PartPath;
  readonly contentType: string;
}

export interface ContentTypes {
  readonly defaults: readonly ContentTypeDefault[];
  readonly overrides: readonly ContentTypeOverride[];
}

/** package part の参照 (path + content type)。 */
export interface PackagePartRef {
  readonly partPath: PartPath;
  readonly contentType: string;
}

/** media asset (画像・音声・動画等) の part。bytes をそのまま保持する。 */
export interface MediaPart {
  readonly partPath: PartPath;
  readonly contentType: string;
  readonly bytes: Uint8Array;
}

/**
 * package 全体の構造。content types / part 一覧 / part ごとの relationship /
 * media、および未編集 part の raw fallback を保持する。
 */
export interface PackageGraph {
  readonly contentTypes: ContentTypes;
  readonly parts: readonly PackagePartRef[];
  readonly relationships: readonly PartRelationships[];
  readonly media: readonly MediaPart[];
  /** typed に表現しない / 未編集の part を書き戻すための raw fallback。 */
  readonly rawParts?: readonly RawPackagePart[];
}
