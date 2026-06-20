/**
 * Source handle 関連の型。
 *
 * CleanDoc source node は、書き戻し (writer) / 編集 (editor) / round-trip の
 * ために「どの package part のどのノードから来たか」を stable な handle として
 * 保持する。handle は直接の mutable pointer ではなく、安定した参照
 * (part path / relationship id / node id / 兄弟内の順序 / raw sidecar 参照) で
 * 構成する (`docs/raw-ooxml-round-trip.md` の Raw Preservation Granularity 参照)。
 */

declare const PartPathBrand: unique symbol;
declare const RelationshipIdBrand: unique symbol;
declare const SourceNodeIdBrand: unique symbol;
declare const RawSidecarIdBrand: unique symbol;

/** OOXML package 内の part path (例: `ppt/slides/slide1.xml`)。 */
export type PartPath = string & { readonly [PartPathBrand]: typeof PartPathBrand };

/** relationship id (例: `rId1`)。`_rels/*.rels` 由来。 */
export type RelationshipId = string & {
  readonly [RelationshipIdBrand]: typeof RelationshipIdBrand;
};

/** source part 内で要素を一意に指す id (spid / 生成 id 等)。 */
export type SourceNodeId = string & { readonly [SourceNodeIdBrand]: typeof SourceNodeIdBrand };

/** raw sidecar を指す id。 */
export type RawSidecarId = string & { readonly [RawSidecarIdBrand]: typeof RawSidecarIdBrand };

export function asPartPath(value: string): PartPath {
  return value as PartPath;
}

export function asRelationshipId(value: string): RelationshipId {
  return value as RelationshipId;
}

export function asSourceNodeId(value: string): SourceNodeId {
  return value as SourceNodeId;
}

export function asRawSidecarId(value: string): RawSidecarId {
  return value as RawSidecarId;
}

/**
 * source node の出自を表す stable handle。writer は handle を見て、生成した
 * ノードを既存 part に splice できるか、より広い scope を再生成すべきかを
 * 判断する。
 */
export interface SourceHandle {
  /** このノードを所有する package part。 */
  readonly partPath: PartPath;
  /** part 内でのノード id (取得できる場合)。 */
  readonly nodeId?: SourceNodeId;
  /** このノードを参照している relationship id (例: blip の `r:embed`)。 */
  readonly relationshipId?: RelationshipId;
  /** 親要素内での子の順序スロット。raw sidecar との順序復元に使う。 */
  readonly orderingSlot?: number;
  /** このノードに紐づく raw sidecar の id 群。 */
  readonly rawSidecarIds?: readonly RawSidecarId[];
}
