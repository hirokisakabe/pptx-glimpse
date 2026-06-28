import { unsafeBrandAssertion } from "../unsafe-type-assertion.js";

/**
 * Source handle related types.
 *
 * PptxSourceModel source node is a writer / editor / round-trip
 * In order to do this, we can use ``which package part and which node it came from'' as a stable handle.
 * A handle is not a direct mutable pointer, but a stable reference
 * (see part path / relationship id / node id / order within siblings / raw sidecar)
 * composes it.
 */

declare const PartPathBrand: unique symbol;
declare const RelationshipIdBrand: unique symbol;
declare const SourceNodeIdBrand: unique symbol;
declare const RawSidecarIdBrand: unique symbol;

/** Part path inside the OOXML package (Example: `ppt/slides/slide1.xml`). */
export type PartPath = string & { readonly [PartPathBrand]: typeof PartPathBrand };

/** Relationship id (Example: `rId1`).`_rels/*.rels` source. */
export type RelationshipId = string & {
  readonly [RelationshipIdBrand]: typeof RelationshipIdBrand;
};

/** An id that uniquely points to an element within the source part (spid / generation id, etc.). */
export type SourceNodeId = string & { readonly [SourceNodeIdBrand]: typeof SourceNodeIdBrand };

/** Id pointing to a raw sidecar. */
export type RawSidecarId = string & { readonly [RawSidecarIdBrand]: typeof RawSidecarIdBrand };

export function asPartPath(value: string): PartPath {
  return unsafeBrandAssertion<PartPath>(value);
}

export function asRelationshipId(value: string): RelationshipId {
  return unsafeBrandAssertion<RelationshipId>(value);
}

export function asSourceNodeId(value: string): SourceNodeId {
  return unsafeBrandAssertion<SourceNodeId>(value);
}

export function asRawSidecarId(value: string): RawSidecarId {
  return unsafeBrandAssertion<RawSidecarId>(value);
}

/**
 * Stable handle describing a source node's provenance. The writer uses the handle to decide whether a generated
 * node can be spliced into an existing part or whether a broader scope must be regenerated.
 *
 */
export interface SourceHandle {
  /** Package part that owns this node. */
  readonly partPath: PartPath;
  /** Node id inside the part (when available). */
  readonly nodeId?: SourceNodeId;
  /** The relationship id referencing this node (e.g. `r:embed` in blip). */
  readonly relationshipId?: RelationshipId;
  /** The child's order slot within the parent element. Used to restore order with raw sidecar. */
  readonly orderingSlot?: number;
  /** Raw sidecar ids associated with this node. */
  readonly rawSidecarIds?: readonly RawSidecarId[];
}
