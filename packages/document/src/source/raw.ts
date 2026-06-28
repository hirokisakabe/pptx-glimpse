/**
 * Types for the raw OOXML escape hatch.
 *
 * PptxSourceModel represents supported semantics as typed fields, while preserving
 * unsupported or partially supported XML as raw sidecars and unedited parts as raw
 * package parts. This enables structural round-tripping; byte equality is not a goal.
 */

import type { PartPath, RawSidecarId } from "./handles.js";

/**
 * Partially parsed raw OOXML node. It preserves the namespace-prefixed qualified name,
 * attributes, child nodes, and text for elements that PptxSourceModel does not represent
 * as typed fields, such as vendor extensions, `mc:AlternateContent`, and unknown
 * DrawingML.
 */
export interface RawOoxmlNode {
  /** Element name including the namespace prefix (Example: `a:extLst`). */
  readonly name: string;
  readonly attributes?: Readonly<Record<string, string>>;
  readonly children?: readonly RawOoxmlNode[];
  /** Element text content (when present). */
  readonly text?: string;
}

/**
 * Raw XML sidecar associated with a PptxSourceModel source node. It is attached to the nearest source node
 * and keeps ordering metadata within the parent element to restore order during write-back.
 */
export interface RawSidecar {
  readonly id: RawSidecarId;
  readonly node: RawOoxmlNode;
  /** Ordering slot within the owning element's child sequence. */
  readonly orderingSlot?: number;
}

/**
 * Raw fallback for writing an unedited package part back as-is. It keeps either bytes
 * for binary assets or an XML tree as a discriminated union. The type rules out invalid
 * states with both or neither representation.
 */
export type RawPackagePart =
  | {
      readonly kind: "binary";
      readonly partPath: PartPath;
      readonly contentType: string;
      /** Original byte string of binary part (image, embedded workbook, etc.). */
      readonly bytes: Uint8Array;
    }
  | {
      readonly kind: "xml";
      readonly partPath: PartPath;
      readonly contentType: string;
      /** The root node when storing the XML part as a tree. */
      readonly xml: RawOoxmlNode;
    };
