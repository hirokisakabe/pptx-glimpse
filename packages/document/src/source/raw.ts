/**
 * Internal note.
 *
 * Internal note.
 * Internal note.
 * enabling structural round trips. Byte equality is not a goal.
 */

import type { PartPath, RawSidecarId } from "./handles.js";

/**
 * Partially parsed raw OOXML node. It preserves the namespace-prefixed qualified name,
 * Internal note.
 * Internal note.
 */
export interface RawOoxmlNode {
  /** Element name including the namespace prefix (Example: `a:extLst`)。 */
  readonly name: string;
  readonly attributes?: Readonly<Record<string, string>>;
  readonly children?: readonly RawOoxmlNode[];
  /** Element text content (when present)。 */
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
 * Raw fallback for writing an unedited package part back as-is. It keeps either bytes (binary
 * asset) or an XML tree as a discriminated union. Having both /
 * having neither is ruled out by the type.
 */
export type RawPackagePart =
  | {
      readonly kind: "binary";
      readonly partPath: PartPath;
      readonly contentType: string;
      /** Internal note. */
      readonly bytes: Uint8Array;
    }
  | {
      readonly kind: "xml";
      readonly partPath: PartPath;
      readonly contentType: string;
      /** Internal note. */
      readonly xml: RawOoxmlNode;
    };
