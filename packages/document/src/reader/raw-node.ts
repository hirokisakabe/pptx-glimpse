/**
 * Conversion helpers for the raw OOXML escape hatch.
 *
 * Internal note.
 * Internal note.
 * Internal note.
 * .
 */

import type { RawOoxmlNode, RawSidecar, RawSidecarId } from "../source/index.js";
import { asRawSidecarId } from "../source/index.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import { localName, type XmlNode } from "./xml.js";

const ATTR_PREFIX = "@_";
const TEXT_KEY = "#text";

/** Creates a factory that issues stable raw sidecar ids per part. */
export function createSidecarIdFactory(partPath: string): () => RawSidecarId {
  let counter = 0;
  return () => asRawSidecarId(`${partPath}#raw-${counter++}`);
}

/**
 * Internal note.
 * qualified name (Example: `a:effectLst`)。attributes are kept after removing the `@_` prefix、
 * Internal note.
 */
function xmlValueToRawNode(name: string, value: unknown): RawOoxmlNode {
  if (typeof value !== "object" || value === null) {
    const text = scalarText(value);
    return text !== undefined && text !== "" ? { name, text } : { name };
  }

  const obj = unsafeOoxmlBoundaryAssertion<Record<string, unknown>>(value);
  const attributes: Record<string, string> = {};
  const children: RawOoxmlNode[] = [];
  let text: string | undefined;

  for (const key of Object.keys(obj)) {
    if (key === TEXT_KEY) {
      text = scalarText(obj[key]);
      continue;
    }
    if (key.startsWith(ATTR_PREFIX)) {
      const attrValue = scalarText(obj[key]);
      if (attrValue !== undefined) attributes[key.slice(ATTR_PREFIX.length)] = attrValue;
      continue;
    }
    const childValue = obj[key];
    const items = Array.isArray(childValue) ? childValue : [childValue];
    for (const item of items) {
      children.push(xmlValueToRawNode(key, item));
    }
  }

  return {
    name,
    ...(Object.keys(attributes).length > 0 ? { attributes } : {}),
    ...(children.length > 0 ? { children } : {}),
    ...(text !== undefined && text !== "" ? { text } : {}),
  };
}

/** Single child element (qualified `name` / value) to one raw sidecar. */
export function makeSidecar(
  name: string,
  value: unknown,
  nextId: () => RawSidecarId,
  orderingSlot?: number,
): RawSidecar {
  return {
    id: nextId(),
    node: xmlValueToRawNode(name, value),
    ...(orderingSlot !== undefined ? { orderingSlot } : {}),
  };
}

/**
 * Internal note.
 * Internal note.
 * Internal note.
 */
export function collectUnknownSidecars(
  parent: XmlNode | undefined,
  knownLocalNames: ReadonlySet<string>,
  nextId: () => RawSidecarId,
): RawSidecar[] {
  if (!parent) return [];
  const sidecars: RawSidecar[] = [];
  let slot = 0;
  for (const key of Object.keys(parent)) {
    if (key.startsWith(ATTR_PREFIX) || key === TEXT_KEY) continue;
    const value = parent[key];
    const items = Array.isArray(value) ? value : [value];
    if (knownLocalNames.has(localName(key))) {
      slot += items.length;
      continue;
    }
    for (const item of items) {
      sidecars.push(makeSidecar(key, item, nextId, slot));
      slot++;
    }
  }
  return sidecars;
}

function scalarText(value: unknown): string | undefined {
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return undefined;
}
