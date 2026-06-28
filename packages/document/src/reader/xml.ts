/**
 * Minimal OOXML XML helpers for document readers.
 *
 * `@pptx-glimpse/document` owns the OOXML reader boundary as a lower-level
 * package, so it parses OOXML parts directly instead of depending on parsing
 * helpers from `@pptx-glimpse/core` or `@pptx-glimpse/renderer`.
 *
 * The object parser keeps namespace prefixes (`removeNSPrefix: false`) because
 * elements such as `p:sldId` can carry both a plain `id` attribute (slide ID)
 * and a relationships `r:id` attribute (relationship reference). Dropping the
 * prefix would collapse those attributes into the same key and lose the
 * relationship reference. Element lookup therefore uses local names while
 * attribute lookup distinguishes plain and namespaced attributes.
 *
 * The ordered parser is used only when reader logic needs element order. It
 * intentionally normalizes prefixes and ignores attributes, so it must not be
 * used for relationship IDs or other attribute-sensitive OOXML data.
 */

import { XMLParser } from "fast-xml-parser";

import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";

export type XmlNode = Record<string, unknown>;
export type XmlOrderedNode = Record<string, unknown>;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseAttributeValue: false,
  // Retain prefix. See the comment at the beginning of the file for the reason.
  removeNSPrefix: false,
  // Do not trim text run (`a:t`) to preserve significant white space at the beginning and end.
  // The PPTX part has been minified, and the blank text from the indentation between the tags is
  // This does not cause spurious text node contamination.
  trimValues: false,
});

const orderedParser = new XMLParser({
  preserveOrder: true,
  removeNSPrefix: true,
  ignoreAttributes: true,
  trimValues: false,
});

/** Parse an XML string and return the root object. */
export function parseXml(xml: string): XmlNode {
  return unsafeOoxmlBoundaryAssertion<XmlNode>(parser.parse(xml));
}

export function parseXmlOrdered(xml: string): XmlOrderedNode[] {
  return unsafeOoxmlBoundaryAssertion<XmlOrderedNode[]>(orderedParser.parse(xml));
}

export function navigateOrdered(
  ordered: readonly XmlOrderedNode[],
  path: readonly string[],
): XmlOrderedNode[] | undefined {
  let current: readonly XmlOrderedNode[] = ordered;
  for (const key of path) {
    const entry = current.find((item) => key in item);
    const value = entry?.[key];
    if (!Array.isArray(value)) return undefined;
    current = unsafeOoxmlBoundaryAssertion<XmlOrderedNode[]>(value);
  }
  return [...current];
}

/** Extracts the local part (`foo`) from a qualified name such as `a:foo`. */
export function localName(key: string): string {
  const colon = key.indexOf(":");
  return colon === -1 ? key : key.slice(colon + 1);
}

/**
 * Get child element by local name (ignoring prefix). Attribute keys (`@_`) are not applicable.
 * If there are multiple elements with the same name, the first match is returned.
 */
export function getChild(node: XmlNode | undefined, name: string): XmlNode | undefined {
  if (!node) return undefined;
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    if (localName(key) === name) {
      const value = node[key];
      return Array.isArray(value)
        ? unsafeOoxmlBoundaryAssertion<XmlNode | undefined>(value[0])
        : unsafeOoxmlBoundaryAssertion<XmlNode | undefined>(value);
    }
  }
  return undefined;
}

/**
 * Determines whether a child element exists by local name.
 *
 * Empty elements such as `<a:noFill/>` can parse as an empty string, so truthiness of
 * `getChild` is not enough. Use this for marker elements whose presence is meaningful.
 */
export function hasChild(node: XmlNode | undefined, name: string): boolean {
  if (!node) return false;
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    if (localName(key) === name) return true;
  }
  return false;
}

/** Get child elements by local name and always return them as an array. */
export function getChildArray(node: XmlNode | undefined, name: string): XmlNode[] {
  if (!node) return [];
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    if (localName(key) === name) {
      const value = node[key];
      if (value === undefined || value === null) return [];
      return unsafeOoxmlBoundaryAssertion<XmlNode[]>(Array.isArray(value) ? value : [value]);
    }
  }
  return [];
}

/** Get an attribute without namespace (`@_<name>`). */
export function getAttr(node: XmlNode | undefined, name: string): string | undefined {
  if (!node) return undefined;
  return scalarToString(node[`@_${name}`]);
}

/**
 * Get namespaced attribute (`@_<prefix>:<localName>`). `p:sldId`
 * To extract relationship references as distinct from plain `id`, such as `r:id`.
 * use. Regardless of the prefix, returns the first attribute that matches the local part.
 */
export function getNamespacedAttr(
  node: XmlNode | undefined,
  localAttr: string,
): string | undefined {
  if (!node) return undefined;
  for (const key of Object.keys(node)) {
    if (!key.startsWith("@_")) continue;
    const attr = key.slice(2);
    const colon = attr.indexOf(":");
    if (colon !== -1 && attr.slice(colon + 1) === localAttr) {
      const value = scalarToString(node[key]);
      if (value !== undefined) return value;
    }
  }
  return undefined;
}

/**
 * Get the text content of the child element by local name. Like `<a:t>foo</a:t>`
 * It corresponds to text nodes, `#text` elements with attributes, and empty elements.
 */
export function getChildText(node: XmlNode | undefined, name: string): string | undefined {
  if (!node) return undefined;
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    if (localName(key) !== name) continue;
    const value = node[key];
    const item: unknown = Array.isArray(value) ? value[0] : value;
    if (typeof item === "string") return item;
    if (typeof item === "number" || typeof item === "boolean") return String(item);
    if (item && typeof item === "object") {
      return scalarToString(unsafeOoxmlBoundaryAssertion<XmlNode>(item)["#text"]);
    }
    return undefined;
  }
  return undefined;
}

/**
 * Return all attributes of the element as a record of `{ name: value }` (remove the `@_` prefix).
 * Used to store an entire attribute set like the logical-name mapping of `p:clrMap`.
 */
export function getAttrs(node: XmlNode | undefined): Record<string, string> {
  const result: Record<string, string> = {};
  if (!node) return result;
  for (const key of Object.keys(node)) {
    if (!key.startsWith("@_")) continue;
    const value = scalarToString(node[key]);
    if (value !== undefined) result[key.slice(2)] = value;
  }
  return result;
}

/** Convert attribute value (string/number/boolean) into a string. object etc. are undefined. */
function scalarToString(value: unknown): string | undefined {
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return undefined;
}
