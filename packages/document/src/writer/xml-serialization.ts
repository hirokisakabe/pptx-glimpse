import { XMLBuilder, XMLParser } from "fast-xml-parser";

import { parseXml, type XmlNode } from "../reader/xml.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const textEncoder = new TextEncoder();
export const textDecoder = new TextDecoder();

export const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: true,
});

const orderedXmlParser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: false,
  trimValues: false,
});

const orderedXmlBuilder = new XMLBuilder({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: true,
});

interface XmlChildOrderEntry {
  readonly key: string;
  readonly value: unknown;
}

const xmlChildOrder = new WeakMap<XmlNode, readonly XmlChildOrderEntry[]>();

/** Parse editable XML while retaining the order of heterogeneous child elements. */
export function parseXmlForEditing(xml: string): XmlNode {
  const root = parseXml(xml);
  const ordered = unsafeOoxmlBoundaryAssertion<XmlNode[]>(orderedXmlParser.parse(xml));
  rememberChildOrder(root, ordered);
  return root;
}

/** Record an appended child in both the grouped edit tree and its ordered sidecar. */
export function appendXmlChildAtEnd(node: XmlNode, key: string, value: XmlNode): void {
  const existing = node[key];
  if (existing === undefined) node[key] = [value];
  else if (Array.isArray(existing)) existing.push(value);
  else node[key] = [existing, value];
  xmlChildOrder.set(node, [...(xmlChildOrder.get(node) ?? []), { key, value }]);
}

/** Serialize an edited XML tree without grouping heterogeneous siblings by tag name. */
export function buildXmlPreservingChildOrder(root: XmlNode): string {
  return orderedXmlBuilder.build(toOrderedChildren(root, true));
}

export function encodeXml(xml: string): Uint8Array {
  return textEncoder.encode(xml);
}

function rememberChildOrder(node: XmlNode, orderedChildren: readonly XmlNode[]): void {
  const occurrences = new Map<string, number>();
  const order: XmlChildOrderEntry[] = [];

  for (const orderedEntry of orderedChildren) {
    const key = Object.keys(orderedEntry).find((entryKey) => entryKey !== ":@");
    if (key === undefined || key === "?xml") continue;
    if (key === "#text") {
      order.push({ key, value: node[key] });
      continue;
    }

    const occurrence = occurrences.get(key) ?? 0;
    occurrences.set(key, occurrence + 1);
    const groupedValue = node[key];
    const value = Array.isArray(groupedValue)
      ? unsafeOoxmlBoundaryAssertion<unknown[]>(groupedValue)[occurrence]
      : groupedValue;
    if (value === undefined) continue;
    order.push({ key, value });

    if (isXmlNode(value)) {
      const childOrder = orderedEntry[key];
      if (Array.isArray(childOrder)) {
        rememberChildOrder(value, unsafeOoxmlBoundaryAssertion<XmlNode[]>(childOrder));
      }
    }
  }

  xmlChildOrder.set(node, order);
}

function toOrderedChildren(node: XmlNode, root = false): XmlNode[] {
  const current = flattenChildren(node);
  const unused = new Set(current);
  const ordered = (xmlChildOrder.get(node) ?? []).flatMap((remembered) => {
    const exact = current.find(
      (child) =>
        unused.has(child) && child.key === remembered.key && child.value === remembered.value,
    );
    const replacement =
      exact ?? current.find((child) => unused.has(child) && child.key === remembered.key);
    if (replacement === undefined) return [];
    unused.delete(replacement);
    return [replacement];
  });

  for (const child of current) {
    if (!unused.has(child)) continue;
    const currentIndex = current.indexOf(child);
    const following = current
      .slice(currentIndex + 1)
      .find((candidate) => ordered.includes(candidate));
    if (following === undefined) ordered.push(child);
    else ordered.splice(ordered.indexOf(following), 0, child);
    unused.delete(child);
  }

  return ordered
    .map(({ key, value }) => {
      if (key === "#text") return { "#text": value };
      const children = isXmlNode(value)
        ? toOrderedChildren(value)
        : value === "" || value === undefined || value === null
          ? []
          : [{ "#text": value }];
      const entry: XmlNode = { [key]: children };
      if (isXmlNode(value)) {
        const attributes = Object.fromEntries(
          Object.entries(value).filter(([attributeKey]) => attributeKey.startsWith("@_")),
        );
        if (Object.keys(attributes).length > 0) entry[":@"] = attributes;
      }
      return entry;
    })
    .filter((entry) => !root || !("?xml" in entry));
}

interface FlattenedChild {
  readonly key: string;
  readonly value: unknown;
}

function flattenChildren(node: XmlNode): FlattenedChild[] {
  return Object.entries(node).flatMap(([key, value]) => {
    if (key.startsWith("@_") || key === "?xml") return [];
    const values = Array.isArray(value) ? unsafeOoxmlBoundaryAssertion<unknown[]>(value) : [value];
    return values.map((childValue) => ({
      key,
      value: childValue,
    }));
  });
}

function isXmlNode(value: unknown): value is XmlNode {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
