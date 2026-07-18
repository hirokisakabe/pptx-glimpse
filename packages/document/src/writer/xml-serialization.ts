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
  return orderedXmlBuilder.build(toOrderedChildren(root));
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

function toOrderedChildren(node: XmlNode): XmlNode[] {
  const current = flattenChildren(node);
  const currentIndexes = indexCurrentChildren(current);
  const matchedIndexes = (xmlChildOrder.get(node) ?? []).flatMap((remembered) => {
    const indexes = currentIndexes.get(remembered.key)?.get(remembered.value);
    if (indexes === undefined || indexes.cursor >= indexes.values.length) return [];
    const index = indexes.values[indexes.cursor];
    indexes.cursor += 1;
    return [index];
  });
  const matched = new Set(matchedIndexes);
  const unmatchedBefore = new Map<number | undefined, FlattenedChild[]>();
  let followingMatchedIndex: number | undefined;
  for (let index = current.length - 1; index >= 0; index -= 1) {
    if (matched.has(index)) {
      followingMatchedIndex = index;
      continue;
    }
    const bucket = unmatchedBefore.get(followingMatchedIndex) ?? [];
    bucket.unshift(current[index]);
    unmatchedBefore.set(followingMatchedIndex, bucket);
  }
  const ordered = matchedIndexes.flatMap((index) => [
    ...(unmatchedBefore.get(index) ?? []),
    current[index],
  ]);
  ordered.push(...(unmatchedBefore.get(undefined) ?? []));

  return ordered.map(({ key, value }) => {
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
  });
}

interface FlattenedChild {
  readonly key: string;
  readonly value: unknown;
}

interface CurrentChildIndexes {
  readonly values: number[];
  cursor: number;
}

function indexCurrentChildren(
  children: readonly FlattenedChild[],
): Map<string, Map<unknown, CurrentChildIndexes>> {
  const indexed = new Map<string, Map<unknown, CurrentChildIndexes>>();
  children.forEach((child, index) => {
    const byValue = indexed.get(child.key) ?? new Map<unknown, CurrentChildIndexes>();
    indexed.set(child.key, byValue);
    const indexes = byValue.get(child.value) ?? { values: [], cursor: 0 };
    indexes.values.push(index);
    byValue.set(child.value, indexes);
  });
  return indexed;
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
