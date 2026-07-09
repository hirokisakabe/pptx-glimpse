import { getChild, localName, type XmlNode } from "../reader/xml.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";

export function namespacedChildKey(node: XmlNode, fallback: string, local: string): string {
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    if (localName(key) === local) return key;
  }
  return fallback;
}

export function namespacedAttributeKey(node: XmlNode, fallback: string, local: string): string {
  for (const key of Object.keys(node)) {
    if (!key.startsWith("@_")) continue;
    const name = key.slice(2);
    const colon = name.indexOf(":");
    if (colon !== -1 && name.slice(colon + 1) === local) return key;
  }
  return `@_${fallback}`;
}

export function stripXmlProcessingInstruction(root: XmlNode): XmlNode {
  const stripped = { ...root };
  delete stripped["?xml"];
  return stripped;
}

export function setChildText(node: XmlNode, name: string, text: string): void {
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    if (localName(key) !== name) continue;
    const value = node[key];
    if (Array.isArray(value)) {
      value[0] = textElementValue(value[0], text);
      return;
    }
    node[key] = textElementValue(value, text);
    return;
  }
  node[`a:${name}`] = textRequiresPreserve(text)
    ? { "@_xml:space": "preserve", "#text": text }
    : text;
}

export function ensureChild(node: XmlNode, name: string): XmlNode {
  const existing = getChild(node, name);
  if (existing !== undefined) return existing;
  node[`a:${name}`] = {};
  return unsafeOoxmlBoundaryAssertion<XmlNode>(node[`a:${name}`]);
}

export function replaceChild(node: XmlNode, name: string, value: XmlNode): void {
  const entries: [string, unknown][] = [];
  let replaced = false;
  for (const [key, entryValue] of Object.entries(node)) {
    if (!key.startsWith("@_") && localName(key) === name) {
      if (!replaced) entries.push([key, value]);
      replaced = true;
      continue;
    }
    entries.push([key, entryValue]);
  }
  if (!replaced) entries.push([`a:${name}`, value]);
  replaceNodeEntries(node, entries);
}

export function deleteChild(node: XmlNode, name: string): boolean {
  let deleted = false;
  replaceNodeEntries(
    node,
    Object.entries(node).filter(([key]) => {
      const keep = key.startsWith("@_") || localName(key) !== name;
      if (!keep) deleted = true;
      return keep;
    }),
  );
  return deleted;
}

export function xmlNodeIsEmpty(node: XmlNode): boolean {
  return Object.keys(node).length === 0;
}

export function replaceNodeEntries(node: XmlNode, entries: readonly [string, unknown][]): void {
  for (const key of Object.keys(node)) delete node[key];
  for (const [key, value] of entries) node[key] = value;
}

export function cloneXmlNode(node: XmlNode): XmlNode {
  return Object.fromEntries(
    Object.entries(node).map(([key, value]) => [key, cloneXmlValue(value)]),
  );
}

function cloneXmlValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(cloneXmlValue);
  if (typeof value === "object" && value !== null) {
    return cloneXmlNode(unsafeOoxmlBoundaryAssertion<XmlNode>(value));
  }
  return value;
}

function textElementValue(existing: unknown, text: string): unknown {
  if (typeof existing === "object" && existing !== null && !Array.isArray(existing)) {
    const next: XmlNode = { ...unsafeOoxmlBoundaryAssertion<XmlNode>(existing), "#text": text };
    if (textRequiresPreserve(text)) next["@_xml:space"] = "preserve";
    else delete next["@_xml:space"];
    return next;
  }
  return textRequiresPreserve(text) ? { "@_xml:space": "preserve", "#text": text } : text;
}

export function textRequiresPreserve(text: string): boolean {
  return text.startsWith(" ") || text.endsWith(" ");
}
