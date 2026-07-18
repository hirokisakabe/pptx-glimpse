import { getChild, localName, type XmlNode } from "../reader/xml.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import { replaceNodeEntries } from "./xml-node-utils.js";
import { appendXmlChildAtEnd, parseXmlForEditing } from "./xml-serialization.js";

export function getDrawingPartRoot(root: XmlNode): XmlNode | undefined {
  return getChild(root, "sld") ?? getChild(root, "sldLayout") ?? getChild(root, "sldMaster");
}

export function ensurePictureNamespaces(drawingPart: XmlNode): void {
  drawingPart["@_xmlns:a"] ??= "http://schemas.openxmlformats.org/drawingml/2006/main";
  drawingPart["@_xmlns:r"] ??=
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
}

export function appendShapeTreeNodeAtEnd(
  spTree: XmlNode,
  preferredKey: string,
  value: XmlNode,
): void {
  const local = localName(preferredKey);
  let existingKey = preferredKey;

  for (const key of Object.keys(spTree)) {
    if (!key.startsWith("@_") && localName(key) === local) existingKey = key;
  }
  appendXmlChildAtEnd(spTree, existingKey, value);
}

export function insertChildByOrder(
  node: XmlNode,
  key: string,
  value: XmlNode,
  shouldInsertBefore: (local: string) => boolean,
): void {
  const entries: [string, unknown][] = [];
  let inserted = false;
  for (const [entryKey, entryValue] of Object.entries(node)) {
    if (!inserted && !entryKey.startsWith("@_") && shouldInsertBefore(localName(entryKey))) {
      entries.push([key, value]);
      inserted = true;
    }
    entries.push([entryKey, entryValue]);
  }
  if (!inserted) entries.push([key, value]);
  replaceNodeEntries(node, entries);
}

export function qualifiedSiblingName(siblingKey: string, local: string): string {
  const prefix = elementPrefix(siblingKey);
  return prefix === "" ? local : `${prefix}:${local}`;
}

export function remapElementPrefix(value: unknown, from: string, to: string): unknown {
  if (Array.isArray(value)) return value.map((entry) => remapElementPrefix(entry, from, to));
  if (typeof value !== "object" || value === null) return value;
  return Object.fromEntries(
    Object.entries(unsafeOoxmlBoundaryAssertion<XmlNode>(value)).map(([key, entry]) => {
      const remappedKey = key.startsWith(`${from}:`)
        ? to === ""
          ? localName(key)
          : `${to}:${localName(key)}`
        : key;
      return [remappedKey, remapElementPrefix(entry, from, to)];
    }),
  );
}

export function preserveNamespaceDeclarations(
  existing: XmlNode | undefined,
  replacement: unknown,
): unknown {
  if (existing === undefined || typeof replacement !== "object" || replacement === null) {
    return replacement;
  }
  const declarations = Object.entries(existing).filter(
    ([key]) => key === "@_xmlns" || key.startsWith("@_xmlns:"),
  );
  if (declarations.length === 0) return replacement;
  return {
    ...Object.fromEntries(declarations),
    ...unsafeOoxmlBoundaryAssertion<XmlNode>(replacement),
  };
}

export function parseShapeFragmentXml(
  xml: string,
  rootLocalName: "sp" | "cxnSp" | "pic" | "graphicFrame",
): XmlNode {
  const node = getChild(parseXmlForEditing(xml), rootLocalName);
  if (node === undefined) {
    throw new Error(
      `writePptx: shape edit XML fragment does not contain a '${rootLocalName}' root element`,
    );
  }
  return node;
}

export function elementPrefix(key: string): string {
  const separatorIndex = key.indexOf(":");
  return separatorIndex === -1 ? "" : key.slice(0, separatorIndex);
}
