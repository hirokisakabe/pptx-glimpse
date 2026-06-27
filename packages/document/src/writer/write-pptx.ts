/**
 * `writePptx(source)` — PptxSourceModel source writer の最初の round-trip slice。
 *
 * この writer は structural round-trip を目的に、reader が保持した raw package
 * material / media bytes / package bookkeeping を PPTX ZIP として再構成する。
 * one plain text-run edit では dirty slide XML part を再シリアライズし、対象
 * run の `a:t` 値だけを stable source handle で差し替える。汎用的な edited
 * writer behavior や node-level XML splicing は後続 slice の責務。
 */

import { XMLBuilder } from "fast-xml-parser";
import { zipSync } from "fflate";

import { parseXml, type XmlNode } from "../reader/xml.js";
import type {
  PartPath,
  PartRelationships,
  PptxSourceModel,
  PptxSourceModelEdit,
  PptxSourceModelTextRunEdit,
  RawOoxmlNode,
  RawPackagePart,
} from "../source/index.js";

/** `writePptx` の出力。 */
export type WritePptxOutput = Uint8Array;

const CONTENT_TYPES_PART = "[Content_Types].xml";
const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: true,
});

/**
 * PptxSourceModel source を PPTX package bytes に書き戻す。
 *
 * round-trip 用の初期 writer であり、未編集 package material を優先して
 * preserved output を作る。dirty part の patch に必要な raw bytes が無い場合や
 * 必要な raw bytes が無い non-bookkeeping part は、
 * 暗黙に再生成せずエラーにする。
 */
export function writePptx(source: PptxSourceModel): WritePptxOutput {
  const textRunEdits = source.edits?.filter(isTextRunEdit) ?? [];
  const dirtyPartPaths = new Set(textRunEdits.map((edit) => edit.handle.partPath));
  const files: Record<string, Uint8Array> = {
    [CONTENT_TYPES_PART]: encodeXml(serializeContentTypes(source.packageGraph.contentTypes)),
  };

  const written = new Set<string>([CONTENT_TYPES_PART]);

  for (const relationships of source.packageGraph.relationships) {
    const relsPath = relationshipsPartPath(relationships.sourcePartPath);
    files[relsPath] = encodeXml(serializeRelationships(relationships));
    written.add(relsPath);
  }

  for (const media of source.packageGraph.media) {
    files[media.partPath] = media.bytes;
    written.add(media.partPath);
  }

  for (const rawPart of source.packageGraph.rawParts ?? []) {
    if (dirtyPartPaths.has(rawPart.partPath)) continue;
    files[rawPart.partPath] = serializeRawPackagePart(rawPart);
    written.add(rawPart.partPath);
  }

  for (const partPath of dirtyPartPaths) {
    files[partPath] = serializeDirtyXmlPart(source, partPath, textRunEdits);
    written.add(partPath);
  }

  for (const part of source.packageGraph.parts) {
    if (written.has(part.partPath)) continue;
    if (part.contentType === RELS_CONTENT_TYPE || isRelationshipPart(part.partPath)) continue;
    throw new Error(
      `writePptx: no preserved package material for part '${part.partPath}'; ` +
        "edited part generation is not implemented in the no-edit writer",
    );
  }

  return zipSync(files);
}

function isTextRunEdit(edit: PptxSourceModelEdit): edit is PptxSourceModelTextRunEdit {
  return edit.kind === "replaceTextRunPlainText";
}

function serializeDirtyXmlPart(
  source: PptxSourceModel,
  partPath: PartPath,
  textRunEdits: readonly PptxSourceModelTextRunEdit[],
): Uint8Array {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === partPath);
  if (rawPart === undefined) {
    throw new Error(`writePptx: dirty part '${partPath}' has no preserved raw package material`);
  }
  if (rawPart.kind !== "binary") {
    throw new Error(`writePptx: dirty XML tree part '${partPath}' patching is not implemented`);
  }

  const root = parseXml(textDecoder.decode(rawPart.bytes));
  const edits = textRunEdits.filter((edit) => edit.handle.partPath === partPath);
  for (const edit of edits) {
    applyTextRunEdit(root, edit);
  }
  return encodeXml(XML_DECLARATION + xmlBuilder.build(root));
}

function applyTextRunEdit(root: XmlNode, edit: PptxSourceModelTextRunEdit): void {
  const locator = parseTextRunLocator(edit.handle.nodeId);
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  const shape = locateShape(spTree, locator);
  const paragraphs = getChildArray(getChild(shape, "txBody"), "p");
  const paragraph = paragraphs[locator.paragraphIndex];
  const run = getChildArray(paragraph, "r")[locator.runIndex];

  if (run === undefined) {
    throw new Error(
      `writePptx: text run handle '${edit.handle.nodeId}' no longer matches source XML`,
    );
  }

  setChildText(run, "t", edit.text);
}

interface TextRunLocator {
  readonly shapeNodeId?: string;
  readonly shapeOrderingSlot?: number;
  readonly paragraphIndex: number;
  readonly runIndex: number;
}

function parseTextRunLocator(
  nodeId: PptxSourceModelTextRunEdit["handle"]["nodeId"],
): TextRunLocator {
  const value = String(nodeId ?? "");
  const byShapeId = /^text:shape:(.+):p:(\d+):r:(\d+)$/.exec(value);
  if (byShapeId !== null) {
    return {
      shapeNodeId: byShapeId[1],
      paragraphIndex: Number(byShapeId[2]),
      runIndex: Number(byShapeId[3]),
    };
  }

  const byShapeSlot = /^text:shapeSlot:(\d+):p:(\d+):r:(\d+)$/.exec(value);
  if (byShapeSlot !== null) {
    return {
      shapeOrderingSlot: Number(byShapeSlot[1]),
      paragraphIndex: Number(byShapeSlot[2]),
      runIndex: Number(byShapeSlot[3]),
    };
  }

  throw new Error(`writePptx: unsupported text run handle '${value}'`);
}

function locateShape(spTree: XmlNode | undefined, locator: TextRunLocator): XmlNode | undefined {
  const shapes = getChildArray(spTree, "sp");
  if (locator.shapeNodeId !== undefined) {
    return shapes.find(
      (shape) =>
        getAttr(getChild(getChild(shape, "nvSpPr"), "cNvPr"), "id") === locator.shapeNodeId,
    );
  }
  if (locator.shapeOrderingSlot === undefined) return undefined;
  return getShapeByOrderingSlot(spTree, locator.shapeOrderingSlot);
}

function getShapeByOrderingSlot(
  spTree: XmlNode | undefined,
  orderingSlot: number,
): XmlNode | undefined {
  if (!spTree) return undefined;

  let currentSlot = 0;
  for (const key of Object.keys(spTree)) {
    if (key.startsWith("@_")) continue;
    const local = localName(key);
    if (local === "nvGrpSpPr" || local === "grpSpPr") continue;

    const value = spTree[key];
    const items = Array.isArray(value) ? value : [value];
    for (const item of items) {
      if (currentSlot === orderingSlot) {
        return local === "sp" ? (item as XmlNode) : undefined;
      }
      currentSlot++;
    }
  }
  return undefined;
}

function getChild(node: XmlNode | undefined, name: string): XmlNode | undefined {
  if (!node) return undefined;
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    if (localName(key) !== name) continue;
    const value = node[key];
    return Array.isArray(value)
      ? (value[0] as XmlNode | undefined)
      : (value as XmlNode | undefined);
  }
  return undefined;
}

function getChildArray(node: XmlNode | undefined, name: string): XmlNode[] {
  if (!node) return [];
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    if (localName(key) === name) {
      const value = node[key];
      if (value === undefined || value === null) return [];
      return (Array.isArray(value) ? value : [value]) as XmlNode[];
    }
  }
  return [];
}

function getAttr(node: XmlNode | undefined, name: string): string | undefined {
  if (!node) return undefined;
  const value = node[`@_${name}`];
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return undefined;
}

function setChildText(node: XmlNode, name: string, text: string): void {
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

function textElementValue(existing: unknown, text: string): unknown {
  if (typeof existing === "object" && existing !== null && !Array.isArray(existing)) {
    const next: XmlNode = { ...(existing as XmlNode), "#text": text };
    if (textRequiresPreserve(text)) next["@_xml:space"] = "preserve";
    else delete next["@_xml:space"];
    return next;
  }
  return textRequiresPreserve(text) ? { "@_xml:space": "preserve", "#text": text } : text;
}

function textRequiresPreserve(text: string): boolean {
  return text.startsWith(" ") || text.endsWith(" ");
}

function localName(key: string): string {
  const colon = key.indexOf(":");
  return colon === -1 ? key : key.slice(colon + 1);
}

function serializeContentTypes(
  contentTypes: PptxSourceModel["packageGraph"]["contentTypes"],
): string {
  const defaults = contentTypes.defaults
    .map(
      (entry) =>
        `<Default Extension="${escapeAttribute(entry.extension)}" ` +
        `ContentType="${escapeAttribute(entry.contentType)}"/>`,
    )
    .join("");
  const overrides = contentTypes.overrides
    .map(
      (entry) =>
        `<Override PartName="/${escapeAttribute(entry.partName)}" ` +
        `ContentType="${escapeAttribute(entry.contentType)}"/>`,
    )
    .join("");
  return (
    XML_DECLARATION +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    defaults +
    overrides +
    `</Types>`
  );
}

function serializeRelationships(partRelationships: PartRelationships): string {
  const relationships = partRelationships.relationships
    .map((relationship) => {
      const targetMode =
        relationship.targetMode === undefined
          ? ""
          : ` TargetMode="${escapeAttribute(relationship.targetMode)}"`;
      return (
        `<Relationship Id="${escapeAttribute(relationship.id)}" ` +
        `Type="${escapeAttribute(relationship.type)}" ` +
        `Target="${escapeAttribute(relationship.target)}"${targetMode}/>`
      );
    })
    .join("");
  return (
    XML_DECLARATION +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    relationships +
    `</Relationships>`
  );
}

function serializeRawPackagePart(rawPart: RawPackagePart): Uint8Array {
  if (rawPart.kind === "binary") return rawPart.bytes;
  return encodeXml(XML_DECLARATION + serializeRawNode(rawPart.xml));
}

function serializeRawNode(node: RawOoxmlNode): string {
  const attributes =
    node.attributes === undefined
      ? ""
      : Object.entries(node.attributes)
          .map(([name, value]) => ` ${name}="${escapeAttribute(value)}"`)
          .join("");
  const text = node.text === undefined ? "" : escapeText(node.text);
  const children = node.children?.map((child) => serializeRawNode(child)).join("") ?? "";
  if (text !== "" && children !== "") {
    throw new Error(
      `writePptx: raw XML part '${node.name}' contains mixed text/element content; ` +
        "ordered mixed-content serialization is not implemented in the no-edit writer",
    );
  }
  if (text === "" && children === "") return `<${node.name}${attributes}/>`;
  return `<${node.name}${attributes}>${text}${children}</${node.name}>`;
}

function relationshipsPartPath(sourcePartPath: PartPath): string {
  if (sourcePartPath === "") return "_rels/.rels";
  const slash = sourcePartPath.lastIndexOf("/");
  const dir = slash === -1 ? "" : sourcePartPath.slice(0, slash + 1);
  const file = slash === -1 ? sourcePartPath : sourcePartPath.slice(slash + 1);
  return `${dir}_rels/${file}.rels`;
}

function isRelationshipPart(path: string): boolean {
  return path.endsWith(".rels") && path.includes("_rels/");
}

function encodeXml(xml: string): Uint8Array {
  return textEncoder.encode(xml);
}

function escapeAttribute(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeText(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}
