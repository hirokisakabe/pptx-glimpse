/**
 * `writePptx(source)` - the first round-trip slice of the PptxSourceModel source writer.
 *
 * The writer targets structural round-trip preservation rather than byte equality. It is
 * not a package patcher that preserves XML attribute order, namespace prefix placement,
 * ZIP metadata, or defaulted OOXML values. Content types and relationships are
 * regenerated structurally from `packageGraph`; media bytes, unknown parts, and
 * non-bookkeeping raw parts prefer the raw package material preserved by the reader.
 * Only dirty scopes are updated according to supported PptxSourceModel operations.
 *
 * The current slice supports one plain text-run edit, reserializing the dirty slide XML
 * part and replacing only the target run's `a:t` value via a stable source handle.
 * Node-level XML splicing, precise unsupported raw-sidecar invalidation, and package
 * topology rewrites belong to later writer slices, but the API and dirty-scope tracking
 * remain shaped for that extension path.
 */

import { XMLBuilder } from "fast-xml-parser";
import { zipSync } from "fflate";

import {
  getAttr,
  getChild,
  getChildArray,
  localName,
  parseXml,
  type XmlNode,
} from "../reader/xml.js";
import type {
  PartPath,
  PartRelationships,
  PptxSourceModel,
  PptxSourceModelEdit,
  PptxSourceModelShapeTransformEdit,
  PptxSourceModelTextRunEdit,
  RawOoxmlNode,
  RawPackagePart,
} from "../source/index.js";
import { isRelationshipPart, relationshipsPartPath } from "../source/package-paths.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";

/** `writePptx` output. */
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
 * Writes a PptxSourceModel source back to PPTX package bytes.
 *
 * This initial round-trip writer prefers unedited package material
 * to create preserved output. If the raw bytes needed to patch a dirty part are unavailable, or
 * a non-bookkeeping part lacks required raw bytes,
 * it throws instead of regenerating content implicitly.
 */
export function writePptx(source: PptxSourceModel): WritePptxOutput {
  const textRunEdits = source.edits?.filter(isTextRunEdit) ?? [];
  const shapeTransformEdits = source.edits?.filter(isShapeTransformEdit) ?? [];
  const dirtyPartPaths = new Set(
    [...textRunEdits, ...shapeTransformEdits].map((edit) => edit.handle.partPath),
  );
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
    files[partPath] = serializeDirtyXmlPart(source, partPath, textRunEdits, shapeTransformEdits);
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

function isShapeTransformEdit(
  edit: PptxSourceModelEdit,
): edit is PptxSourceModelShapeTransformEdit {
  return edit.kind === "updateShapeTransform";
}

function serializeDirtyXmlPart(
  source: PptxSourceModel,
  partPath: PartPath,
  textRunEdits: readonly PptxSourceModelTextRunEdit[],
  shapeTransformEdits: readonly PptxSourceModelShapeTransformEdit[],
): Uint8Array {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === partPath);
  if (rawPart === undefined) {
    throw new Error(`writePptx: dirty part '${partPath}' has no preserved raw package material`);
  }
  if (rawPart.kind !== "binary") {
    throw new Error(`writePptx: dirty XML tree part '${partPath}' patching is not implemented`);
  }

  const root = parseXml(textDecoder.decode(rawPart.bytes));
  for (const edit of textRunEdits.filter((edit) => edit.handle.partPath === partPath)) {
    applyTextRunEdit(root, edit);
  }
  for (const edit of shapeTransformEdits.filter((edit) => edit.handle.partPath === partPath)) {
    applyShapeTransformEdit(root, edit);
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

function applyShapeTransformEdit(root: XmlNode, edit: PptxSourceModelShapeTransformEdit): void {
  const locator = parseShapeLocator(edit.handle);
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  const shape = locateShapeTreeNode(spTree, locator);
  const xfrm = getShapeTransformNode(shape);

  if (xfrm === undefined) {
    throw new Error(
      `writePptx: shape transform handle '${String(
        edit.handle.nodeId ?? edit.handle.orderingSlot ?? "",
      )}' no longer matches source XML with xfrm`,
    );
  }

  const off = getChild(xfrm, "off");
  const ext = getChild(xfrm, "ext");
  if (off === undefined || ext === undefined) {
    throw new Error("writePptx: shape transform xfrm must contain off and ext");
  }

  off["@_x"] = String(edit.offsetX);
  off["@_y"] = String(edit.offsetY);
  ext["@_cx"] = String(edit.width);
  ext["@_cy"] = String(edit.height);
}

interface TextRunLocator {
  readonly shapeNodeId?: string;
  readonly shapeOrderingSlot?: number;
  readonly paragraphIndex: number;
  readonly runIndex: number;
}

interface ShapeLocator {
  readonly nodeId?: string;
  readonly orderingSlot?: number;
}

function parseShapeLocator(handle: PptxSourceModelShapeTransformEdit["handle"]): ShapeLocator {
  if (handle.nodeId !== undefined) return { nodeId: String(handle.nodeId) };
  if (handle.orderingSlot !== undefined) return { orderingSlot: handle.orderingSlot };
  throw new Error("writePptx: shape transform edit requires nodeId or orderingSlot in handle");
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

function locateShapeTreeNode(
  spTree: XmlNode | undefined,
  locator: ShapeLocator,
): XmlNode | undefined {
  const shapeKeys = new Set(["sp", "pic", "cxnSp", "graphicFrame", "grpSp"]);
  if (!spTree) return undefined;
  if (locator.nodeId !== undefined) {
    for (const key of Object.keys(spTree)) {
      if (key.startsWith("@_") || !shapeKeys.has(localName(key))) continue;
      const value = spTree[key];
      const items = Array.isArray(value) ? unsafeOoxmlBoundaryAssertion<unknown[]>(value) : [value];
      const found = items.find(
        (item) =>
          getShapeTreeNodeId(unsafeOoxmlBoundaryAssertion<XmlNode>(item)) === locator.nodeId,
      );
      if (found !== undefined) return unsafeOoxmlBoundaryAssertion<XmlNode>(found);
    }
    return undefined;
  }
  if (locator.orderingSlot === undefined) return undefined;
  return getShapeTreeNodeByOrderingSlot(spTree, locator.orderingSlot);
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
        return local === "sp" ? unsafeOoxmlBoundaryAssertion<XmlNode>(item) : undefined;
      }
      currentSlot++;
    }
  }
  return undefined;
}

function getShapeTreeNodeByOrderingSlot(
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
      if (currentSlot === orderingSlot) return unsafeOoxmlBoundaryAssertion<XmlNode>(item);
      currentSlot++;
    }
  }
  return undefined;
}

function getShapeTreeNodeId(node: XmlNode): string | undefined {
  const nonVisualProperties =
    getChild(node, "nvSpPr") ??
    getChild(node, "nvPicPr") ??
    getChild(node, "nvCxnSpPr") ??
    getChild(node, "nvGrpSpPr") ??
    getChild(node, "nvGraphicFramePr");
  return getAttr(getChild(nonVisualProperties, "cNvPr"), "id");
}

function getShapeTransformNode(shape: XmlNode | undefined): XmlNode | undefined {
  if (shape === undefined) return undefined;
  return (
    getChild(getChild(shape, "spPr"), "xfrm") ??
    getChild(getChild(shape, "grpSpPr"), "xfrm") ??
    getChild(shape, "xfrm")
  );
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
    const next: XmlNode = { ...unsafeOoxmlBoundaryAssertion<XmlNode>(existing), "#text": text };
    if (textRequiresPreserve(text)) next["@_xml:space"] = "preserve";
    else delete next["@_xml:space"];
    return next;
  }
  return textRequiresPreserve(text) ? { "@_xml:space": "preserve", "#text": text } : text;
}

function textRequiresPreserve(text: string): boolean {
  return text.startsWith(" ") || text.endsWith(" ");
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
