import { getChild, getChildArray, localName, parseXml, type XmlNode } from "../reader/xml.js";
import { editDirtyPartPath } from "../source/edit-descriptors.js";
import type {
  EditableTextRunProperties,
  PartPath,
  PptxSourceModel,
  PptxSourceModelAddConnectorEdit,
  PptxSourceModelAddTextBoxEdit,
  PptxSourceModelDeleteShapeEdit,
  PptxSourceModelEdit,
  PptxSourceModelParagraphTextEdit,
  PptxSourceModelShapeTransformEdit,
  PptxSourceModelTextRunEdit,
  PptxSourceModelTextRunPropertiesEdit,
} from "../source/index.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import {
  deleteShapeXml,
  getShapeTransformNode,
  locateShape,
  locateShapeTreeNode,
  type ParagraphTextLocator,
  parseParagraphLocator,
  parseShapeLocator,
  parseTextRunLocator,
} from "./xml-locators.js";
import {
  appendChild,
  cloneXmlNode,
  deleteChild,
  ensureChild,
  replaceChild,
  replaceNodeEntries,
  setChildText,
  stripXmlProcessingInstruction,
  textRequiresPreserve,
  xmlNodeIsEmpty,
} from "./xml-node-utils.js";
import { encodeXml, textDecoder, XML_DECLARATION, xmlBuilder } from "./xml-serialization.js";

export function serializeDirtyXmlPart(
  source: PptxSourceModel,
  partPath: PartPath,
  edits: readonly PptxSourceModelEdit[],
): Uint8Array {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === partPath);
  if (rawPart === undefined) {
    throw new Error(`writePptx: dirty part '${partPath}' has no preserved raw package material`);
  }
  if (rawPart.kind !== "binary") {
    throw new Error(`writePptx: dirty XML tree part '${partPath}' patching is not implemented`);
  }

  const root = parseXml(textDecoder.decode(rawPart.bytes));
  // Edits are applied in chronological order. This relies on the editing API
  // invariant that deleteShape drops earlier edits targeting the deleted shape,
  // so no stale-target edit can follow its delete within one part. Hand-built
  // edit arrays that violate the invariant fail fast in the apply functions.
  for (const edit of edits) {
    if (editDirtyPartPath(edit) !== partPath) continue;
    applyDirtyPartEdit(root, edit);
  }
  return encodeXml(XML_DECLARATION + xmlBuilder.build(stripXmlProcessingInstruction(root)));
}

/**
 * The writer-side apply switch: the one place that maps an edit kind to its XML
 * patch. Kinds that never dirty an XML part (see `editDirtyPartPath`) throw so
 * a descriptor/apply mismatch fails fast instead of being silently skipped.
 */
function applyDirtyPartEdit(root: XmlNode, edit: PptxSourceModelEdit): void {
  switch (edit.kind) {
    case "replaceTextRunPlainText":
      applyTextRunEdit(root, edit);
      return;
    case "updateTextRunProperties":
      applyTextRunPropertiesEdit(root, edit);
      return;
    case "replaceParagraphPlainText":
      applyParagraphTextEdit(root, edit);
      return;
    case "updateShapeTransform":
      applyShapeTransformEdit(root, edit);
      return;
    case "addTextBox":
      applyAddTextBoxEdit(root, edit);
      return;
    case "addConnector":
      applyAddConnectorEdit(root, edit);
      return;
    case "deleteShape":
      applyDeleteShapeEdit(root, edit);
      return;
    case "replaceImage":
    case "addEmptySlideFromLayout":
    case "duplicateSlide":
    case "moveSlide":
    case "deleteSlide":
      throw new Error(`writePptx: edit kind '${edit.kind}' does not patch a dirty XML part`);
  }
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

function applyTextRunPropertiesEdit(
  root: XmlNode,
  edit: PptxSourceModelTextRunPropertiesEdit,
): void {
  assertTextRunPropertiesEdit(edit);
  const run = locateTextRun(root, edit.handle.nodeId);

  if (run === undefined) {
    throw new Error(
      `writePptx: text run properties handle '${edit.handle.nodeId}' no longer matches source XML`,
    );
  }

  const set = edit.set ?? {};
  const hasSet = hasTextRunPropertiesSetValues(set);
  const existingRunProperties = getChild(run, "rPr");
  if (existingRunProperties === undefined && !hasSet) return;

  const rPr = existingRunProperties ?? ensureRunProperties(run);
  let cleared = false;
  for (const property of edit.clear ?? []) {
    cleared = clearRunProperty(rPr, property) || cleared;
  }
  if (set.bold !== undefined) rPr["@_b"] = booleanOoxmlValue(set.bold);
  if (set.italic !== undefined) rPr["@_i"] = booleanOoxmlValue(set.italic);
  if (set.underline !== undefined) rPr["@_u"] = set.underline ? "sng" : "none";
  if (set.fontSize !== undefined) rPr["@_sz"] = String(Math.round(set.fontSize * 100));
  if (set.typeface !== undefined) ensureChild(rPr, "latin")["@_typeface"] = set.typeface;
  if (set.color !== undefined) {
    replaceChild(rPr, "solidFill", {
      "a:srgbClr": {
        "@_val": set.color.hex.toUpperCase(),
      },
    });
  }
  if (!hasSet && cleared && xmlNodeIsEmpty(rPr)) deleteChild(run, "rPr");
}

function applyParagraphTextEdit(root: XmlNode, edit: PptxSourceModelParagraphTextEdit): void {
  const locator = parseParagraphLocator(edit.handle.nodeId);
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  const shape = locateShape(spTree, locator);
  const paragraphs = getChildArray(getChild(shape, "txBody"), "p");
  const paragraph = locatePhysicalParagraphForTextEdit(paragraphs, locator, edit.handle.nodeId);

  if (paragraph === undefined) {
    throw new Error(
      `writePptx: paragraph handle '${edit.handle.nodeId}' no longer matches source XML`,
    );
  }

  replaceParagraphRunsWithSingleTextRun(paragraph, edit.text);
}

function locateTextRun(
  root: XmlNode,
  nodeId: PptxSourceModelTextRunEdit["handle"]["nodeId"],
): XmlNode | undefined {
  const locator = parseTextRunLocator(nodeId);
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  const shape = locateShape(spTree, locator);
  const paragraphs = getChildArray(getChild(shape, "txBody"), "p");
  const paragraph = paragraphs[locator.paragraphIndex];
  return getChildArray(paragraph, "r")[locator.runIndex];
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

function applyAddTextBoxEdit(root: XmlNode, edit: PptxSourceModelAddTextBoxEdit): void {
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  if (spTree === undefined) {
    throw new Error(`writePptx: slide '${edit.slidePartPath}' has no spTree`);
  }
  if (locateShapeTreeNode(spTree, { nodeId: edit.shapeId }) !== undefined) {
    throw new Error(`writePptx: shape id '${edit.shapeId}' already exists in source XML`);
  }

  appendChild(spTree, "p:sp", parseShapeFragmentXml(edit.xml, "sp"));
}

function applyAddConnectorEdit(root: XmlNode, edit: PptxSourceModelAddConnectorEdit): void {
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  if (spTree === undefined) {
    throw new Error(`writePptx: slide '${edit.slidePartPath}' has no spTree`);
  }
  if (locateShapeTreeNode(spTree, { nodeId: edit.shapeId }) !== undefined) {
    throw new Error(`writePptx: shape id '${edit.shapeId}' already exists in source XML`);
  }
  if (
    edit.startShapeId !== undefined &&
    locateShapeTreeNode(spTree, { nodeId: edit.startShapeId }) === undefined
  ) {
    throw new Error(`writePptx: connector start shape '${edit.startShapeId}' was not found`);
  }
  if (
    edit.endShapeId !== undefined &&
    locateShapeTreeNode(spTree, { nodeId: edit.endShapeId }) === undefined
  ) {
    throw new Error(`writePptx: connector end shape '${edit.endShapeId}' was not found`);
  }

  appendChild(spTree, "p:cxnSp", parseShapeFragmentXml(edit.xml, "cxnSp"));
}

function applyDeleteShapeEdit(root: XmlNode, edit: PptxSourceModelDeleteShapeEdit): void {
  const locator = parseShapeLocator(edit.handle);
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  if (!deleteShapeXml(spTree, locator.nodeId)) {
    throw new Error(
      `writePptx: shape delete handle '${locator.nodeId}' no longer matches p:sp or p:cxnSp`,
    );
  }
}

/**
 * Parse a shape XML fragment finalized at edit time and return its single root
 * element. The writer does not generate shape XML content; it only splices the
 * pre-serialized fragment into the target `p:spTree`.
 */
function parseShapeFragmentXml(xml: string, rootLocalName: "sp" | "cxnSp"): XmlNode {
  const node = getChild(parseXml(xml), rootLocalName);
  if (node === undefined) {
    throw new Error(
      `writePptx: shape edit XML fragment does not contain a '${rootLocalName}' root element`,
    );
  }
  return node;
}

function ensureRunProperties(run: XmlNode): XmlNode {
  const existing = getChild(run, "rPr");
  if (existing !== undefined) return existing;

  const entries: [string, unknown][] = [];
  let inserted = false;
  for (const [key, value] of Object.entries(run)) {
    if (!inserted && !key.startsWith("@_")) {
      entries.push(["a:rPr", {}]);
      inserted = true;
    }
    entries.push([key, value]);
  }
  if (!inserted) entries.push(["a:rPr", {}]);
  replaceNodeEntries(run, entries);
  return getChild(run, "rPr") ?? {};
}

function clearRunProperty(
  rPr: XmlNode,
  property: NonNullable<PptxSourceModelTextRunPropertiesEdit["clear"]>[number],
): boolean {
  switch (property) {
    case "bold":
      if (rPr["@_b"] === undefined) return false;
      delete rPr["@_b"];
      return true;
    case "italic":
      if (rPr["@_i"] === undefined) return false;
      delete rPr["@_i"];
      return true;
    case "underline":
      if (rPr["@_u"] === undefined) return false;
      delete rPr["@_u"];
      return true;
    case "fontSize":
      if (rPr["@_sz"] === undefined) return false;
      delete rPr["@_sz"];
      return true;
    case "typeface": {
      const latin = getChild(rPr, "latin");
      if (latin?.["@_typeface"] === undefined) return false;
      delete latin["@_typeface"];
      if (xmlNodeIsEmpty(latin)) deleteChild(rPr, "latin");
      return true;
    }
    case "color":
      return deleteChild(rPr, "solidFill");
  }
}

function booleanOoxmlValue(value: boolean): string {
  return value ? "1" : "0";
}

function hasTextRunPropertiesSetValues(properties: EditableTextRunProperties): boolean {
  return (
    properties.bold !== undefined ||
    properties.italic !== undefined ||
    properties.underline !== undefined ||
    properties.fontSize !== undefined ||
    properties.color !== undefined ||
    properties.typeface !== undefined
  );
}

function assertTextRunPropertiesEdit(edit: PptxSourceModelTextRunPropertiesEdit): void {
  const clear = edit.clear ?? [];
  if (!hasTextRunPropertiesSetValues(edit.set ?? {}) && clear.length === 0) {
    throw new Error("writePptx: text run properties edit must set or clear at least one property");
  }
}

function replaceParagraphRunsWithSingleTextRun(paragraph: XmlNode, text: string): void {
  const firstRunProperties = getChild(getFirstRunLikeNode(paragraph), "rPr");
  const replacementRun: XmlNode = {
    ...(firstRunProperties !== undefined ? { "a:rPr": cloneXmlNode(firstRunProperties) } : {}),
    "a:t": textRequiresPreserve(text) ? { "@_xml:space": "preserve", "#text": text } : text,
  };
  const attrs: [string, unknown][] = [];
  const paragraphProperties: [string, unknown][] = [];
  const middleChildren: [string, unknown][] = [];
  const endProperties: [string, unknown][] = [];

  for (const [key, value] of Object.entries(paragraph)) {
    if (key.startsWith("@_")) {
      attrs.push([key, value]);
      continue;
    }

    const local = localName(key);
    if (isRunLikeLocalName(local)) continue;
    if (local === "pPr") paragraphProperties.push([key, value]);
    else if (local === "endParaRPr") endProperties.push([key, value]);
    else middleChildren.push([key, value]);
  }

  replaceNodeEntries(paragraph, [
    ...attrs,
    ...paragraphProperties,
    ["a:r", replacementRun],
    ...middleChildren,
    ...endProperties,
  ]);
}

function getFirstRunLikeNode(paragraph: XmlNode): XmlNode | undefined {
  for (const key of Object.keys(paragraph)) {
    if (key.startsWith("@_")) continue;
    if (!isRunLikeLocalName(localName(key))) continue;
    const value = paragraph[key];
    return Array.isArray(value)
      ? unsafeOoxmlBoundaryAssertion<XmlNode | undefined>(value[0])
      : unsafeOoxmlBoundaryAssertion<XmlNode | undefined>(value);
  }
  return undefined;
}

function isRunLikeLocalName(name: string): boolean {
  return name === "r" || name === "fld" || name === "br";
}

function locatePhysicalParagraphForTextEdit(
  paragraphs: readonly XmlNode[],
  locator: ParagraphTextLocator,
  handleNodeId: PptxSourceModelParagraphTextEdit["handle"]["nodeId"],
): XmlNode | undefined {
  let logicalParagraphIndex = 0;
  for (const paragraph of paragraphs) {
    const logicalCount = getLogicalParagraphCount(paragraph);
    if (
      locator.paragraphIndex >= logicalParagraphIndex &&
      locator.paragraphIndex < logicalParagraphIndex + logicalCount
    ) {
      if (logicalCount > 1) {
        throw new Error(
          `writePptx: paragraph handle '${handleNodeId}' references an interleaved bullet paragraph split by the reader; paragraph replacement is not supported for this source XML`,
        );
      }
      return paragraph;
    }
    logicalParagraphIndex += logicalCount;
  }
  return undefined;
}

function getLogicalParagraphCount(paragraph: XmlNode): number {
  const bulletParagraphProperties = getChildArray(paragraph, "pPr").filter(
    (properties) =>
      getChild(properties, "buChar") !== undefined ||
      getChild(properties, "buAutoNum") !== undefined,
  );
  return Math.max(1, bulletParagraphProperties.length);
}
