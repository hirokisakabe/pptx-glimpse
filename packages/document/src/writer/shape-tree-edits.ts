import { getAttr, getChild, localName, type XmlNode } from "../reader/xml.js";
import type {
  PptxSourceModelAddChartEdit,
  PptxSourceModelAddConnectorEdit,
  PptxSourceModelAddPictureEdit,
  PptxSourceModelAddShapeEdit,
  PptxSourceModelAddTableEdit,
  PptxSourceModelAddTextBoxEdit,
  PptxSourceModelDeleteShapeEdit,
  PptxSourceModelReorderShapesEdit,
  PptxSourceModelSetSlideBackgroundEdit,
} from "../source/index.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import {
  appendShapeTreeNodeAtEnd,
  elementPrefix,
  ensurePictureNamespaces,
  getDrawingPartRoot,
  parseShapeFragmentXml,
  preserveNamespaceDeclarations,
  qualifiedSiblingName,
  remapElementPrefix,
} from "./dirty-part-xml-helpers.js";
import { deleteShapeXml, locateShapeTreeNode, parseShapeLocator } from "./xml-locators.js";
import { replaceNodeEntries } from "./xml-node-utils.js";
import { parseXmlForEditing } from "./xml-serialization.js";
import { setXmlChildOrder } from "./xml-serialization.js";

export function applyAddTextBoxEdit(root: XmlNode, edit: PptxSourceModelAddTextBoxEdit): void {
  applyAddSpEdit(root, edit);
}

export function applyAddShapeEdit(root: XmlNode, edit: PptxSourceModelAddShapeEdit): void {
  applyAddSpEdit(root, edit);
}

export function applyAddTableEdit(root: XmlNode, edit: PptxSourceModelAddTableEdit): void {
  const slide = getChild(root, "sld");
  if (slide !== undefined) ensurePictureNamespaces(slide);
  const spTree = getChild(getChild(slide, "cSld"), "spTree");
  if (spTree === undefined)
    throw new Error(`writePptx: slide '${edit.slidePartPath}' has no spTree`);
  assertShapeIdAvailable(spTree, edit.shapeId);
  appendShapeTreeNodeAtEnd(
    spTree,
    "p:graphicFrame",
    parseShapeFragmentXml(edit.xml, "graphicFrame"),
  );
}

export function applyAddConnectorEdit(root: XmlNode, edit: PptxSourceModelAddConnectorEdit): void {
  const spTree = getShapeTree(root, edit.slidePartPath);
  assertShapeIdAvailable(spTree, edit.shapeId);
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
  appendShapeTreeNodeAtEnd(spTree, "p:cxnSp", parseShapeFragmentXml(edit.xml, "cxnSp"));
}

export function applyAddPictureEdit(root: XmlNode, edit: PptxSourceModelAddPictureEdit): void {
  const drawingPart = getDrawingPartRoot(root);
  if (drawingPart !== undefined) ensurePictureNamespaces(drawingPart);
  const spTree = getShapeTree(root, edit.slidePartPath);
  assertShapeIdAvailable(spTree, edit.shapeId);
  appendShapeTreeNodeAtEnd(spTree, "p:pic", parseShapeFragmentXml(edit.xml, "pic"));
}

export function applyAddChartEdit(root: XmlNode, edit: PptxSourceModelAddChartEdit): void {
  const slide = getChild(root, "sld");
  if (slide !== undefined) ensurePictureNamespaces(slide);
  const spTree = getChild(getChild(slide, "cSld"), "spTree");
  if (spTree === undefined)
    throw new Error(`writePptx: slide '${edit.slidePartPath}' has no spTree`);
  assertShapeIdAvailable(spTree, edit.shapeId);
  appendShapeTreeNodeAtEnd(
    spTree,
    "p:graphicFrame",
    parseShapeFragmentXml(edit.xml, "graphicFrame"),
  );
}

export function applyDeleteShapeEdit(root: XmlNode, edit: PptxSourceModelDeleteShapeEdit): void {
  const locator = parseShapeLocator(edit.handle, "shape delete edit");
  const spTree = getChild(getChild(getChild(root, "sld"), "cSld"), "spTree");
  if (!deleteShapeXml(spTree, locator.nodeId)) {
    throw new Error(
      `writePptx: shape delete handle '${locator.nodeId}' no longer matches p:sp or p:cxnSp`,
    );
  }
}

export function applyReorderShapesEdit(
  root: XmlNode,
  edit: PptxSourceModelReorderShapesEdit,
): void {
  const spTree = getShapeTree(root, edit.targetPartPath);
  const current: { key: string; value: unknown }[] = [];
  const shapeById = new Map<string, { key: string; value: unknown }>();
  for (const [key, grouped] of Object.entries(spTree)) {
    if (key.startsWith("@_")) continue;
    const values = Array.isArray(grouped)
      ? unsafeOoxmlBoundaryAssertion<unknown[]>(grouped)
      : [grouped];
    for (const value of values) {
      const entry = { key, value };
      current.push(entry);
      if (localName(key) === "nvGrpSpPr" || localName(key) === "grpSpPr") continue;
      const nodeId = shapeTreeEntryNodeId(value);
      if (nodeId !== undefined) shapeById.set(nodeId, entry);
    }
  }
  if (shapeById.size !== edit.shapeIds.length) {
    throw new Error("writePptx: reordered shape ids must contain every shape exactly once");
  }
  const orderedShapes = edit.shapeIds.map((shapeId) => {
    const entry = shapeById.get(shapeId);
    if (entry === undefined) {
      throw new Error(`writePptx: reordered shape '${shapeId}' was not found`);
    }
    return entry;
  });
  let orderedShapeIndex = 0;
  setXmlChildOrder(
    spTree,
    current.map((entry) =>
      shapeTreeEntryNodeId(entry.value) === undefined ? entry : orderedShapes[orderedShapeIndex++],
    ),
  );
}

export function applySetSlideBackgroundEdit(
  root: XmlNode,
  edit: PptxSourceModelSetSlideBackgroundEdit,
): void {
  const slide = getChild(root, "sld");
  const cSldKey =
    slide === undefined
      ? undefined
      : Object.keys(slide).find((key) => !key.startsWith("@_") && localName(key) === "cSld");
  const cSld = getChild(slide, "cSld");
  if (slide === undefined || cSldKey === undefined || cSld === undefined) {
    throw new Error(`writePptx: slide '${edit.slidePartPath}' has no p:cSld`);
  }
  const spTreeKey = Object.keys(cSld).find(
    (key) => !key.startsWith("@_") && localName(key) === "spTree",
  );
  if (spTreeKey === undefined) {
    throw new Error(`writePptx: slide '${edit.slidePartPath}' has no p:spTree`);
  }
  slide["@_xmlns:a"] ??= "http://schemas.openxmlformats.org/drawingml/2006/main";
  if (edit.relationshipId !== undefined) {
    slide["@_xmlns:r"] ??= "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  }
  const parsedBackground = getChild(parseXmlForEditing(edit.xml), "bg");
  if (parsedBackground === undefined) {
    throw new Error("writePptx: background edit XML fragment does not contain a p:bg root element");
  }
  const existingBackgroundKey = Object.keys(cSld).find(
    (key) => !key.startsWith("@_") && localName(key) === "bg",
  );
  const existingBackground = getChild(cSld, "bg");
  const backgroundKey = existingBackgroundKey ?? qualifiedSiblingName(cSldKey, "bg");
  const remappedBackground = remapElementPrefix(
    parsedBackground,
    "p",
    elementPrefix(backgroundKey),
  );
  const background = preserveNamespaceDeclarations(existingBackground, remappedBackground);

  const entries: [string, unknown][] = [];
  let inserted = false;
  for (const [key, value] of Object.entries(cSld)) {
    if (!key.startsWith("@_") && localName(key) === "bg") continue;
    if (!inserted && key === spTreeKey) {
      entries.push([backgroundKey, background]);
      inserted = true;
    }
    entries.push([key, value]);
  }
  replaceNodeEntries(cSld, entries);
}

function applyAddSpEdit(
  root: XmlNode,
  edit: PptxSourceModelAddTextBoxEdit | PptxSourceModelAddShapeEdit,
): void {
  const spTree = getShapeTree(root, edit.slidePartPath);
  assertShapeIdAvailable(spTree, edit.shapeId);
  appendShapeTreeNodeAtEnd(spTree, "p:sp", parseShapeFragmentXml(edit.xml, "sp"));
}

function getShapeTree(root: XmlNode, partPath: string): XmlNode {
  const spTree = getChild(getChild(getDrawingPartRoot(root), "cSld"), "spTree");
  if (spTree === undefined) throw new Error(`writePptx: drawing part '${partPath}' has no spTree`);
  return spTree;
}

function assertShapeIdAvailable(spTree: XmlNode, shapeId: string): void {
  if (locateShapeTreeNode(spTree, { nodeId: shapeId }) !== undefined) {
    throw new Error(`writePptx: shape id '${shapeId}' already exists in source XML`);
  }
}

function shapeTreeNodeId(node: XmlNode): string | undefined {
  const nonVisualProperties =
    getChild(node, "nvSpPr") ??
    getChild(node, "nvPicPr") ??
    getChild(node, "nvCxnSpPr") ??
    getChild(node, "nvGrpSpPr") ??
    getChild(node, "nvGraphicFramePr");
  return getAttr(getChild(nonVisualProperties, "cNvPr"), "id");
}

function shapeTreeEntryNodeId(value: unknown): string | undefined {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return undefined;
  return shapeTreeNodeId(unsafeOoxmlBoundaryAssertion<XmlNode>(value));
}
