import { getChild, localName, type XmlNode } from "../reader/xml.js";
import type {
  EditableShapeFill,
  EditableShapeOutline,
  PptxSourceModelShapeFillEdit,
  PptxSourceModelShapeOutlineEdit,
  PptxSourceModelShapeTransformEdit,
  SourceHandle,
} from "../source/index.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import { insertChildByOrder } from "./dirty-part-xml-helpers.js";
import { getShapeTransformNode, locateShapeTreeNode, parseShapeLocator } from "./xml-locators.js";
import { replaceNodeEntries } from "./xml-node-utils.js";

const FILL_CHILD_LOCAL_NAMES: ReadonlySet<string> = new Set([
  "noFill",
  "solidFill",
  "gradFill",
  "blipFill",
  "pattFill",
  "grpFill",
]);

export function applyShapeTransformEdit(
  root: XmlNode,
  edit: PptxSourceModelShapeTransformEdit,
): void {
  const locator = parseShapeLocator(edit.handle, "shape transform edit");
  const spTree = getChild(getChild(getChild(root, "sld"), "cSld"), "spTree");
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

export function applyShapeFillEdit(root: XmlNode, edit: PptxSourceModelShapeFillEdit): void {
  const shape = locateEditableShapeTreeNode(root, edit.handle, "shape fill edit");
  if (shape.localName === "cxnSp") {
    throw new Error(
      `writePptx: shape fill handle '${String(edit.handle.nodeId)}' references a connector`,
    );
  }
  replaceFillChild(ensureShapeProperties(shape.node), edit.fill);
}

export function applyShapeOutlineEdit(root: XmlNode, edit: PptxSourceModelShapeOutlineEdit): void {
  const shape = locateEditableShapeTreeNode(root, edit.handle, "shape outline edit");
  const ln = ensureLineProperties(ensureShapeProperties(shape.node));
  applyOutlinePatch(ln, edit.outline);
}

interface ShapeTreeNodeLocation {
  readonly node: XmlNode;
  readonly localName: string;
}

function locateEditableShapeTreeNode(
  root: XmlNode,
  handle: SourceHandle,
  editName: string,
): ShapeTreeNodeLocation {
  const locator = parseShapeLocator(handle, editName);
  const spTree = getChild(getChild(getChild(root, "sld"), "cSld"), "spTree");
  const shape = locateShapeTreeNodeWithLocalName(spTree, locator.nodeId);
  if (shape === undefined) {
    throw new Error(
      `writePptx: ${editName} handle '${String(handle.nodeId)}' no longer matches source XML`,
    );
  }
  if (shape.localName !== "sp" && shape.localName !== "cxnSp") {
    throw new Error(
      `writePptx: ${editName} handle '${String(handle.nodeId)}' does not reference p:sp or p:cxnSp`,
    );
  }
  return shape;
}

function locateShapeTreeNodeWithLocalName(
  spTree: XmlNode | undefined,
  nodeId: string,
): ShapeTreeNodeLocation | undefined {
  if (spTree === undefined) return undefined;
  for (const key of Object.keys(spTree)) {
    if (key.startsWith("@_")) continue;
    const keyLocalName = localName(key);
    const value = spTree[key];
    const items = Array.isArray(value) ? value : [value];
    for (const item of items) {
      const node = unsafeOoxmlBoundaryAssertion<XmlNode>(item);
      const nonVisualProperties =
        getChild(node, "nvSpPr") ??
        getChild(node, "nvPicPr") ??
        getChild(node, "nvCxnSpPr") ??
        getChild(node, "nvGrpSpPr") ??
        getChild(node, "nvGraphicFramePr");
      if (getChild(nonVisualProperties, "cNvPr")?.["@_id"] === nodeId) {
        return { node, localName: keyLocalName };
      }
    }
  }
  return undefined;
}

function ensureShapeProperties(shape: XmlNode): XmlNode {
  const existing = getChild(shape, "spPr");
  if (existing !== undefined) return existing;
  const entries: [string, unknown][] = [];
  let inserted = false;
  for (const [key, value] of Object.entries(shape)) {
    entries.push([key, value]);
    if (!inserted && !key.startsWith("@_") && localName(key).startsWith("nv")) {
      entries.push(["p:spPr", {}]);
      inserted = true;
    }
  }
  if (!inserted) entries.push(["p:spPr", {}]);
  replaceNodeEntries(shape, entries);
  return getChild(shape, "spPr") ?? {};
}

function ensureLineProperties(spPr: XmlNode): XmlNode {
  const existing = getChild(spPr, "ln");
  if (existing !== undefined) return existing;
  insertChildByOrder(spPr, "a:ln", {}, (name) =>
    ["effectLst", "effectDag", "scene3d", "sp3d", "extLst"].includes(name),
  );
  return getChild(spPr, "ln") ?? {};
}

function applyOutlinePatch(ln: XmlNode, outline: EditableShapeOutline): void {
  if (outline.width !== undefined) ln["@_w"] = String(outline.width);
  if (outline.fill !== undefined) replaceFillChild(ln, outline.fill);
}

function replaceFillChild(parent: XmlNode, fill: EditableShapeFill): void {
  const fillNode =
    fill.kind === "none"
      ? { key: "a:noFill", value: {} }
      : {
          key: "a:solidFill",
          value: { "a:srgbClr": { "@_val": fill.color.hex.toUpperCase() } },
        };
  const entries = Object.entries(parent).filter(
    ([key]) => key.startsWith("@_") || !FILL_CHILD_LOCAL_NAMES.has(localName(key)),
  );
  replaceNodeEntries(parent, entries);
  insertChildByOrder(parent, fillNode.key, fillNode.value, (name) =>
    [
      "ln",
      "effectLst",
      "effectDag",
      "scene3d",
      "sp3d",
      "extLst",
      "prstDash",
      "custDash",
      "round",
      "bevel",
      "miter",
      "headEnd",
      "tailEnd",
    ].includes(name),
  );
}
