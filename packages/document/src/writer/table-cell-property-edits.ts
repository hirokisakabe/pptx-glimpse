import { getChild, getChildArray, hasChild, localName, type XmlNode } from "../reader/xml.js";
import type {
  EditableShapeFill,
  EditableTableCellBorder,
  PptxSourceModelTableCellPropertiesEdit,
} from "../source/index.js";
import { insertChildByOrder } from "./dirty-part-xml-helpers.js";
import { locateShapeTreeNodeLocation, parseShapeLocator } from "./xml-locators.js";
import { deleteChild, replaceNodeEntries, xmlNodeIsEmpty } from "./xml-node-utils.js";

const FILL_CHILD_LOCAL_NAMES: ReadonlySet<string> = new Set([
  "noFill",
  "solidFill",
  "gradFill",
  "blipFill",
  "pattFill",
  "grpFill",
]);

export function applyTableCellPropertiesEdit(
  root: XmlNode,
  edit: PptxSourceModelTableCellPropertiesEdit,
): void {
  const locator = parseShapeLocator(edit.address.tableHandle, "table cell properties edit");
  const spTree = getChild(getChild(getChild(root, "sld"), "cSld"), "spTree");
  const location = locateShapeTreeNodeLocation(spTree, locator);
  if (location?.nodeKind !== "graphicFrame") {
    throw new Error(
      `writePptx: table cell properties handle '${String(edit.address.tableHandle.nodeId)}' no longer matches a graphicFrame`,
    );
  }
  const graphicData = getChild(getChild(location.node, "graphic"), "graphicData");
  const table = getChild(graphicData, "tbl");
  const row = getChildArray(table, "tr")[edit.address.rowIndex];
  const cell = getChildArray(row, "tc")[edit.address.cellIndex];
  if (cell === undefined) {
    throw new Error(
      `writePptx: table cell address row ${edit.address.rowIndex}, cell ${edit.address.cellIndex} no longer matches source XML`,
    );
  }

  const hasSet = hasSetValues(edit.set);
  if (!hasChild(cell, "tcPr") && !hasSet) return;
  const tcPr = ensureCellProperties(cell);

  for (const property of edit.clear ?? []) clearCellProperty(tcPr, property);
  const set = edit.set;
  if (set?.marginLeft !== undefined) tcPr["@_marL"] = String(set.marginLeft);
  if (set?.marginRight !== undefined) tcPr["@_marR"] = String(set.marginRight);
  if (set?.marginTop !== undefined) tcPr["@_marT"] = String(set.marginTop);
  if (set?.marginBottom !== undefined) tcPr["@_marB"] = String(set.marginBottom);
  if (set?.fill !== undefined) replaceFillChild(tcPr, set.fill, false);
  if (set?.borders !== undefined) {
    for (const side of ["left", "right", "top", "bottom"] as const) {
      const border = set.borders[side];
      if (border !== undefined) applyBorderPatch(tcPr, side, border);
    }
  }
  if (xmlNodeIsEmpty(tcPr)) deleteChild(cell, "tcPr");
}

function ensureCellProperties(cell: XmlNode): XmlNode {
  for (const [key, value] of Object.entries(cell)) {
    if (key.startsWith("@_") || localName(key) !== "tcPr") continue;
    if (isXmlNode(value)) return value;
    if (Array.isArray(value) || value !== "") {
      throw new Error("writePptx: table cell has unsupported duplicate or malformed tcPr XML");
    }
    const replacement: XmlNode = {};
    cell[key] = replacement;
    return replacement;
  }
  const properties: XmlNode = {};
  insertChildByOrder(cell, "a:tcPr", properties, (name) => name === "extLst");
  return properties;
}

function isXmlNode(value: unknown): value is XmlNode {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function clearCellProperty(
  tcPr: XmlNode,
  property: NonNullable<PptxSourceModelTableCellPropertiesEdit["clear"]>[number],
): void {
  switch (property) {
    case "fill":
      removeFillChildren(tcPr);
      return;
    case "borderTop":
      deleteChild(tcPr, "lnT");
      return;
    case "borderBottom":
      deleteChild(tcPr, "lnB");
      return;
    case "borderLeft":
      deleteChild(tcPr, "lnL");
      return;
    case "borderRight":
      deleteChild(tcPr, "lnR");
      return;
    case "marginLeft":
      delete tcPr["@_marL"];
      return;
    case "marginRight":
      delete tcPr["@_marR"];
      return;
    case "marginTop":
      delete tcPr["@_marT"];
      return;
    case "marginBottom":
      delete tcPr["@_marB"];
      return;
  }
}

function applyBorderPatch(
  tcPr: XmlNode,
  side: "left" | "right" | "top" | "bottom",
  patch: EditableTableCellBorder,
): void {
  const local = { left: "lnL", right: "lnR", top: "lnT", bottom: "lnB" }[side];
  let line = getChild(tcPr, local);
  if (line === undefined) {
    insertChildByOrder(tcPr, `a:${local}`, {}, (name) =>
      [
        "lnTlToBr",
        "lnBlToTr",
        "cell3D",
        "noFill",
        "solidFill",
        "gradFill",
        "blipFill",
        "pattFill",
        "grpFill",
        "extLst",
      ].includes(name),
    );
    line = getChild(tcPr, local) ?? {};
  }
  if (patch.width !== undefined) line["@_w"] = String(patch.width);
  if (patch.fill !== undefined) replaceFillChild(line, patch.fill, true);
}

function replaceFillChild(parent: XmlNode, fill: EditableShapeFill, lineFill: boolean): void {
  removeFillChildren(parent);
  const key = fill.kind === "none" ? "a:noFill" : "a:solidFill";
  const value =
    fill.kind === "none" ? {} : { "a:srgbClr": { "@_val": fill.color.hex.toUpperCase() } };
  insertChildByOrder(parent, key, value, (name) =>
    lineFill
      ? [
          "prstDash",
          "custDash",
          "round",
          "bevel",
          "miter",
          "headEnd",
          "tailEnd",
          "extLst",
        ].includes(name)
      : name === "extLst",
  );
}

function removeFillChildren(parent: XmlNode): void {
  replaceNodeEntries(
    parent,
    Object.entries(parent).filter(
      ([key]) => key.startsWith("@_") || !FILL_CHILD_LOCAL_NAMES.has(localName(key)),
    ),
  );
}

function hasSetValues(set: PptxSourceModelTableCellPropertiesEdit["set"]): boolean {
  return (
    set !== undefined &&
    (set.fill !== undefined ||
      set.borders !== undefined ||
      set.marginLeft !== undefined ||
      set.marginRight !== undefined ||
      set.marginTop !== undefined ||
      set.marginBottom !== undefined)
  );
}
