/**
 * `p:spTree` を CleanDoc source の shape node 列へ読み取る。
 *
 * simple autoshape (`p:sp`) と embedded raster image (`p:pic`) を typed に表し、
 * group / connector / graphicFrame 等の未対応ノードは raw shape node として
 * 保存する (`docs/cleandoc-minimal-poc-scope.md`)。typed node 内でも、未対応の
 * 子要素・属性は raw sidecar として保持する。
 *
 * 注意: fast-xml-parser は同名要素のみを配列化するため、子の反復順は
 * 「タグ種別の初出順 × 同名内の文書順」となり、異種タグ間の完全な文書順
 * (z-order) は保証されない。effective element ordering の解決は computed view
 * (#465) の責務であり、PoC 対象 fixture は sp/pic が交互配置されないため
 * 実害はない。
 */

import type {
  PartPath,
  RawSidecarId,
  SourceCellBorders,
  SourceChart,
  SourceImage,
  SourceImageCrop,
  SourceNodeId,
  SourcePlaceholder,
  SourceRawShapeNode,
  SourceShape,
  SourceShapeNode,
  SourceSmartArt,
  SourceTable,
  SourceTableCell,
  SourceTableColumn,
  SourceTableRow,
} from "../source/index.js";
import { asEmu, asOoxmlPercent, asRelationshipId, asSourceNodeId } from "../source/index.js";
import {
  isTrue,
  numericAttr,
  parseFill,
  parseLine,
  parseOutline,
  parseTransform,
} from "./drawing.js";
import { collectUnknownSidecars, makeSidecar } from "./raw-node.js";
import { parseTextBody } from "./text.js";
import {
  getAttr,
  getChild,
  getChildArray,
  getChildText,
  getNamespacedAttr,
  localName,
  type XmlNode,
} from "./xml.js";

const KNOWN_SHAPE_CHILDREN: ReadonlySet<string> = new Set(["nvSpPr", "spPr", "txBody"]);
const KNOWN_PICTURE_CHILDREN: ReadonlySet<string> = new Set(["nvPicPr", "blipFill", "spPr"]);
const KNOWN_BLIP_FILL_CHILDREN: ReadonlySet<string> = new Set(["blip", "srcRect"]);
const KNOWN_GRAPHIC_FRAME_CHILDREN: ReadonlySet<string> = new Set([
  "nvGraphicFramePr",
  "xfrm",
  "graphic",
]);
const KNOWN_GRAPHIC_CHILDREN: ReadonlySet<string> = new Set(["graphicData"]);
const KNOWN_GRAPHIC_DATA_CHILDREN: ReadonlySet<string> = new Set(["tbl"]);
const KNOWN_TABLE_CHILDREN: ReadonlySet<string> = new Set(["tblPr", "tblGrid", "tr"]);
const KNOWN_TABLE_CELL_CHILDREN: ReadonlySet<string> = new Set(["txBody", "tcPr"]);
const KNOWN_TABLE_CELL_PROPERTIES_CHILDREN: ReadonlySet<string> = new Set([
  "lnL",
  "lnR",
  "lnT",
  "lnB",
  "solidFill",
  "noFill",
  "gradFill",
  "blipFill",
  "pattFill",
  "grpFill",
]);
// `a:spPr` のうち typed に解釈する子。これ以外 (custGeom / effectLst / scene3d /
// extLst 等) は raw sidecar として保持する。fill 系は parseFill が typed/raw を
// 判別するため known 扱いにして二重計上を防ぐ。
const KNOWN_SP_PR_CHILDREN: ReadonlySet<string> = new Set([
  "xfrm",
  "prstGeom",
  "solidFill",
  "noFill",
  "gradFill",
  "blipFill",
  "pattFill",
  "grpFill",
  "ln",
]);

const CHART_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const SMARTART_DIAGRAM_URIS = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/diagram",
  "http://purl.oclc.org/ooxml/drawingml/diagram",
]);

/** `p:spTree` を読み、shape node 列を返す。 */
export function parseShapeTree(
  spTree: XmlNode | undefined,
  partPath: PartPath,
  nextId: () => RawSidecarId,
): SourceShapeNode[] {
  if (!spTree) return [];

  const nodes: SourceShapeNode[] = [];
  let orderingSlot = 0;
  for (const key of Object.keys(spTree)) {
    if (key.startsWith("@_")) continue;
    if (key === "#text") continue;
    const local = localName(key);
    // group 自身の非可視プロパティはノードではないため除外する。
    if (local === "nvGrpSpPr" || local === "grpSpPr") continue;

    const value = spTree[key];
    const items = Array.isArray(value) ? value : [value];
    for (const item of items) {
      const node = item as XmlNode;
      if (local === "sp") {
        nodes.push(parseShape(node, partPath, nextId, orderingSlot));
      } else if (local === "pic") {
        nodes.push(parseImage(node, partPath, nextId, orderingSlot));
      } else if (local === "graphicFrame") {
        nodes.push(parseGraphicFrame(node, partPath, nextId, orderingSlot));
      } else if (local === "AlternateContent") {
        const parsed = parseAlternateContent(node, partPath, nextId, orderingSlot);
        nodes.push(...parsed);
      } else {
        nodes.push(parseRawShapeNode(key, node, partPath, nextId, orderingSlot));
      }
      orderingSlot++;
    }
  }
  return nodes;
}

function parseShape(
  sp: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceShape {
  const nvSpPr = getChild(sp, "nvSpPr");
  const cNvPr = getChild(nvSpPr, "cNvPr");
  const nodeId = sourceNodeId(cNvPr);
  const name = getAttr(cNvPr, "name");
  const placeholder = parsePlaceholder(getChild(getChild(nvSpPr, "nvPr"), "ph"));

  const spPr = getChild(sp, "spPr");
  const transform = parseTransform(spPr);
  const geometry = parsePresetGeometry(spPr);
  const fill = parseFill(spPr, nextId);
  const outline = parseOutline(spPr, nextId);
  const textBody = parseTextBody(getChild(sp, "txBody"), partPath, nextId, nodeId, orderingSlot);

  const rawSidecars = [
    ...collectUnknownSidecars(sp, KNOWN_SHAPE_CHILDREN, nextId),
    ...collectUnknownSidecars(spPr, KNOWN_SP_PR_CHILDREN, nextId),
  ];

  return {
    kind: "shape",
    ...(nodeId !== undefined ? { nodeId } : {}),
    ...(name !== undefined ? { name } : {}),
    ...(transform !== undefined ? { transform } : {}),
    ...(geometry !== undefined ? { geometry } : {}),
    ...(fill !== undefined ? { fill } : {}),
    ...(outline !== undefined ? { outline } : {}),
    ...(textBody !== undefined ? { textBody } : {}),
    ...(placeholder !== undefined ? { placeholder } : {}),
    handle: { partPath, ...(nodeId !== undefined ? { nodeId } : {}), orderingSlot },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseImage(
  pic: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceImage {
  const cNvPr = getChild(getChild(pic, "nvPicPr"), "cNvPr");
  const nodeId = sourceNodeId(cNvPr);
  const name = getAttr(cNvPr, "name");

  const blipFill = getChild(pic, "blipFill");
  const blip = getChild(blipFill, "blip");
  const embed = getNamespacedAttr(blip, "embed");
  const crop = parseCrop(getChild(blipFill, "srcRect"));

  const spPr = getChild(pic, "spPr");
  const transform = parseTransform(spPr);

  const rawSidecars = [
    ...collectUnknownSidecars(pic, KNOWN_PICTURE_CHILDREN, nextId),
    ...collectUnknownSidecars(spPr, KNOWN_SP_PR_CHILDREN, nextId),
    // `a:stretch` / `a:tile` 等の fill mode と blip 配下の recolor 操作を保持する。
    ...collectUnknownSidecars(blipFill, KNOWN_BLIP_FILL_CHILDREN, nextId),
    ...collectUnknownSidecars(blip, EMPTY_KNOWN, nextId),
  ];

  return {
    kind: "image",
    ...(nodeId !== undefined ? { nodeId } : {}),
    ...(name !== undefined ? { name } : {}),
    ...(transform !== undefined ? { transform } : {}),
    ...(embed !== undefined ? { blipRelationshipId: asRelationshipId(embed) } : {}),
    ...(crop !== undefined ? { crop } : {}),
    handle: {
      partPath,
      ...(nodeId !== undefined ? { nodeId } : {}),
      ...(embed !== undefined ? { relationshipId: asRelationshipId(embed) } : {}),
      orderingSlot,
    },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseGraphicFrame(
  graphicFrame: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceShapeNode {
  const graphic = getChild(graphicFrame, "graphic");
  const graphicData = getChild(graphic, "graphicData");
  const uri = getAttr(graphicData, "uri");
  if (uri === CHART_GRAPHIC_DATA_URI) {
    const parsedChart = parseChartGraphicFrame(graphicFrame, partPath, nextId, orderingSlot);
    if (parsedChart !== undefined) return parsedChart;
  }
  if (uri !== undefined && SMARTART_DIAGRAM_URIS.has(uri)) {
    const parsedSmartArt = parseSmartArtGraphicFrame(graphicFrame, partPath, nextId, orderingSlot);
    if (parsedSmartArt !== undefined) return parsedSmartArt;
  }

  const table = getChild(graphicData, "tbl");
  if (table === undefined) {
    return parseRawShapeNode("p:graphicFrame", graphicFrame, partPath, nextId, orderingSlot);
  }

  const parsedTable = parseTable(table, graphicFrame, partPath, nextId, orderingSlot);
  if (parsedTable === undefined) {
    return parseRawShapeNode("p:graphicFrame", graphicFrame, partPath, nextId, orderingSlot);
  }
  return parsedTable;
}

function parseAlternateContent(
  alternateContent: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceShapeNode[] {
  const branches = [
    ...getChildArray(alternateContent, "Choice"),
    ...getChildArray(alternateContent, "Fallback"),
  ];
  for (const branch of branches) {
    const parsed = parseAlternateContentBranch(branch, partPath, nextId, orderingSlot);
    if (parsed.length > 0) return parsed;
  }
  return [
    parseRawShapeNode("mc:AlternateContent", alternateContent, partPath, nextId, orderingSlot),
  ];
}

function parseAlternateContentBranch(
  branch: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceShapeNode[] {
  const nodes: SourceShapeNode[] = [];
  for (const key of Object.keys(branch)) {
    if (key.startsWith("@_") || key === "#text") continue;
    const local = localName(key);
    const value = branch[key];
    const items = Array.isArray(value) ? value : [value];
    for (const item of items) {
      const node = item as XmlNode;
      if (local === "sp") nodes.push(parseShape(node, partPath, nextId, orderingSlot));
      else if (local === "pic") nodes.push(parseImage(node, partPath, nextId, orderingSlot));
      else if (local === "graphicFrame") {
        nodes.push(parseGraphicFrame(node, partPath, nextId, orderingSlot));
      }
    }
  }
  return nodes;
}

function parseChartGraphicFrame(
  graphicFrame: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceChart | undefined {
  const graphic = getChild(graphicFrame, "graphic");
  const graphicData = getChild(graphic, "graphicData");
  const chart = getChild(graphicData, "chart");
  const rId = getNamespacedAttr(chart, "id") ?? getAttr(chart, "id");
  if (rId === undefined) return undefined;

  const nvGraphicFramePr = getChild(graphicFrame, "nvGraphicFramePr");
  const cNvPr = getChild(nvGraphicFramePr, "cNvPr");
  const nodeId = sourceNodeId(cNvPr);
  const name = getAttr(cNvPr, "name");
  const transform = parseTransform(graphicFrame);
  const rawSidecars = [
    ...collectUnknownSidecars(graphicFrame, KNOWN_GRAPHIC_FRAME_CHILDREN, nextId),
    ...collectUnknownSidecars(graphic, KNOWN_GRAPHIC_CHILDREN, nextId),
    ...collectUnknownSidecars(graphicData, new Set(["chart"]), nextId),
  ];

  return {
    kind: "chart",
    ...(nodeId !== undefined ? { nodeId } : {}),
    ...(name !== undefined ? { name } : {}),
    ...(transform !== undefined ? { transform } : {}),
    chartRelationshipId: asRelationshipId(rId),
    handle: {
      partPath,
      ...(nodeId !== undefined ? { nodeId } : {}),
      relationshipId: asRelationshipId(rId),
      orderingSlot,
    },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseSmartArtGraphicFrame(
  graphicFrame: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceSmartArt | undefined {
  const graphic = getChild(graphicFrame, "graphic");
  const graphicData = getChild(graphic, "graphicData");
  const relIds = getChild(graphicData, "relIds");
  const dmRId = getNamespacedAttr(relIds, "dm") ?? getAttr(relIds, "dm");
  if (dmRId === undefined) return undefined;

  const nvGraphicFramePr = getChild(graphicFrame, "nvGraphicFramePr");
  const cNvPr = getChild(nvGraphicFramePr, "cNvPr");
  const nodeId = sourceNodeId(cNvPr);
  const name = getAttr(cNvPr, "name");
  const transform = parseTransform(graphicFrame);
  const rawSidecars = [
    ...collectUnknownSidecars(graphicFrame, KNOWN_GRAPHIC_FRAME_CHILDREN, nextId),
    ...collectUnknownSidecars(graphic, KNOWN_GRAPHIC_CHILDREN, nextId),
    ...collectUnknownSidecars(graphicData, new Set(["relIds"]), nextId),
  ];

  return {
    kind: "smartArt",
    ...(nodeId !== undefined ? { nodeId } : {}),
    ...(name !== undefined ? { name } : {}),
    ...(transform !== undefined ? { transform } : {}),
    dataRelationshipId: asRelationshipId(dmRId),
    handle: {
      partPath,
      ...(nodeId !== undefined ? { nodeId } : {}),
      relationshipId: asRelationshipId(dmRId),
      orderingSlot,
    },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseTable(
  tbl: XmlNode,
  graphicFrame: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceTable | undefined {
  const columns = parseTableColumns(getChild(tbl, "tblGrid"));
  if (columns.length === 0) return undefined;

  const nvGraphicFramePr = getChild(graphicFrame, "nvGraphicFramePr");
  const cNvPr = getChild(nvGraphicFramePr, "cNvPr");
  const nodeId = sourceNodeId(cNvPr);
  const name = getAttr(cNvPr, "name");
  const tblPr = getChild(tbl, "tblPr");
  const tableStyleId = getChildText(tblPr, "tableStyleId");
  const graphic = getChild(graphicFrame, "graphic");
  const graphicData = getChild(graphic, "graphicData");
  const transform = parseTransform(graphicFrame);

  const rows = parseTableRows(tbl, partPath, nextId, nodeId, orderingSlot);
  const rawSidecars = [
    ...collectUnknownSidecars(graphicFrame, KNOWN_GRAPHIC_FRAME_CHILDREN, nextId),
    ...collectUnknownSidecars(graphic, KNOWN_GRAPHIC_CHILDREN, nextId),
    ...collectUnknownSidecars(graphicData, KNOWN_GRAPHIC_DATA_CHILDREN, nextId),
    ...collectUnknownSidecars(tbl, KNOWN_TABLE_CHILDREN, nextId),
  ];

  return {
    kind: "table",
    ...(nodeId !== undefined ? { nodeId } : {}),
    ...(name !== undefined ? { name } : {}),
    ...(transform !== undefined ? { transform } : {}),
    table: {
      columns,
      rows,
      ...(tableStyleId !== undefined ? { tableStyleId } : {}),
    },
    handle: { partPath, ...(nodeId !== undefined ? { nodeId } : {}), orderingSlot },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseTableColumns(tblGrid: XmlNode | undefined): SourceTableColumn[] {
  return getChildArray(tblGrid, "gridCol").map((col) => ({ width: emuAttr(col, "w") }));
}

function parseTableRows(
  tbl: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  tableNodeId: SourceNodeId | undefined,
  tableOrderingSlot: number,
): SourceTableRow[] {
  return getChildArray(tbl, "tr").map((tr, rowIndex) => ({
    height: emuAttr(tr, "h"),
    cells: getChildArray(tr, "tc").map((tc, cellIndex) =>
      parseTableCell(tc, partPath, nextId, tableNodeId, tableOrderingSlot, rowIndex, cellIndex),
    ),
  }));
}

function parseTableCell(
  tc: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  tableNodeId: SourceNodeId | undefined,
  tableOrderingSlot: number,
  rowIndex: number,
  cellIndex: number,
): SourceTableCell {
  const tcPr = getChild(tc, "tcPr");
  const textOrderingSlot = tableOrderingSlot * 1_000_000_000 + rowIndex * 1_000_000 + cellIndex;
  const textBody = parseTextBody(
    getChild(tc, "txBody"),
    partPath,
    nextId,
    tableNodeId,
    textOrderingSlot,
  );
  const fill = parseFill(tcPr, nextId);
  const borders = parseCellBorders(tcPr, nextId);
  const rawSidecars = [
    ...collectUnknownSidecars(tc, KNOWN_TABLE_CELL_CHILDREN, nextId),
    ...collectUnknownSidecars(tcPr, KNOWN_TABLE_CELL_PROPERTIES_CHILDREN, nextId),
  ];

  return {
    ...(textBody !== undefined ? { textBody } : {}),
    ...(fill !== undefined ? { fill } : {}),
    ...(borders !== undefined ? { borders } : {}),
    gridSpan: numericAttr(tc, "gridSpan") ?? 1,
    rowSpan: numericAttr(tc, "rowSpan") ?? 1,
    hMerge: isTrue(getAttr(tc, "hMerge") ?? getAttr(tcPr, "hMerge")),
    vMerge: isTrue(getAttr(tc, "vMerge") ?? getAttr(tcPr, "vMerge")),
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseCellBorders(
  tcPr: XmlNode | undefined,
  nextId: () => RawSidecarId,
): SourceCellBorders | undefined {
  const left = parseLine(getChild(tcPr, "lnL"), nextId);
  const right = parseLine(getChild(tcPr, "lnR"), nextId);
  const top = parseLine(getChild(tcPr, "lnT"), nextId);
  const bottom = parseLine(getChild(tcPr, "lnB"), nextId);
  if (left === undefined && right === undefined && top === undefined && bottom === undefined) {
    return undefined;
  }
  return {
    ...(top !== undefined ? { top } : {}),
    ...(bottom !== undefined ? { bottom } : {}),
    ...(left !== undefined ? { left } : {}),
    ...(right !== undefined ? { right } : {}),
  };
}

function parseRawShapeNode(
  key: string,
  node: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceRawShapeNode {
  const nodeId = sourceNodeId(getChild(getChild(node, "nvSpPr"), "cNvPr"));
  return {
    kind: "raw",
    ...(nodeId !== undefined ? { nodeId } : {}),
    raw: makeSidecar(key, node, nextId),
    handle: { partPath, ...(nodeId !== undefined ? { nodeId } : {}), orderingSlot },
  };
}

function parsePlaceholder(ph: XmlNode | undefined): SourcePlaceholder | undefined {
  if (!ph) return undefined;
  const type = getAttr(ph, "type");
  const index = numericAttr(ph, "idx");
  if (type === undefined && index === undefined) return undefined;
  return {
    ...(type !== undefined ? { type } : {}),
    ...(index !== undefined ? { index } : {}),
  };
}

function parsePresetGeometry(spPr: XmlNode | undefined) {
  const prstGeom = getChild(spPr, "prstGeom");
  if (!prstGeom) return undefined;
  const preset = getAttr(prstGeom, "prst");
  return preset !== undefined ? { preset } : undefined;
}

function parseCrop(srcRect: XmlNode | undefined): SourceImageCrop | undefined {
  if (!srcRect) return undefined;
  const left = numericAttr(srcRect, "l");
  const top = numericAttr(srcRect, "t");
  const right = numericAttr(srcRect, "r");
  const bottom = numericAttr(srcRect, "b");
  const crop: SourceImageCrop = {
    ...(left !== undefined ? { left: asOoxmlPercent(left) } : {}),
    ...(top !== undefined ? { top: asOoxmlPercent(top) } : {}),
    ...(right !== undefined ? { right: asOoxmlPercent(right) } : {}),
    ...(bottom !== undefined ? { bottom: asOoxmlPercent(bottom) } : {}),
  };
  return Object.keys(crop).length > 0 ? crop : undefined;
}

function sourceNodeId(cNvPr: XmlNode | undefined): SourceNodeId | undefined {
  const id = getAttr(cNvPr, "id");
  return id !== undefined ? asSourceNodeId(id) : undefined;
}

function emuAttr(node: XmlNode | undefined, attrName: string) {
  return asEmu(numericAttr(node, attrName) ?? 0);
}

const EMPTY_KNOWN: ReadonlySet<string> = new Set();
