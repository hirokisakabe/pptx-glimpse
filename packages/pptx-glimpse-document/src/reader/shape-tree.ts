/**
 * `p:spTree` を CleanDoc source の shape node 列へ読み取る。
 *
 * simple autoshape (`p:sp`), embedded raster image (`p:pic`), connector
 * (`p:cxnSp`), and group (`p:grpSp`) を typed に表す。graphicFrame は table /
 * chart / SmartArt の supported subset を typed 化し、それ以外の未対応ノードは
 * raw shape node として保存する。typed node 内でも、未対応の子要素・属性は raw
 * sidecar として保持する。
 *
 * `orderedChildren` が渡された場合は preserve-order XML parse 結果を使い、異種
 * タグ間の z-order を維持する。未指定時は従来通りタグ種別ごとの順序に fallback
 * する。
 */

import type {
  PartPath,
  RawSidecarId,
  SourceCellBorders,
  SourceChart,
  SourceConnector,
  SourceGeometry,
  SourceGroup,
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
  SourceTransform,
} from "../source/index.js";
import { asEmu, asOoxmlPercent, asRelationshipId, asSourceNodeId } from "../source/index.js";
import { parseCustomGeometry } from "./custom-geometry.js";
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
  type XmlOrderedNode,
} from "./xml.js";

const KNOWN_SHAPE_CHILDREN: ReadonlySet<string> = new Set(["nvSpPr", "spPr", "txBody"]);
const KNOWN_CONNECTOR_CHILDREN: ReadonlySet<string> = new Set(["nvCxnSpPr", "spPr", "style"]);
const KNOWN_GROUP_CHILDREN: ReadonlySet<string> = new Set(["nvGrpSpPr", "grpSpPr"]);
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

const SHAPE_TREE_NODE_TAGS: ReadonlySet<string> = new Set([
  "sp",
  "pic",
  "cxnSp",
  "grpSp",
  "graphicFrame",
]);

const CHART_GRAPHIC_DATA_URIS: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/chart",
  "http://purl.oclc.org/ooxml/drawingml/chart",
]);
const SMARTART_DIAGRAM_URIS: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/diagram",
  "http://purl.oclc.org/ooxml/drawingml/diagram",
]);

/** `p:spTree` を読み、shape node 列を返す。 */
export function parseShapeTree(
  spTree: XmlNode | undefined,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderedChildren?: readonly XmlOrderedNode[],
): SourceShapeNode[] {
  if (!spTree) return [];
  if (orderedChildren !== undefined) {
    return parseShapeTreeOrdered(spTree, partPath, nextId, orderedChildren);
  }

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
      } else if (local === "cxnSp") {
        nodes.push(parseConnector(node, partPath, nextId, orderingSlot));
      } else if (local === "grpSp") {
        nodes.push(parseGroup(node, partPath, nextId, orderingSlot));
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

function parseShapeTreeOrdered(
  spTree: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderedChildren: readonly XmlOrderedNode[],
): SourceShapeNode[] {
  const nodes: SourceShapeNode[] = [];
  const tagCounters: Record<string, number> = {};
  let orderingSlot = 0;

  for (const child of orderedChildren) {
    const key = Object.keys(child).find((candidate) => candidate !== ":@");
    if (key === undefined) continue;
    const local = localName(key);

    if (local === "AlternateContent") {
      const index = tagCounters[local] ?? 0;
      tagCounters[local] = index + 1;
      const alternateContents = getChildArray(spTree, "AlternateContent");
      const alternateContent = alternateContents[index];
      if (alternateContent !== undefined) {
        nodes.push(...parseAlternateContent(alternateContent, partPath, nextId, orderingSlot));
        orderingSlot++;
      }
      continue;
    }

    if (!SHAPE_TREE_NODE_TAGS.has(local)) continue;
    const index = tagCounters[local] ?? 0;
    tagCounters[local] = index + 1;
    const node = getChildArray(spTree, local)[index];
    if (node === undefined) continue;

    nodes.push(parseShapeTreeNode(local, node, child, partPath, nextId, orderingSlot));
    orderingSlot++;
  }

  return nodes;
}

function parseShapeTreeNode(
  local: string,
  node: XmlNode,
  orderedNode: XmlOrderedNode | undefined,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
): SourceShapeNode {
  if (local === "sp") return parseShape(node, partPath, nextId, orderingSlot, orderedNode);
  if (local === "pic") return parseImage(node, partPath, nextId, orderingSlot);
  if (local === "cxnSp") {
    return parseConnector(node, partPath, nextId, orderingSlot, orderedNode);
  }
  if (local === "grpSp") {
    const orderedGroupChildren = orderedNode?.[local] as readonly XmlOrderedNode[] | undefined;
    return parseGroup(node, partPath, nextId, orderingSlot, orderedGroupChildren);
  }
  if (local === "graphicFrame") return parseGraphicFrame(node, partPath, nextId, orderingSlot);
  return parseRawShapeNode(`p:${local}`, node, partPath, nextId, orderingSlot);
}

function parseShape(
  sp: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
  orderedNode?: XmlOrderedNode,
): SourceShape {
  const nvSpPr = getChild(sp, "nvSpPr");
  const cNvPr = getChild(nvSpPr, "cNvPr");
  const nodeId = sourceNodeId(cNvPr);
  const name = getAttr(cNvPr, "name");
  const placeholder = parsePlaceholder(getChild(getChild(nvSpPr, "nvPr"), "ph"));

  const spPr = getChild(sp, "spPr");
  const transform = parseTransform(spPr);
  const geometry = parseGeometry(spPr, orderedNestedChildChildren(orderedNode, "sp", "spPr"));
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

function parseConnector(
  cxnSp: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
  orderedNode?: XmlOrderedNode,
): SourceConnector {
  const nvCxnSpPr = getChild(cxnSp, "nvCxnSpPr");
  const cNvPr = getChild(nvCxnSpPr, "cNvPr");
  const nodeId = sourceNodeId(cNvPr);
  const name = getAttr(cNvPr, "name");
  const spPr = getChild(cxnSp, "spPr");
  const transform = parseTransform(spPr);
  const geometry = parseGeometry(spPr, orderedNestedChildChildren(orderedNode, "cxnSp", "spPr"));
  const outline = parseOutline(spPr, nextId);
  const rawSidecars = [
    ...collectUnknownSidecars(cxnSp, KNOWN_CONNECTOR_CHILDREN, nextId),
    ...collectUnknownSidecars(spPr, KNOWN_SP_PR_CHILDREN, nextId),
  ];

  return {
    kind: "connector",
    ...(nodeId !== undefined ? { nodeId } : {}),
    ...(name !== undefined ? { name } : {}),
    ...(transform !== undefined ? { transform } : {}),
    ...(geometry !== undefined ? { geometry } : {}),
    ...(outline !== undefined ? { outline } : {}),
    handle: { partPath, ...(nodeId !== undefined ? { nodeId } : {}), orderingSlot },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseGroup(
  grpSp: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  orderingSlot: number,
  orderedChildren?: readonly XmlOrderedNode[],
): SourceGroup {
  const nvGrpSpPr = getChild(grpSp, "nvGrpSpPr");
  const cNvPr = getChild(nvGrpSpPr, "cNvPr");
  const nodeId = sourceNodeId(cNvPr);
  const name = getAttr(cNvPr, "name");
  const grpSpPr = getChild(grpSp, "grpSpPr");
  const transform = parseTransform(grpSpPr);
  const childTransform = parseChildTransform(grpSpPr, transform);
  const rawSidecars = collectUnknownSidecars(grpSp, KNOWN_GROUP_CHILDREN, nextId);

  return {
    kind: "group",
    ...(nodeId !== undefined ? { nodeId } : {}),
    ...(name !== undefined ? { name } : {}),
    ...(transform !== undefined ? { transform } : {}),
    ...(childTransform !== undefined ? { childTransform } : {}),
    children: parseShapeTree(grpSp, partPath, nextId, orderedChildren),
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
  if (uri !== undefined && CHART_GRAPHIC_DATA_URIS.has(uri)) {
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
    if (parsed.length > 0) {
      if (parsed.every((node) => node.kind === "raw")) {
        continue;
      }
      return parsed.map((node) =>
        attachRawSidecar(node, makeSidecar("mc:AlternateContent", alternateContent, nextId)),
      );
    }
  }
  return [
    parseRawShapeNode("mc:AlternateContent", alternateContent, partPath, nextId, orderingSlot),
  ];
}

function attachRawSidecar<T extends SourceShapeNode>(
  node: T,
  sidecar: ReturnType<typeof makeSidecar>,
): T {
  if (node.kind === "raw") return node;
  return {
    ...node,
    rawSidecars: [...(node.rawSidecars ?? []), sidecar],
  };
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
      else if (local === "cxnSp") {
        nodes.push(parseConnector(node, partPath, nextId, orderingSlot));
      } else if (local === "grpSp") {
        nodes.push(parseGroup(node, partPath, nextId, orderingSlot));
      } else if (local === "graphicFrame") {
        nodes.push(parseGraphicFrame(node, partPath, nextId, orderingSlot));
      } else {
        nodes.push(parseRawShapeNode(key, node, partPath, nextId, orderingSlot));
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
    ...collectUnknownSidecars(graphicData, EMPTY_KNOWN, nextId),
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
    ...collectUnknownSidecars(graphicData, EMPTY_KNOWN, nextId),
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

function parseGeometry(
  spPr: XmlNode | undefined,
  orderedSpPr?: readonly XmlOrderedNode[],
): SourceGeometry | undefined {
  const prstGeom = getChild(spPr, "prstGeom");
  if (prstGeom) {
    const preset = getAttr(prstGeom, "prst");
    const adjustValues = parseAdjustValues(getChild(prstGeom, "avLst"));
    return preset !== undefined
      ? {
          preset,
          ...(Object.keys(adjustValues).length > 0 ? { adjustValues } : {}),
        }
      : undefined;
  }

  const customPaths = parseCustomGeometry(
    getChild(spPr, "custGeom"),
    orderedChildChildren(orderedSpPr, "custGeom"),
  );
  if (customPaths !== undefined) return { kind: "custom", paths: customPaths };
  return undefined;
}

function parseAdjustValues(avLst: XmlNode | undefined): Record<string, number> {
  const adjustValues: Record<string, number> = {};
  for (const guide of getChildArray(avLst, "gd")) {
    const name = getAttr(guide, "name");
    const formula = getAttr(guide, "fmla");
    const match = formula?.match(/val\s+(-?\d+)/);
    if (name !== undefined && match !== undefined && match !== null) {
      adjustValues[name] = Number(match[1]);
    }
  }
  return adjustValues;
}

function parseChildTransform(
  grpSpPr: XmlNode | undefined,
  fallback: SourceTransform | undefined,
): SourceTransform | undefined {
  const xfrm = getChild(grpSpPr, "xfrm");
  const childOff = getChild(xfrm, "chOff");
  const childExt = getChild(xfrm, "chExt");
  const offsetX = numericAttr(childOff, "x") ?? 0;
  const offsetY = numericAttr(childOff, "y") ?? 0;
  const width = numericAttr(childExt, "cx") ?? fallback?.width;
  const height = numericAttr(childExt, "cy") ?? fallback?.height;
  if (width === undefined || height === undefined) return undefined;
  return {
    offsetX: asEmu(offsetX),
    offsetY: asEmu(offsetY),
    width: asEmu(Number(width)),
    height: asEmu(Number(height)),
  };
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

function orderedChildChildren(
  parent: readonly XmlOrderedNode[] | undefined,
  childLocalName: string,
): readonly XmlOrderedNode[] | undefined {
  const child = parent?.find((entry) => {
    const key = Object.keys(entry).find((candidate) => candidate !== ":@");
    return key !== undefined && localName(key) === childLocalName;
  });
  if (child === undefined) return undefined;
  const key = Object.keys(child).find((candidate) => candidate !== ":@");
  const value = key !== undefined ? child[key] : undefined;
  return Array.isArray(value) ? (value as readonly XmlOrderedNode[]) : undefined;
}

function orderedNestedChildChildren(
  node: XmlOrderedNode | undefined,
  parentLocalName: string,
  childLocalName: string,
): readonly XmlOrderedNode[] | undefined {
  if (node === undefined) return undefined;
  const parentChildren = node[parentLocalName];
  return orderedChildChildren(
    Array.isArray(parentChildren) ? (parentChildren as readonly XmlOrderedNode[]) : undefined,
    childLocalName,
  );
}
