import type { TableData, TableRow, TableColumn, TableCell, CellBorders } from "../model/table.js";
import type { Outline } from "../model/line.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { FontScheme } from "../model/theme.js";
import { parseFillFromNode, parseOutline } from "./fill-parser.js";
import { parseTextBody } from "./slide-parser.js";
import type { XmlNode } from "./xml-parser.js";
import { asEmu } from "../utils/unit-types.js";

export function parseTable(
  tblNode: XmlNode,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): TableData | null {
  if (!tblNode) return null;

  const columns = parseColumns(tblNode.tblGrid as XmlNode);
  if (columns.length === 0) return null;

  const tblPr = tblNode.tblPr as XmlNode | undefined;
  const hasTableStyle = tblPr?.tableStyleId !== undefined;
  const defaultBorders = hasTableStyle ? createDefaultBorders() : null;

  const rows = parseRows(tblNode.tr as XmlNode, colorResolver, fontScheme, defaultBorders);

  return { rows, columns };
}

function parseColumns(tblGrid: XmlNode): TableColumn[] {
  if (!tblGrid) return [];

  const gridCols = (tblGrid.gridCol as XmlNode[] | undefined) ?? [];
  return gridCols.map((col) => ({
    width: asEmu(Number(col["@_w"] ?? 0)),
  }));
}

function parseRows(
  trList: XmlNode,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
  defaultBorders?: CellBorders | null,
): TableRow[] {
  if (!trList) return [];

  const trArr = Array.isArray(trList) ? (trList as XmlNode[]) : [trList];
  const rows: TableRow[] = [];
  for (const tr of trArr) {
    const height = asEmu(Number(tr["@_h"] ?? 0));
    const cells = parseCells(tr.tc as XmlNode, colorResolver, fontScheme, defaultBorders);
    rows.push({ height, cells });
  }
  return rows;
}

function parseCells(
  tcList: XmlNode,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
  defaultBorders?: CellBorders | null,
): TableCell[] {
  if (!tcList) return [];

  const tcArr = Array.isArray(tcList) ? (tcList as XmlNode[]) : [tcList];
  const cells: TableCell[] = [];
  for (const tc of tcArr) {
    const textBody = parseTextBody(tc.txBody as XmlNode, colorResolver, undefined, fontScheme);
    const tcPr = tc.tcPr as XmlNode | undefined;
    const fill = tcPr ? parseFillFromNode(tcPr, colorResolver) : null;
    const inlineBorders = tcPr ? parseCellBorders(tcPr, colorResolver) : null;
    const borders = inlineBorders ?? defaultBorders ?? null;
    const gridSpan = Number(tc["@_gridSpan"] ?? 1);
    const rowSpan = Number(tc["@_rowSpan"] ?? 1);
    const hMerge =
      (tcPr?.["@_hMerge"] as string | undefined) === "1" ||
      (tcPr?.["@_hMerge"] as string | undefined) === "true";
    const vMerge =
      (tcPr?.["@_vMerge"] as string | undefined) === "1" ||
      (tcPr?.["@_vMerge"] as string | undefined) === "true";

    cells.push({ textBody, fill, borders, gridSpan, rowSpan, hMerge, vMerge });
  }
  return cells;
}

function parseCellBorders(tcPr: XmlNode, colorResolver: ColorResolver): CellBorders | null {
  const top = parseOutline(tcPr.lnT as XmlNode, colorResolver);
  const bottom = parseOutline(tcPr.lnB as XmlNode, colorResolver);
  const left = parseOutline(tcPr.lnL as XmlNode, colorResolver);
  const right = parseOutline(tcPr.lnR as XmlNode, colorResolver);

  for (const border of [top, bottom, left, right]) {
    if (border && !border.fill) {
      border.fill = { type: "solid", color: { hex: "#000000", alpha: 1 } };
    }
  }

  if (!top && !bottom && !left && !right) return null;

  return { top, bottom, left, right };
}

function createDefaultBorders(): CellBorders {
  const defaultOutline: Outline = {
    width: asEmu(12700),
    fill: { type: "solid", color: { hex: "#000000", alpha: 1 } },
    dashStyle: "solid",
    headEnd: null,
    tailEnd: null,
  };
  return {
    top: { ...defaultOutline },
    bottom: { ...defaultOutline },
    left: { ...defaultOutline },
    right: { ...defaultOutline },
  };
}
