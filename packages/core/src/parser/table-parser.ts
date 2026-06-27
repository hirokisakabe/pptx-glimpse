import type { Outline } from "@pptx-glimpse/renderer";
import type {
  CellBorders,
  TableCell,
  TableColumn,
  TableData,
  TableRow,
} from "@pptx-glimpse/renderer";
import type { FontScheme } from "@pptx-glimpse/renderer";
import { asEmu } from "@pptx-glimpse/renderer";

import type { ColorResolver } from "../color/color-resolver.js";
import { unsafeXmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import { parseFillFromNode, parseOutline } from "./fill-parser.js";
import { parseTextBody } from "./slide-parser.js";
import type { XmlNode } from "./xml-parser.js";

export function parseTable(
  tblNode: XmlNode,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): TableData | null {
  if (!tblNode) return null;

  const columns = parseColumns(unsafeXmlBoundaryAssertion<XmlNode>(tblNode.tblGrid));
  if (columns.length === 0) return null;

  const tblPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(tblNode.tblPr);
  const hasTableStyle = tblPr?.tableStyleId !== undefined;
  const defaultBorders = hasTableStyle ? createDefaultBorders() : null;

  const rows = parseRows(
    unsafeXmlBoundaryAssertion<XmlNode>(tblNode.tr),
    colorResolver,
    fontScheme,
    defaultBorders,
  );

  return { rows, columns };
}

function parseColumns(tblGrid: XmlNode): TableColumn[] {
  if (!tblGrid) return [];

  const gridCols = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(tblGrid.gridCol) ?? [];
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
    const cells = parseCells(
      unsafeXmlBoundaryAssertion<XmlNode>(tr.tc),
      colorResolver,
      fontScheme,
      defaultBorders,
    );
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
    const textBody = parseTextBody(
      unsafeXmlBoundaryAssertion<XmlNode>(tc.txBody),
      colorResolver,
      undefined,
      fontScheme,
    );
    const tcPr = unsafeXmlBoundaryAssertion<XmlNode | undefined>(tc.tcPr);
    const fill = tcPr ? parseFillFromNode(tcPr, colorResolver) : null;
    const inlineBorders = tcPr ? parseCellBorders(tcPr, colorResolver) : null;
    const borders = inlineBorders ?? defaultBorders ?? null;
    const gridSpan = Number(tc["@_gridSpan"] ?? 1);
    const rowSpan = Number(tc["@_rowSpan"] ?? 1);
    const hMerge =
      unsafeXmlBoundaryAssertion<string | undefined>(tcPr?.["@_hMerge"]) === "1" ||
      unsafeXmlBoundaryAssertion<string | undefined>(tcPr?.["@_hMerge"]) === "true";
    const vMerge =
      unsafeXmlBoundaryAssertion<string | undefined>(tcPr?.["@_vMerge"]) === "1" ||
      unsafeXmlBoundaryAssertion<string | undefined>(tcPr?.["@_vMerge"]) === "true";

    cells.push({ textBody, fill, borders, gridSpan, rowSpan, hMerge, vMerge });
  }
  return cells;
}

function parseCellBorders(tcPr: XmlNode, colorResolver: ColorResolver): CellBorders | null {
  const top = parseOutline(unsafeXmlBoundaryAssertion<XmlNode>(tcPr.lnT), colorResolver);
  const bottom = parseOutline(unsafeXmlBoundaryAssertion<XmlNode>(tcPr.lnB), colorResolver);
  const left = parseOutline(unsafeXmlBoundaryAssertion<XmlNode>(tcPr.lnL), colorResolver);
  const right = parseOutline(unsafeXmlBoundaryAssertion<XmlNode>(tcPr.lnR), colorResolver);

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
