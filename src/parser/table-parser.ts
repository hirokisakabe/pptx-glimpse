import type { TableData, TableRow, TableColumn, TableCell, CellBorders } from "../model/table.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { FontScheme } from "../model/theme.js";
import { parseFillFromNode, parseOutline } from "./fill-parser.js";
import { parseTextBody } from "./slide-parser.js";

export function parseTable(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  tblNode: any,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): TableData | null {
  if (!tblNode) return null;

  const columns = parseColumns(tblNode.tblGrid);
  if (columns.length === 0) return null;

  const rows = parseRows(tblNode.tr, colorResolver, fontScheme);

  return { rows, columns };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseColumns(tblGrid: any): TableColumn[] {
  if (!tblGrid) return [];

  const gridCols = tblGrid.gridCol ?? [];
  return gridCols.map((col: Record<string, string>) => ({
    width: Number(col["@_w"] ?? 0),
  }));
}

function parseRows(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  trList: any,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): TableRow[] {
  if (!trList) return [];

  const rows: TableRow[] = [];
  for (const tr of trList) {
    const height = Number(tr["@_h"] ?? 0);
    const cells = parseCells(tr.tc, colorResolver, fontScheme);
    rows.push({ height, cells });
  }
  return rows;
}

function parseCells(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  tcList: any,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): TableCell[] {
  if (!tcList) return [];

  const cells: TableCell[] = [];
  for (const tc of tcList) {
    const textBody = parseTextBody(tc.txBody, colorResolver, undefined, fontScheme);
    const tcPr = tc.tcPr;
    const fill = tcPr ? parseFillFromNode(tcPr, colorResolver) : null;
    const borders = tcPr ? parseCellBorders(tcPr, colorResolver) : null;
    const gridSpan = Number(tc["@_gridSpan"] ?? 1);
    const rowSpan = Number(tc["@_rowSpan"] ?? 1);
    const hMerge = tcPr?.["@_hMerge"] === "1" || tcPr?.["@_hMerge"] === "true";
    const vMerge = tcPr?.["@_vMerge"] === "1" || tcPr?.["@_vMerge"] === "true";

    cells.push({ textBody, fill, borders, gridSpan, rowSpan, hMerge, vMerge });
  }
  return cells;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseCellBorders(tcPr: any, colorResolver: ColorResolver): CellBorders | null {
  const top = parseOutline(tcPr.lnT, colorResolver);
  const bottom = parseOutline(tcPr.lnB, colorResolver);
  const left = parseOutline(tcPr.lnL, colorResolver);
  const right = parseOutline(tcPr.lnR, colorResolver);

  if (!top && !bottom && !left && !right) return null;

  return { top, bottom, left, right };
}
