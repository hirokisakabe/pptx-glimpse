import type { Fill } from "./fill.js";
import type { Outline } from "./line.js";
import type { TextBody } from "./text.js";
import type { Transform } from "./shape.js";

export interface TableElement {
  type: "table";
  transform: Transform;
  table: TableData;
}

export interface TableData {
  rows: TableRow[];
  columns: TableColumn[];
}

export interface TableRow {
  height: number;
  cells: TableCell[];
}

export interface TableColumn {
  width: number;
}

export interface TableCell {
  textBody: TextBody | null;
  fill: Fill | null;
  borders: CellBorders | null;
  gridSpan: number;
  rowSpan: number;
  hMerge: boolean;
  vMerge: boolean;
}

export interface CellBorders {
  top: Outline | null;
  bottom: Outline | null;
  left: Outline | null;
  right: Outline | null;
}
