/** Existing native Table cell property editing for PptxSourceModel. */

import { sourceHandlesEqual } from "./edit-descriptors.js";
import type {
  EditableShapeFill,
  EditableTableCellBorder,
  EditableTableCellProperties,
  EditableTableCellProperty,
  PptxSourceModel,
  SourceCellBorders,
  SourceFill,
  SourceGroup,
  SourceOutline,
  SourceShapeNode,
  SourceSlide,
  SourceTable,
  SourceTableCell,
  TableCellAddress,
} from "./index.js";

const EDITABLE_PROPERTIES: ReadonlySet<string> = new Set([
  "fill",
  "borderTop",
  "borderBottom",
  "borderLeft",
  "borderRight",
  "marginLeft",
  "marginRight",
  "marginTop",
  "marginBottom",
]);
const SET_PROPERTY_KEYS: ReadonlySet<string> = new Set([
  "fill",
  "borders",
  "marginLeft",
  "marginRight",
  "marginTop",
  "marginBottom",
]);
const BORDER_SIDE_KEYS: ReadonlySet<string> = new Set(["top", "bottom", "left", "right"]);
const BORDER_PROPERTY_KEYS: ReadonlySet<string> = new Set(["width", "fill"]);
const NONE_FILL_KEYS: ReadonlySet<string> = new Set(["kind"]);
const SOLID_FILL_KEYS: ReadonlySet<string> = new Set(["kind", "color"]);
const SRGB_COLOR_KEYS: ReadonlySet<string> = new Set(["kind", "hex"]);

type MutableTableCell = { -readonly [K in keyof SourceTableCell]: SourceTableCell[K] };
type MutableCellBorders = { -readonly [K in keyof SourceCellBorders]: SourceCellBorders[K] };

export function setTableCellProperties(
  source: PptxSourceModel,
  address: TableCellAddress,
  properties: EditableTableCellProperties,
): PptxSourceModel {
  assertAddress(address, "setTableCellProperties");
  assertProperties(properties, "setTableCellProperties");
  return updateTableCellProperties(source, address, properties, []);
}

export function clearTableCellProperties(
  source: PptxSourceModel,
  address: TableCellAddress,
  properties: readonly EditableTableCellProperty[],
): PptxSourceModel {
  assertAddress(address, "clearTableCellProperties");
  if (!Array.isArray(properties) || properties.length === 0) {
    throw new Error("clearTableCellProperties: properties must contain at least one property name");
  }
  for (const property of properties) {
    if (typeof property !== "string" || !EDITABLE_PROPERTIES.has(property)) {
      throw new Error(`clearTableCellProperties: unsupported table cell property '${property}'`);
    }
  }
  return updateTableCellProperties(source, address, {}, [...new Set(properties)]);
}

function updateTableCellProperties(
  source: PptxSourceModel,
  address: TableCellAddress,
  set: EditableTableCellProperties,
  clear: readonly EditableTableCellProperty[],
): PptxSourceModel {
  const target = findTarget(source, address);
  if (target === undefined) {
    throw new Error(
      "updateTableCellProperties: table cell address was not found in PptxSourceModel source",
    );
  }
  const nextCell = patchCell(target.cell, set, clear);
  if (stableValueEqual(target.cell, nextCell)) return source;

  const slides = source.slides.map((slide) =>
    slide === target.slide
      ? {
          ...slide,
          shapes: replaceTableCell(slide.shapes, address, nextCell),
        }
      : slide,
  );
  return {
    ...source,
    slides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "updateTableCellProperties",
        address,
        ...(hasSetValues(set) ? { set } : {}),
        ...(clear.length > 0 ? { clear } : {}),
      },
    ],
  };
}

function findTarget(
  source: PptxSourceModel,
  address: TableCellAddress,
):
  | { readonly slide: SourceSlide; readonly table: SourceTable; readonly cell: SourceTableCell }
  | undefined {
  const matches = source.slides.flatMap((slide) =>
    findTables(slide.shapes).flatMap((match) =>
      sourceHandlesEqual(match.table.handle, address.tableHandle) ? [{ slide, ...match }] : [],
    ),
  );
  if (matches.length > 1) {
    throw new Error(
      "updateTableCellProperties: table handle is ambiguous in PptxSourceModel source",
    );
  }
  const match = matches[0];
  if (match?.insideAlternateContent) {
    throw new Error("updateTableCellProperties: tables inside AlternateContent are not supported");
  }
  const cell = match?.table.table.rows[address.rowIndex]?.cells[address.cellIndex];
  return match === undefined || cell === undefined ? undefined : { ...match, cell };
}

function replaceTableCell(
  nodes: readonly SourceShapeNode[],
  address: TableCellAddress,
  cell: SourceTableCell,
): readonly SourceShapeNode[] {
  return nodes.map((node) => {
    if (node.kind === "group") {
      if (!containsTable(node.children, address)) return node;
      const children = replaceTableCell(node.children, address, cell);
      return { ...node, children } satisfies SourceGroup;
    }
    if (node.kind !== "table" || !sourceHandlesEqual(node.handle, address.tableHandle)) return node;
    return {
      ...node,
      table: {
        ...node.table,
        rows: node.table.rows.map((row, rowIndex) =>
          rowIndex === address.rowIndex
            ? {
                ...row,
                cells: row.cells.map((current, cellIndex) =>
                  cellIndex === address.cellIndex ? cell : current,
                ),
              }
            : row,
        ),
      },
    } satisfies SourceTable;
  });
}

function containsTable(nodes: readonly SourceShapeNode[], address: TableCellAddress): boolean {
  return findTables(nodes).some((match) =>
    sourceHandlesEqual(match.table.handle, address.tableHandle),
  );
}

interface TableMatch {
  readonly table: SourceTable;
  readonly insideAlternateContent: boolean;
}

function findTables(
  nodes: readonly SourceShapeNode[],
  insideAlternateContent = false,
): TableMatch[] {
  return nodes.flatMap((node): TableMatch[] => {
    const nodeInsideAlternateContent =
      insideAlternateContent || hasAlternateContentWrapperSidecar(node);
    if (node.kind === "table") {
      return [{ table: node, insideAlternateContent: nodeInsideAlternateContent }];
    }
    return node.kind === "group" ? findTables(node.children, nodeInsideAlternateContent) : [];
  });
}

function hasAlternateContentWrapperSidecar(node: SourceShapeNode): boolean {
  if (node.kind === "raw") return false;
  return (
    node.rawSidecars?.some(
      (sidecar) =>
        sidecar.node.name === "mc:AlternateContent" && sidecar.orderingSlot === undefined,
    ) ?? false
  );
}

function patchCell(
  current: SourceTableCell,
  set: EditableTableCellProperties,
  clear: readonly EditableTableCellProperty[],
): SourceTableCell {
  const next: MutableTableCell = { ...current };
  for (const property of clear) clearProperty(next, property);
  if (set.fill !== undefined) next.fill = toSourceFill(set.fill);
  if (set.marginLeft !== undefined) next.marginLeft = set.marginLeft;
  if (set.marginRight !== undefined) next.marginRight = set.marginRight;
  if (set.marginTop !== undefined) next.marginTop = set.marginTop;
  if (set.marginBottom !== undefined) next.marginBottom = set.marginBottom;
  if (set.borders !== undefined) {
    const borders: MutableCellBorders = { ...next.borders };
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const patch = set.borders[side];
      if (patch !== undefined) borders[side] = patchBorder(borders[side], patch);
    }
    next.borders = borders;
  }
  return next;
}

function clearProperty(cell: MutableTableCell, property: EditableTableCellProperty): void {
  switch (property) {
    case "fill":
      delete cell.fill;
      return;
    case "marginLeft":
      delete cell.marginLeft;
      return;
    case "marginRight":
      delete cell.marginRight;
      return;
    case "marginTop":
      delete cell.marginTop;
      return;
    case "marginBottom":
      delete cell.marginBottom;
      return;
    case "borderTop":
      clearBorder(cell, "top");
      return;
    case "borderBottom":
      clearBorder(cell, "bottom");
      return;
    case "borderLeft":
      clearBorder(cell, "left");
      return;
    case "borderRight":
      clearBorder(cell, "right");
      return;
  }
}

function clearBorder(cell: MutableTableCell, side: keyof SourceCellBorders): void {
  const borders: MutableCellBorders = { ...cell.borders };
  delete borders[side];
  if (Object.keys(borders).length === 0) delete cell.borders;
  else cell.borders = borders;
}

function patchBorder(
  current: SourceOutline | undefined,
  patch: EditableTableCellBorder,
): SourceOutline {
  return {
    ...(current ?? {}),
    ...(patch.width !== undefined ? { width: patch.width } : {}),
    ...(patch.fill !== undefined ? { fill: toSourceFill(patch.fill) } : {}),
  };
}

function toSourceFill(fill: EditableShapeFill): SourceFill {
  return fill.kind === "none"
    ? { kind: "none" }
    : { kind: "solid", color: { kind: "srgb", hex: fill.color.hex } };
}

function assertAddress(address: TableCellAddress, operation: string): void {
  if (address.tableHandle.nodeId === undefined) {
    throw new Error(`${operation}: table cell edit requires a table node id`);
  }
  if (!Number.isInteger(address.rowIndex) || address.rowIndex < 0) {
    throw new Error(`${operation}: rowIndex must be a non-negative integer`);
  }
  if (!Number.isInteger(address.cellIndex) || address.cellIndex < 0) {
    throw new Error(`${operation}: cellIndex must be a non-negative integer`);
  }
}

function assertProperties(properties: EditableTableCellProperties, operation: string): void {
  if (typeof properties !== "object" || properties === null || !hasSetValues(properties)) {
    throw new Error(`${operation}: properties must set at least one table cell property`);
  }
  for (const property of Object.keys(properties)) {
    if (!SET_PROPERTY_KEYS.has(property)) {
      throw new Error(`${operation}: unsupported table cell property '${property}'`);
    }
  }
  if (properties.fill !== undefined) assertFill(properties.fill, operation, "fill");
  for (const margin of ["marginLeft", "marginRight", "marginTop", "marginBottom"] as const) {
    const value = properties[margin];
    if (value !== undefined && (!Number.isInteger(value) || value < 0)) {
      throw new Error(`${operation}: ${margin} must be a non-negative integer EMU value`);
    }
  }
  if (properties.borders !== undefined) {
    if (typeof properties.borders !== "object" || properties.borders === null) {
      throw new Error(`${operation}: borders must be an object`);
    }
    for (const side of Object.keys(properties.borders)) {
      if (!BORDER_SIDE_KEYS.has(side)) {
        throw new Error(`${operation}: unsupported table cell border side '${side}'`);
      }
    }
    let hasBorder = false;
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const border = properties.borders[side];
      if (border === undefined) continue;
      hasBorder = true;
      assertBorder(border, operation, `borders.${side}`);
    }
    if (!hasBorder) throw new Error(`${operation}: borders must set at least one side`);
  }
}

function assertBorder(border: EditableTableCellBorder, operation: string, path: string): void {
  if (
    typeof border !== "object" ||
    border === null ||
    (border.width === undefined && border.fill === undefined)
  ) {
    throw new Error(`${operation}: ${path} must set width or fill`);
  }
  for (const property of Object.keys(border)) {
    if (!BORDER_PROPERTY_KEYS.has(property)) {
      throw new Error(`${operation}: ${path}.${property} is not a supported border style`);
    }
  }
  if (border.width !== undefined && (!Number.isInteger(border.width) || border.width <= 0)) {
    throw new Error(`${operation}: ${path}.width must be a positive integer EMU value`);
  }
  if (border.fill !== undefined) assertFill(border.fill, operation, `${path}.fill`);
}

function assertFill(fill: unknown, operation: string, path: string): void {
  if (!isRecord(fill)) throw new Error(`${operation}: ${path} must be a fill object`);
  if (fill.kind === "none") {
    assertOnlyKeys(fill, NONE_FILL_KEYS, operation, path);
    return;
  }
  if (fill.kind !== "solid")
    throw new Error(`${operation}: ${path} supports only solid and none fills`);
  assertOnlyKeys(fill, SOLID_FILL_KEYS, operation, path);
  if (!isRecord(fill.color) || fill.color.kind !== "srgb")
    throw new Error(`${operation}: ${path} supports only srgb solid colors`);
  assertOnlyKeys(fill.color, SRGB_COLOR_KEYS, operation, `${path}.color`);
  if (typeof fill.color.hex !== "string" || !/^[0-9A-Fa-f]{6}$/.test(fill.color.hex)) {
    throw new Error(`${operation}: ${path} srgb color must be a 6-digit hex value`);
  }
}

function assertOnlyKeys(
  value: Record<string, unknown>,
  allowed: ReadonlySet<string>,
  operation: string,
  path: string,
): void {
  for (const property of Object.keys(value)) {
    if (!allowed.has(property)) {
      throw new Error(`${operation}: ${path}.${property} is not a supported fill property`);
    }
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function hasSetValues(properties: EditableTableCellProperties): boolean {
  return (
    properties.fill !== undefined ||
    properties.borders !== undefined ||
    properties.marginLeft !== undefined ||
    properties.marginRight !== undefined ||
    properties.marginTop !== undefined ||
    properties.marginBottom !== undefined
  );
}

function stableValueEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right)) return false;
    if (left.length !== right.length) return false;
    return left.every((value, index) => stableValueEqual(value, right[index]));
  }
  if (isRecord(left) || isRecord(right)) {
    if (!isRecord(left) || !isRecord(right)) return false;
    const leftKeys = Object.keys(left).sort();
    const rightKeys = Object.keys(right).sort();
    if (!stableValueEqual(leftKeys, rightKeys)) return false;
    return leftKeys.every((key) => stableValueEqual(left[key], right[key]));
  }
  return false;
}
