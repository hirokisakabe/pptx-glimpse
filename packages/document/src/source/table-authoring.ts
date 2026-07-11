import { XMLBuilder } from "fast-xml-parser";

import { editReservedShapeId, sourceHandlesEqual } from "./edit-descriptors.js";
import type { PartPath, RelationshipId, SourceHandle } from "./handles.js";
import { nextNumberedName, nextRelationshipId } from "./package-graph-mutations.js";
import type { PptxSourceModel, PptxSourceModelAddTableEdit } from "./pptx-source-model.js";
import { parseShapeNodeXml } from "./shape-xml.js";
import type { Emu, Pt } from "./units.js";

export interface AddTableRunPropertiesInput {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly fontSize?: Pt;
  readonly fontFace?: string;
  readonly color?: string;
}
export interface AddTableRunInput {
  readonly text: string;
  readonly properties?: AddTableRunPropertiesInput;
  readonly hyperlink?: string;
}
export interface AddTableBorderInput {
  readonly width?: Emu;
  readonly color?: string;
  readonly dash?: "solid" | "dash" | "dot" | "dashDot";
}
export interface AddTableCellInput {
  readonly text?: string;
  readonly runs?: readonly AddTableRunInput[];
  readonly fill?: string;
  readonly align?: "left" | "center" | "right" | "justify";
  readonly verticalAlign?: "top" | "middle" | "bottom";
  readonly borders?: {
    readonly top?: AddTableBorderInput;
    readonly right?: AddTableBorderInput;
    readonly bottom?: AddTableBorderInput;
    readonly left?: AddTableBorderInput;
  };
  readonly colspan?: number;
  readonly rowspan?: number;
}
export interface AddTableRowInput {
  readonly height: Emu;
  readonly cells: readonly AddTableCellInput[];
}
export interface AddTableInput {
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly columnWidths: readonly Emu[];
  readonly rows: readonly AddTableRowInput[];
  readonly name?: string;
  readonly tableStyleId?: string;
}

const HYPERLINK_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const DEFAULT_TABLE_STYLE = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: true,
});

export function addTable(
  source: PptxSourceModel,
  slideHandle: SourceHandle,
  input: AddTableInput,
): PptxSourceModel {
  assertInput(input);
  const slideIndex = source.slides.findIndex((slide) =>
    sourceHandlesEqual(slide.handle, slideHandle),
  );
  if (slideIndex < 0)
    throw new Error("addTable: slide handle was not found in PptxSourceModel source");
  const slide = source.slides[slideIndex];
  const shapeId = nextShapeId(source, slide.partPath);
  const group = source.packageGraph.relationships.find(
    (item) => item.sourcePartPath === slide.partPath,
  );
  let relationships = group?.relationships ?? [];
  const hyperlinkIds = new Map<string, RelationshipId>();
  for (const row of input.rows)
    for (const cell of row.cells)
      for (const run of cell.runs ?? []) {
        if (run.hyperlink === undefined || hyperlinkIds.has(run.hyperlink)) continue;
        const id = nextRelationshipId(relationships);
        hyperlinkIds.set(run.hyperlink, id);
        relationships = [
          ...relationships,
          { id, type: HYPERLINK_REL, target: run.hyperlink, targetMode: "External" },
        ];
      }
  const xml = buildTableXml(input, shapeId, input.name?.trim() || `Table ${shapeId}`, hyperlinkIds);
  const table = parseShapeNodeXml(xml, slide.partPath, nextOrderingSlot(slide.shapes));
  if (table.kind !== "table") throw new Error("addTable: finalized XML did not parse as a table");
  const edit = {
    kind: "addTable",
    slidePartPath: slide.partPath,
    shapeId,
    xml,
  } satisfies PptxSourceModelAddTableEdit;
  return {
    ...source,
    packageGraph:
      hyperlinkIds.size === 0
        ? source.packageGraph
        : {
            ...source.packageGraph,
            relationships:
              group === undefined
                ? [
                    ...source.packageGraph.relationships,
                    { sourcePartPath: slide.partPath, relationships },
                  ]
                : source.packageGraph.relationships.map((item) =>
                    item === group ? { ...item, relationships } : item,
                  ),
          },
    slides: source.slides.map((candidate, index) =>
      index === slideIndex ? { ...candidate, shapes: [...candidate.shapes, table] } : candidate,
    ),
    edits: [...(source.edits ?? []), edit],
  };
}

function buildTableXml(
  input: AddTableInput,
  shapeId: string,
  name: string,
  links: ReadonlyMap<string, RelationshipId>,
): string {
  const mergeFlags = computeMergeFlags(input);
  return xmlBuilder.build({
    "p:graphicFrame": {
      "p:nvGraphicFramePr": {
        "p:cNvPr": { "@_id": shapeId, "@_name": name },
        "p:cNvGraphicFramePr": {},
        "p:nvPr": {},
      },
      "p:xfrm": {
        "a:off": { "@_x": String(input.offsetX), "@_y": String(input.offsetY) },
        "a:ext": { "@_cx": String(input.width), "@_cy": String(input.height) },
      },
      "a:graphic": {
        "a:graphicData": {
          "@_uri": "http://schemas.openxmlformats.org/drawingml/2006/table",
          "a:tbl": {
            "a:tblPr": {
              "@_firstRow": "1",
              "@_bandRow": "1",
              "a:tableStyleId": input.tableStyleId ?? DEFAULT_TABLE_STYLE,
            },
            "a:tblGrid": {
              "a:gridCol": input.columnWidths.map((width) => ({ "@_w": String(width) })),
            },
            "a:tr": input.rows.map((row, rowIndex) => ({
              "@_h": String(row.height),
              "a:tc": row.cells.map((cell, columnIndex) =>
                cellXml(cell, links, mergeFlags.get(`${rowIndex}:${columnIndex}`)),
              ),
            })),
          },
        },
      },
    },
  });
}

function cellXml(
  cell: AddTableCellInput,
  links: ReadonlyMap<string, RelationshipId>,
  merge?: { readonly horizontal: boolean; readonly vertical: boolean },
): Record<string, unknown> {
  const runs = cell.runs ?? [{ text: cell.text ?? "" }];
  return {
    ...(cell.colspan !== undefined ? { "@_gridSpan": String(cell.colspan) } : {}),
    ...(cell.rowspan !== undefined ? { "@_rowSpan": String(cell.rowspan) } : {}),
    ...(merge?.horizontal ? { "@_hMerge": "1" } : {}),
    ...(merge?.vertical ? { "@_vMerge": "1" } : {}),
    "a:txBody": {
      "a:bodyPr": {
        "@_anchor": ({ top: "t", middle: "ctr", bottom: "b" } as const)[
          cell.verticalAlign ?? "top"
        ],
      },
      "a:lstStyle": {},
      "a:p": {
        ...(cell.align !== undefined
          ? {
              "a:pPr": {
                "@_algn": ({ left: "l", center: "ctr", right: "r", justify: "just" } as const)[
                  cell.align
                ],
              },
            }
          : {}),
        "a:r": runs.map((run) => ({
          ...(run.properties !== undefined || run.hyperlink !== undefined
            ? { "a:rPr": runPropertiesXml(run, links) }
            : {}),
          "a:t": textElementValue(run.text),
        })),
        "a:endParaRPr": {},
      },
    },
    "a:tcPr": {
      ...(cell.fill !== undefined ? { "a:solidFill": colorXml(cell.fill) } : {}),
      ...bordersXml(cell.borders),
    },
  };
}

function computeMergeFlags(
  input: AddTableInput,
): ReadonlyMap<string, { readonly horizontal: boolean; readonly vertical: boolean }> {
  const result = new Map<string, { horizontal: boolean; vertical: boolean }>();
  input.rows.forEach((row, rowIndex) =>
    row.cells.forEach((cell, columnIndex) => {
      const currentKey = `${rowIndex}:${columnIndex}`;
      if (result.has(currentKey)) {
        if (!isEmptyMergeContinuation(cell)) {
          throw new Error("addTable: cells covered by a span must be empty placeholders");
        }
        return;
      }
      const colspan = cell.colspan ?? 1;
      const rowspan = cell.rowspan ?? 1;
      if (
        columnIndex + colspan > input.columnWidths.length ||
        rowIndex + rowspan > input.rows.length
      ) {
        throw new Error("addTable: cell span exceeds the table grid");
      }
      for (let y = rowIndex; y < rowIndex + rowspan; y += 1) {
        for (let x = columnIndex; x < columnIndex + colspan; x += 1) {
          if (x === columnIndex && y === rowIndex) continue;
          const key = `${y}:${x}`;
          if (result.has(key)) throw new Error("addTable: cell spans overlap");
          result.set(key, { horizontal: x > columnIndex, vertical: y > rowIndex });
        }
      }
    }),
  );
  return result;
}

function runPropertiesXml(
  run: AddTableRunInput,
  links: ReadonlyMap<string, RelationshipId>,
): Record<string, unknown> {
  const p = run.properties;
  return {
    ...(p?.bold !== undefined ? { "@_b": p.bold ? "1" : "0" } : {}),
    ...(p?.italic !== undefined ? { "@_i": p.italic ? "1" : "0" } : {}),
    ...(p?.underline !== undefined ? { "@_u": p.underline ? "sng" : "none" } : {}),
    ...(p?.fontSize !== undefined ? { "@_sz": String(Math.round(p.fontSize * 100)) } : {}),
    ...(p?.color !== undefined ? { "a:solidFill": colorXml(p.color) } : {}),
    ...(p?.fontFace !== undefined
      ? { "a:latin": { "@_typeface": p.fontFace }, "a:ea": { "@_typeface": p.fontFace } }
      : {}),
    ...(run.hyperlink !== undefined
      ? { "a:hlinkClick": { "@_r:id": links.get(run.hyperlink) } }
      : {}),
  };
}
function bordersXml(borders: AddTableCellInput["borders"]): Record<string, unknown> {
  if (borders === undefined) return {};
  return Object.fromEntries(
    (["left", "right", "top", "bottom"] as const).flatMap((side) => {
      const border = borders[side];
      if (border === undefined) return [];
      const tag = { left: "a:lnL", right: "a:lnR", top: "a:lnT", bottom: "a:lnB" }[side];
      return [
        [
          tag,
          {
            ...(border.width !== undefined ? { "@_w": String(border.width) } : {}),
            "a:solidFill": colorXml(border.color ?? "000000"),
            ...(border.dash !== undefined ? { "a:prstDash": { "@_val": border.dash } } : {}),
          },
        ],
      ];
    }),
  );
}
function colorXml(hex: string): Record<string, unknown> {
  if (!/^[0-9a-f]{6}$/i.test(hex)) throw new Error("addTable: colors must be 6-digit RGB hex");
  return { "a:srgbClr": { "@_val": hex.toUpperCase() } };
}
function textElementValue(text: string): unknown {
  return text.startsWith(" ") || text.endsWith(" ")
    ? { "@_xml:space": "preserve", "#text": text }
    : text;
}
function isEmptyMergeContinuation(cell: AddTableCellInput): boolean {
  return Object.keys(cell).length === 0;
}
function assertInput(input: AddTableInput): void {
  for (const [name, value] of Object.entries({
    offsetX: input.offsetX,
    offsetY: input.offsetY,
    width: input.width,
    height: input.height,
  }))
    if (
      typeof value !== "number" ||
      !Number.isFinite(value) ||
      ((name === "width" || name === "height") && value <= 0)
    )
      throw new Error(`addTable: ${name} must be a valid EMU value`);
  if (input.columnWidths.length === 0 || input.rows.length === 0)
    throw new Error("addTable: columns and rows must not be empty");
  if (input.columnWidths.some((v) => !Number.isFinite(v) || v <= 0))
    throw new Error("addTable: column widths must be positive");
  for (const row of input.rows) {
    if (!Number.isFinite(row.height) || row.height <= 0)
      throw new Error("addTable: row heights must be positive");
    if (row.cells.length !== input.columnWidths.length)
      throw new Error("addTable: each row must contain one cell per grid column");
    for (const cell of row.cells) {
      if (cell.text !== undefined && cell.runs !== undefined)
        throw new Error("addTable: specify cell text or runs, not both");
      if (
        !Number.isInteger(cell.colspan ?? 1) ||
        !Number.isInteger(cell.rowspan ?? 1) ||
        (cell.colspan ?? 1) < 1 ||
        (cell.rowspan ?? 1) < 1
      )
        throw new Error("addTable: spans must be positive integers");
      if (cell.runs?.length === 0) throw new Error("addTable: runs must not be empty");
      for (const run of cell.runs ?? []) {
        if (typeof run.text !== "string") throw new Error("addTable: run text must be a string");
        if (run.hyperlink !== undefined && run.hyperlink.trim().length === 0)
          throw new Error("addTable: hyperlink must not be empty");
        if (
          run.properties?.fontSize !== undefined &&
          (!Number.isFinite(run.properties.fontSize) || run.properties.fontSize <= 0)
        )
          throw new Error("addTable: font size must be positive");
      }
    }
  }
}
function nextShapeId(source: PptxSourceModel, partPath: PartPath): string {
  // p:spTree reserves id 1 for its non-visual group properties.
  const used = new Set<string>(["1"]);
  const slide = source.slides.find((s) => s.partPath === partPath);
  for (const shape of slide?.shapes ?? [])
    if (shape.nodeId !== undefined) used.add(String(shape.nodeId));
  for (const edit of source.edits ?? []) {
    const id = editReservedShapeId(edit, partPath);
    if (id !== undefined) used.add(id);
  }
  return nextNumberedName(used, /^(\d+)$/, String);
}
function nextOrderingSlot(
  shapes: readonly { readonly handle?: { readonly orderingSlot?: number } }[],
): number {
  return Math.max(0, ...shapes.map((shape) => shape.handle?.orderingSlot ?? 0)) + 1;
}
