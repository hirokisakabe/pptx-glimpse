import { zipSync } from "fflate";

import { editReservedShapeId, sourceHandlesEqual } from "./edit-descriptors.js";
import { relativeTarget } from "./editing-shared.js";
import { asRelationshipId, type PartPath } from "./handles.js";
import type {
  PackageGraph,
  PartRelationships,
  PptxSourceModel,
  PptxSourceModelAddChartEdit,
  SourceHandle,
  SourceShapeNode,
} from "./index.js";
import {
  addPackagePart,
  addPartRelationship,
  nextNumberedName,
  nextNumberedPartPath,
  nextRelationshipId,
} from "./package-graph-mutations.js";
import { parseShapeNodeXml } from "./shape-xml.js";
import type { Emu } from "./units.js";

export type NativeChartType = "bar" | "line" | "pie" | "area" | "doughnut" | "radar";
export type NativeRadarStyle = "standard" | "marker" | "filled";
export type NativeChartLegendPosition = "b" | "t" | "l" | "r" | "tr";

export interface AddChartSeriesInput {
  readonly name?: string;
  readonly categories: readonly string[];
  readonly values: readonly number[];
  /** Six-digit sRGB color, with or without `#`. */
  readonly color?: string;
}

export interface AddChartAxisInput {
  readonly hidden?: boolean;
  readonly lineVisible?: boolean;
  readonly gridLinesVisible?: boolean;
  readonly title?: string;
}

export interface AddChartPlotLayoutInput {
  /** Fraction of chart width/height in the range 0..1. */
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
}

export interface AddChartInput {
  readonly chartType: NativeChartType;
  readonly series: readonly AddChartSeriesInput[];
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly name?: string;
  readonly title?: string;
  readonly showLegend?: boolean;
  readonly legendPosition?: NativeChartLegendPosition;
  readonly radarStyle?: NativeRadarStyle;
  readonly holeSize?: number;
  readonly categoryAxis?: AddChartAxisInput;
  readonly valueAxis?: AddChartAxisInput;
  readonly plotLayout?: AddChartPlotLayoutInput;
}

const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const PACKAGE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
const encoder = new TextEncoder();

export function addChart(
  source: PptxSourceModel,
  slideHandle: SourceHandle,
  input: AddChartInput,
): PptxSourceModel {
  assertChartInput(input);
  const slideIndex = source.slides.findIndex((slide) =>
    sourceHandlesEqual(slide.handle, slideHandle),
  );
  if (slideIndex === -1)
    throw new Error("addChart: slide handle was not found in PptxSourceModel source");
  const slide = source.slides[slideIndex];
  const slideRelationships = relationshipGroup(source.packageGraph, slide.partPath);
  const relationshipId = nextRelationshipId(slideRelationships.relationships);
  const reserved = (source.edits ?? []).flatMap((edit) =>
    edit.kind === "addChart" ? [edit.chartPartPath, edit.workbookPartPath] : [],
  );
  const chartPartPath = nextNumberedPartPath(
    source.packageGraph,
    reserved,
    "ppt/charts/chart",
    ".xml",
  );
  const workbookPartPath = nextNumberedPartPath(
    source.packageGraph,
    reserved,
    "ppt/embeddings/Microsoft_Excel_Worksheet",
    ".xlsx",
  );
  const shapeId = nextChartShapeId(source, slide.partPath);
  const chartNumber = Number(/(\d+)\.xml$/.exec(chartPartPath)?.[1] ?? 1);
  const chartXml = buildChartXml(input, chartNumber);
  const workbookBytes = buildEmbeddedWorkbook(input.series);
  const chartRelationships: PartRelationships = {
    sourcePartPath: chartPartPath,
    relationships: [
      {
        id: asRelationshipId("rId1"),
        type: PACKAGE_REL_TYPE,
        target: relativeTarget(chartPartPath, workbookPartPath),
      },
    ],
  };
  let graph = addPackagePart(source.packageGraph, {
    partPath: workbookPartPath,
    contentType: XLSX_CONTENT_TYPE,
    bytes: workbookBytes,
  });
  graph = addPackagePart(graph, {
    partPath: chartPartPath,
    contentType: CHART_CONTENT_TYPE,
    bytes: encoder.encode(chartXml),
    relationships: chartRelationships,
  });
  graph = addPartRelationship(graph, slide.partPath, {
    id: relationshipId,
    type: CHART_REL_TYPE,
    target: relativeTarget(slide.partPath, chartPartPath),
  });
  const name = input.name?.trim() || `Chart ${shapeId}`;
  const xml = buildChartFrameXml(shapeId, name, relationshipId, input);
  const chart = parseShapeNodeXml(xml, slide.partPath, nextOrderingSlot(slide.shapes));
  if (chart.kind !== "chart")
    throw new Error("addChart: finalized chart XML did not parse as a chart");
  const edit = {
    kind: "addChart",
    slidePartPath: slide.partPath,
    shapeId,
    relationshipId,
    chartPartPath,
    workbookPartPath,
    xml,
  } satisfies PptxSourceModelAddChartEdit;
  return {
    ...source,
    packageGraph: graph,
    slides: source.slides.map((candidate, index) =>
      index === slideIndex ? { ...candidate, shapes: [...candidate.shapes, chart] } : candidate,
    ),
    edits: [...(source.edits ?? []), edit],
  };
}

function buildChartFrameXml(
  shapeId: string,
  name: string,
  relationshipId: string,
  input: AddChartInput,
): string {
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${shapeId}" name="${escapeXml(name)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${input.offsetX}" y="${input.offsetY}"/><a:ext cx="${input.width}" cy="${input.height}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${relationshipId}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

function buildChartXml(input: AddChartInput, chartNumber: number): string {
  const categories = input.series[0].categories;
  const chartTag = `${input.chartType}Chart`;
  const seriesXml = input.series
    .map((series, index) => buildSeriesXml(input.chartType, series, index, categories.length))
    .join("");
  const axisIds = [100000 + chartNumber * 2, 100001 + chartNumber * 2];
  const chartProperties = chartTypeProperties(input, axisIds);
  const axes = usesAxes(input.chartType) ? buildAxes(input, axisIds) : "";
  const title = input.title === undefined ? "" : buildTitleXml(input.title);
  const legend =
    input.showLegend === true
      ? `<c:legend><c:legendPos val="${input.legendPosition ?? "r"}"/><c:layout/><c:overlay val="0"/></c:legend>`
      : "";
  const layout =
    input.plotLayout === undefined ? "<c:layout/>" : buildManualLayout(input.plotLayout);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${title}<c:autoTitleDeleted val="${input.title === undefined ? 1 : 0}"/><c:plotArea>${layout}<c:${chartTag}>${chartProperties.beforeSeries}${seriesXml}${chartProperties.afterSeries}</c:${chartTag}>${axes}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;
}

function chartTypeProperties(
  input: AddChartInput,
  axisIds: readonly number[],
): { readonly beforeSeries: string; readonly afterSeries: string } {
  const axisRefs = usesAxes(input.chartType)
    ? `<c:axId val="${axisIds[0]}"/><c:axId val="${axisIds[1]}"/>`
    : "";
  switch (input.chartType) {
    case "bar":
      return {
        beforeSeries: `<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>`,
        afterSeries: axisRefs,
      };
    case "line":
      return {
        beforeSeries: `<c:grouping val="standard"/><c:varyColors val="0"/>`,
        afterSeries: axisRefs,
      };
    case "area":
      return {
        beforeSeries: `<c:grouping val="standard"/><c:varyColors val="0"/>`,
        afterSeries: axisRefs,
      };
    case "pie":
      return { beforeSeries: `<c:varyColors val="1"/>`, afterSeries: "" };
    case "doughnut":
      return {
        beforeSeries: `<c:varyColors val="1"/>`,
        afterSeries: `<c:holeSize val="${input.holeSize ?? 50}"/>`,
      };
    case "radar":
      return {
        beforeSeries: `<c:radarStyle val="${input.radarStyle ?? "standard"}"/><c:varyColors val="0"/>`,
        afterSeries: axisRefs,
      };
  }
}

function buildSeriesXml(
  chartType: NativeChartType,
  series: AddChartSeriesInput,
  index: number,
  pointCount: number,
): string {
  const column = spreadsheetColumn(index + 2);
  const name = series.name ?? `Series ${index + 1}`;
  const color =
    series.color === undefined
      ? ""
      : `<c:spPr><a:solidFill><a:srgbClr val="${normalizeColor(series.color)}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${normalizeColor(series.color)}"/></a:solidFill></a:ln></c:spPr>`;
  const marker =
    chartType === "line" || chartType === "radar"
      ? `<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>`
      : "";
  const smooth = chartType === "line" ? `<c:smooth val="0"/>` : "";
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:strRef><c:f>Sheet1!$${column}$1</c:f>${stringCache([name])}</c:strRef></c:tx>${color}${marker}<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${pointCount + 1}</c:f>${stringCache(series.categories)}</c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$${column}$2:$${column}$${pointCount + 1}</c:f>${numberCache(series.values)}</c:numRef></c:val>${smooth}</c:ser>`;
}

function stringCache(values: readonly string[]): string {
  return `<c:strCache><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${escapeXml(value)}</c:v></c:pt>`).join("")}</c:strCache>`;
}

function numberCache(values: readonly number[]): string {
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join("")}</c:numCache>`;
}

function buildAxes(input: AddChartInput, ids: readonly number[]): string {
  const cat = input.categoryAxis ?? {};
  const val = input.valueAxis ?? {};
  return `<c:catAx><c:axId val="${ids[0]}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${cat.hidden === true ? 1 : 0}"/><c:axPos val="b"/>${cat.gridLinesVisible === true ? "<c:majorGridlines/>" : ""}${cat.title === undefined ? "" : buildTitleXml(cat.title)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axisShapeProperties(cat.lineVisible)}<c:crossAx val="${ids[1]}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx><c:valAx><c:axId val="${ids[1]}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${val.hidden === true ? 1 : 0}"/><c:axPos val="l"/>${val.gridLinesVisible === false ? "" : "<c:majorGridlines/>"}${val.title === undefined ? "" : buildTitleXml(val.title)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axisShapeProperties(val.lineVisible)}<c:crossAx val="${ids[0]}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
}

function axisShapeProperties(visible: boolean | undefined): string {
  return visible === false ? `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>` : "";
}

function buildTitleXml(title: string): string {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>`;
}

function buildManualLayout(layout: AddChartPlotLayoutInput): string {
  const value = (tag: string, number: number | undefined) =>
    number === undefined ? "" : `<c:${tag} val="${number}"/>`;
  return `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/>${value("x", layout.x)}${value("y", layout.y)}${value("w", layout.width)}${value("h", layout.height)}</c:manualLayout></c:layout>`;
}

function buildEmbeddedWorkbook(series: readonly AddChartSeriesInput[]): Uint8Array {
  const rows = Math.max(...series.map((item) => item.categories.length));
  const cells: string[] = [`<c r="A1" t="inlineStr"><is><t>Category</t></is></c>`];
  series.forEach((item, index) =>
    cells.push(
      `<c r="${spreadsheetColumn(index + 2)}1" t="inlineStr"><is><t>${escapeXml(item.name ?? `Series ${index + 1}`)}</t></is></c>`,
    ),
  );
  const rowXml = [`<row r="1">${cells.join("")}</row>`];
  for (let row = 0; row < rows; row += 1) {
    const rowCells = [
      `<c r="A${row + 2}" t="inlineStr"><is><t>${escapeXml(series[0].categories[row])}</t></is></c>`,
    ];
    series.forEach((item, index) =>
      rowCells.push(
        `<c r="${spreadsheetColumn(index + 2)}${row + 2}"><v>${item.values[row]}</v></c>`,
      ),
    );
    rowXml.push(`<row r="${row + 2}">${rowCells.join("")}</row>`);
  }
  const xml = (value: string) =>
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${value}`);
  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
    "xl/workbook.xml": xml(
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": xml(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${spreadsheetColumn(series.length + 1)}${rows + 1}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${rowXml.join("")}</sheetData></worksheet>`,
    ),
    "xl/styles.xml": xml(
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
    ),
  });
}

function relationshipGroup(graph: PackageGraph, partPath: PartPath): PartRelationships {
  return (
    graph.relationships.find((item) => item.sourcePartPath === partPath) ?? {
      sourcePartPath: partPath,
      relationships: [],
    }
  );
}

function nextChartShapeId(source: PptxSourceModel, slidePartPath: PartPath): string {
  const used = new Set<string>();
  const collect = (shapes: readonly SourceShapeNode[]) =>
    shapes.forEach((shape) => {
      if (shape.nodeId !== undefined) used.add(String(shape.nodeId));
      if (shape.kind === "group") collect(shape.children);
    });
  collect(source.slides.find((slide) => slide.partPath === slidePartPath)?.shapes ?? []);
  for (const edit of source.edits ?? []) {
    const id = editReservedShapeId(edit, slidePartPath);
    if (id !== undefined) used.add(id);
  }
  return nextNumberedName(used, /^(\d+)$/, String);
}

function nextOrderingSlot(shapes: readonly { readonly handle?: SourceHandle }[]): number {
  return shapes.reduce((max, shape) => Math.max(max, shape.handle?.orderingSlot ?? -1), -1) + 1;
}

function assertChartInput(input: AddChartInput): void {
  for (const [field, value] of [
    ["offsetX", input.offsetX],
    ["offsetY", input.offsetY],
  ] as const) {
    if (!Number.isFinite(value)) throw new Error(`addChart: ${field} must be a finite EMU value`);
  }
  for (const [field, value] of [
    ["width", input.width],
    ["height", input.height],
  ] as const) {
    if (!Number.isFinite(value) || value <= 0)
      throw new Error(`addChart: ${field} must be a finite positive EMU value`);
  }
  if (input.series.length === 0) throw new Error("addChart: series must not be empty");
  const count = input.series[0].categories.length;
  if (count === 0) throw new Error("addChart: series categories must not be empty");
  input.series.forEach((series) => {
    if (series.categories.length !== count || series.values.length !== count)
      throw new Error("addChart: every series must have matching category and value counts");
    if (series.categories.some((category, index) => category !== input.series[0].categories[index]))
      throw new Error("addChart: every series must use identical category labels");
    if (series.values.some((value) => !Number.isFinite(value)))
      throw new Error("addChart: values must be finite numbers");
    if (series.color !== undefined) normalizeColor(series.color);
  });
  if (input.radarStyle !== undefined && input.chartType !== "radar")
    throw new Error("addChart: radarStyle is only valid for radar charts");
  if (
    input.holeSize !== undefined &&
    (input.chartType !== "doughnut" ||
      !Number.isInteger(input.holeSize) ||
      input.holeSize < 10 ||
      input.holeSize > 90)
  )
    throw new Error("addChart: holeSize must be an integer from 10 through 90 for doughnut charts");
  for (const value of Object.values(input.plotLayout ?? {}))
    if (!Number.isFinite(value) || value < 0 || value > 1)
      throw new Error("addChart: plotLayout values must be finite numbers from 0 through 1");
}

function normalizeColor(value: string): string {
  const color = value.replace(/^#/, "").toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(color))
    throw new Error("addChart: series color must be a six-digit sRGB color");
  return color;
}
function usesAxes(type: NativeChartType): boolean {
  return type !== "pie" && type !== "doughnut";
}
function spreadsheetColumn(index: number): string {
  let value = "";
  for (let n = index; n > 0; n = Math.floor((n - 1) / 26))
    value = String.fromCharCode(65 + ((n - 1) % 26)) + value;
  return value;
}
function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
