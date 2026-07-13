import { zipSync } from "fflate";

import { getAttr, getChild, parseXml } from "../reader/xml.js";
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
import type { AuthoringColorInput } from "./shape-xml.js";
import { parseShapeNodeXml } from "./shape-xml.js";
import type { SourceDashStyle } from "./shapes.js";
import type { Emu, Pt } from "./units.js";

export type NativeChartType = "bar" | "line" | "pie" | "area" | "doughnut" | "radar";
export type NativeRadarStyle = "standard" | "marker" | "filled";
export type NativeChartLegendPosition = "b" | "t" | "l" | "r" | "tr";
export type AddChartColorInput = AuthoringColorInput;
export type AddChartFillInput =
  | { readonly kind: "none" }
  | { readonly kind: "solid"; readonly color: AddChartColorInput };
export interface AddChartOutlineInput {
  readonly width?: Emu;
  readonly fill?: AddChartFillInput;
  readonly dash?: SourceDashStyle;
}
export interface AddChartAreaStyleInput {
  readonly fill?: AddChartFillInput;
  readonly outline?: AddChartOutlineInput;
}
export interface AddChartTextStyleInput {
  readonly fontFace?: string;
  readonly fontSize?: Pt;
  readonly color?: AddChartColorInput;
  readonly bold?: boolean;
  readonly italic?: boolean;
}
export interface AddChartNumberFormatInput {
  readonly formatCode: string;
  readonly sourceLinked?: boolean;
}
export type NativeChartTickMark = "none" | "inside" | "outside" | "cross";
export type NativeChartLabelPosition = "nextTo" | "high" | "low" | "none";
export type NativeChartMarkerSymbol =
  | "auto"
  | "circle"
  | "dash"
  | "diamond"
  | "dot"
  | "none"
  | "plus"
  | "square"
  | "star"
  | "triangle"
  | "x";

export interface AddChartMarkerInput {
  readonly symbol?: NativeChartMarkerSymbol;
  /** Marker size in points, from 2 through 72. */
  readonly size?: number;
  readonly fill?: AddChartFillInput;
  readonly outline?: AddChartOutlineInput;
}

export interface AddChartDataPointInput {
  readonly index: number;
  readonly fill?: AddChartFillInput;
  readonly outline?: AddChartOutlineInput;
}

export interface AddChartSeriesInput {
  readonly name?: string;
  readonly categories: readonly string[];
  readonly values: readonly number[];
  /** Six-digit sRGB color, with or without `#`. */
  readonly color?: string;
  readonly fill?: AddChartFillInput;
  readonly outline?: AddChartOutlineInput;
  readonly marker?: AddChartMarkerInput;
  readonly dataPoints?: readonly AddChartDataPointInput[];
}

export interface AddChartAxisInput {
  readonly hidden?: boolean;
  readonly lineVisible?: boolean;
  readonly gridLinesVisible?: boolean;
  readonly title?: string;
  readonly majorTickMark?: NativeChartTickMark;
  readonly labelPosition?: NativeChartLabelPosition;
  readonly numberFormat?: AddChartNumberFormatInput;
  readonly line?: AddChartOutlineInput;
  readonly majorGridline?: AddChartOutlineInput;
  readonly textStyle?: AddChartTextStyleInput;
  readonly showMultiLevelLabels?: boolean;
}

export interface AddChartPlotLayoutInput {
  readonly coordinateMode?: "factor" | "edge";
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
  readonly titleStyle?: AddChartTextStyleInput;
  readonly displayBlanksAs?: "gap" | "zero" | "span";
  readonly roundedCorners?: boolean;
  readonly chartArea?: AddChartAreaStyleInput;
  readonly plotArea?: AddChartAreaStyleInput;
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
const decoder = new TextDecoder();
const CHART_TYPES: ReadonlySet<string> = new Set([
  "bar",
  "line",
  "pie",
  "area",
  "doughnut",
  "radar",
]);
const LEGEND_POSITIONS: ReadonlySet<string> = new Set(["b", "t", "l", "r", "tr"]);
const RADAR_STYLES: ReadonlySet<string> = new Set(["standard", "marker", "filled"]);
const BLANK_DISPLAY_MODES: ReadonlySet<string> = new Set(["gap", "zero", "span"]);
const TICK_MARKS: ReadonlySet<string> = new Set(["none", "inside", "outside", "cross"]);
const LABEL_POSITIONS: ReadonlySet<string> = new Set(["nextTo", "high", "low", "none"]);
const MARKER_SYMBOLS: ReadonlySet<string> = new Set([
  "auto",
  "circle",
  "dash",
  "diamond",
  "dot",
  "none",
  "plus",
  "square",
  "star",
  "triangle",
  "x",
]);
const DASH_STYLES: ReadonlySet<string> = new Set([
  "solid",
  "dash",
  "dot",
  "dashDot",
  "lgDash",
  "lgDashDot",
  "lgDashDotDot",
  "sysDash",
  "sysDot",
]);
const XLSX_MAX_SERIES = 16_383;
const XLSX_MAX_DATA_POINTS = 1_048_575;
const MAX_LINE_WIDTH = 20_116_800;
const MIN_TEXT_SIZE = 100;
const MAX_TEXT_SIZE = 400_000;

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
  const title = input.title === undefined ? "" : buildTitleXml(input.title, input.titleStyle);
  const legend =
    input.showLegend === true
      ? `<c:legend><c:legendPos val="${input.legendPosition ?? "r"}"/><c:layout/><c:overlay val="0"/></c:legend>`
      : "";
  const layout =
    input.plotLayout === undefined ? "<c:layout/>" : buildManualLayout(input.plotLayout);
  const plotAreaStyle = buildAreaStyleXml(input.plotArea);
  const chartAreaStyle = buildAreaStyleXml(input.chartArea);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="${input.roundedCorners === true ? 1 : 0}"/><c:chart>${title}<c:autoTitleDeleted val="${input.title === undefined ? 1 : 0}"/><c:plotArea>${layout}<c:${chartTag}>${chartProperties.beforeSeries}${seriesXml}${chartProperties.afterSeries}</c:${chartTag}>${axes}${plotAreaStyle}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="${input.displayBlanksAs ?? "gap"}"/><c:showDLblsOverMax val="0"/></c:chart>${chartAreaStyle}<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;
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
  const legacyColor =
    series.color === undefined
      ? undefined
      : ({ kind: "solid", color: { kind: "srgb", hex: normalizeColor(series.color) } } as const);
  const seriesProperties = buildShapePropertiesXml(
    series.fill ?? legacyColor,
    series.outline ?? (legacyColor === undefined ? undefined : { fill: legacyColor }),
  );
  const marker = chartType === "line" || chartType === "radar" ? buildMarkerXml(series.marker) : "";
  const dataPoints = (series.dataPoints ?? []).map(buildDataPointXml).join("");
  const smooth = chartType === "line" ? `<c:smooth val="0"/>` : "";
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:strRef><c:f>Sheet1!$${column}$1</c:f>${stringCache([name])}</c:strRef></c:tx>${seriesProperties}${marker}${dataPoints}<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${pointCount + 1}</c:f>${stringCache(series.categories)}</c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$${column}$2:$${column}$${pointCount + 1}</c:f>${numberCache(series.values)}</c:numRef></c:val>${smooth}</c:ser>`;
}

function stringCache(values: readonly string[]): string {
  return `<c:strCache><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}">${textElement("c:v", value)}</c:pt>`).join("")}</c:strCache>`;
}

function numberCache(values: readonly number[]): string {
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join("")}</c:numCache>`;
}

function buildAxes(input: AddChartInput, ids: readonly number[]): string {
  const cat = input.categoryAxis ?? {};
  const val = input.valueAxis ?? {};
  return `<c:catAx><c:axId val="${ids[0]}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${cat.hidden === true ? 1 : 0}"/><c:axPos val="b"/>${buildGridlineXml(cat, false)}${cat.title === undefined ? "" : buildTitleXml(cat.title)}${buildNumberFormatXml(cat.numberFormat)}<c:majorTickMark val="${tickMarkToken(cat.majorTickMark)}"/><c:minorTickMark val="none"/><c:tickLblPos val="${cat.labelPosition ?? "nextTo"}"/>${axisShapeProperties(cat)}${buildTextPropertiesXml(cat.textStyle)}<c:crossAx val="${ids[1]}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>${cat.showMultiLevelLabels === undefined ? "" : `<c:noMultiLvlLbl val="${cat.showMultiLevelLabels ? 0 : 1}"/>`}</c:catAx><c:valAx><c:axId val="${ids[1]}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${val.hidden === true ? 1 : 0}"/><c:axPos val="l"/>${buildGridlineXml(val, true)}${val.title === undefined ? "" : buildTitleXml(val.title)}${buildNumberFormatXml(val.numberFormat)}<c:majorTickMark val="${tickMarkToken(val.majorTickMark)}"/><c:minorTickMark val="none"/><c:tickLblPos val="${val.labelPosition ?? "nextTo"}"/>${axisShapeProperties(val)}${buildTextPropertiesXml(val.textStyle)}<c:crossAx val="${ids[0]}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
}

function buildGridlineXml(axis: AddChartAxisInput, visibleByDefault: boolean): string {
  const visible = axis.gridLinesVisible ?? (axis.majorGridline !== undefined || visibleByDefault);
  if (!visible) return "";
  const properties =
    axis.majorGridline === undefined ? "" : buildShapePropertiesXml(undefined, axis.majorGridline);
  return `<c:majorGridlines>${properties}</c:majorGridlines>`;
}

function buildNumberFormatXml(format: AddChartNumberFormatInput | undefined): string {
  return `<c:numFmt formatCode="${escapeXml(format?.formatCode ?? "General")}" sourceLinked="${format?.sourceLinked === false ? 0 : 1}"/>`;
}

function tickMarkToken(value: NativeChartTickMark | undefined): string {
  if (value === "inside") return "in";
  if (value === "outside") return "out";
  return value ?? "none";
}

function axisShapeProperties(axis: AddChartAxisInput): string {
  if (axis.line !== undefined) return buildShapePropertiesXml(undefined, axis.line);
  return axis.lineVisible === false
    ? buildShapePropertiesXml(undefined, { fill: { kind: "none" } })
    : "";
}

function buildTitleXml(title: string, style?: AddChartTextStyleInput): string {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>${buildRunPropertiesXml(style)}${textElement("a:t", title)}</a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>`;
}

function buildManualLayout(layout: AddChartPlotLayoutInput): string {
  const value = (tag: string, number: number | undefined) =>
    number === undefined ? "" : `<c:${tag} val="${number}"/>`;
  const mode = layout.coordinateMode ?? "factor";
  return `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="${mode}"/><c:yMode val="${mode}"/><c:wMode val="${mode}"/><c:hMode val="${mode}"/>${value("x", layout.x)}${value("y", layout.y)}${value("w", layout.width)}${value("h", layout.height)}</c:manualLayout></c:layout>`;
}

function buildAreaStyleXml(style: AddChartAreaStyleInput | undefined): string {
  return style === undefined ? "" : buildShapePropertiesXml(style.fill, style.outline);
}

function buildShapePropertiesXml(
  fill: AddChartFillInput | undefined,
  outline: AddChartOutlineInput | undefined,
): string {
  return fill === undefined && outline === undefined
    ? ""
    : `<c:spPr>${buildFillXml(fill)}${outline === undefined ? "" : buildOutlineXml(outline)}</c:spPr>`;
}

function buildFillXml(fill: AddChartFillInput | undefined): string {
  if (fill === undefined) return "";
  return fill.kind === "none"
    ? `<a:noFill/>`
    : `<a:solidFill>${buildColorXml(fill.color)}</a:solidFill>`;
}

function buildOutlineXml(outline: AddChartOutlineInput): string {
  const width = outline.width === undefined ? "" : ` w="${outline.width}"`;
  const dash = outline.dash === undefined ? "" : `<a:prstDash val="${outline.dash}"/>`;
  return `<a:ln${width}>${buildFillXml(outline.fill)}${dash}</a:ln>`;
}

function buildColorXml(color: AddChartColorInput): string {
  const transforms = (color.transforms ?? [])
    .map((transform) => `<a:alpha val="${transform.value}"/>`)
    .join("");
  return `<a:srgbClr val="${normalizeColor(color.hex)}">${transforms}</a:srgbClr>`;
}

function buildRunPropertiesXml(style: AddChartTextStyleInput | undefined): string {
  const attributes = [
    `lang="en-US"`,
    ...(style?.fontSize === undefined ? [] : [`sz="${Math.round(style.fontSize * 100)}"`]),
    ...(style?.bold === undefined ? [] : [`b="${style.bold ? 1 : 0}"`]),
    ...(style?.italic === undefined ? [] : [`i="${style.italic ? 1 : 0}"`]),
  ].join(" ");
  return `<a:rPr ${attributes}>${buildTextStyleChildren(style)}</a:rPr>`;
}

function buildTextPropertiesXml(style: AddChartTextStyleInput | undefined): string {
  if (style === undefined) return "";
  const attributes = [
    ...(style.fontSize === undefined ? [] : [`sz="${Math.round(style.fontSize * 100)}"`]),
    ...(style.bold === undefined ? [] : [`b="${style.bold ? 1 : 0}"`]),
    ...(style.italic === undefined ? [] : [`i="${style.italic ? 1 : 0}"`]),
  ].join(" ");
  const attributeText = attributes.length === 0 ? "" : ` ${attributes}`;
  return `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr${attributeText}>${buildTextStyleChildren(style)}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
}

function buildTextStyleChildren(style: AddChartTextStyleInput | undefined): string {
  if (style === undefined) return "";
  const color =
    style.color === undefined ? "" : `<a:solidFill>${buildColorXml(style.color)}</a:solidFill>`;
  const fonts =
    style.fontFace === undefined
      ? ""
      : `<a:latin typeface="${escapeXml(style.fontFace)}"/><a:ea typeface="${escapeXml(style.fontFace)}"/><a:cs typeface="${escapeXml(style.fontFace)}"/>`;
  return color + fonts;
}

function buildMarkerXml(marker: AddChartMarkerInput | undefined): string {
  const resolved = marker ?? {};
  const properties = buildShapePropertiesXml(resolved.fill, resolved.outline);
  return `<c:marker><c:symbol val="${resolved.symbol ?? "circle"}"/><c:size val="${resolved.size ?? 5}"/>${properties}</c:marker>`;
}

function buildDataPointXml(point: AddChartDataPointInput): string {
  return `<c:dPt><c:idx val="${point.index}"/>${buildShapePropertiesXml(point.fill, point.outline)}</c:dPt>`;
}

function buildEmbeddedWorkbook(series: readonly AddChartSeriesInput[]): Uint8Array {
  const rows = Math.max(...series.map((item) => item.categories.length));
  const cells: string[] = [`<c r="A1" t="inlineStr"><is><t>Category</t></is></c>`];
  series.forEach((item, index) =>
    cells.push(
      `<c r="${spreadsheetColumn(index + 2)}1" t="inlineStr"><is>${textElement("t", item.name ?? `Series ${index + 1}`)}</is></c>`,
    ),
  );
  const rowXml = [`<row r="1">${cells.join("")}</row>`];
  for (let row = 0; row < rows; row += 1) {
    const rowCells = [
      `<c r="A${row + 2}" t="inlineStr"><is>${textElement("t", series[0].categories[row])}</is></c>`,
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
  const rawSlide = source.packageGraph.rawParts?.find((part) => part.partPath === slidePartPath);
  if (rawSlide?.kind === "binary") {
    const root = parseXml(decoder.decode(rawSlide.bytes));
    const rootId = getAttr(
      getChild(
        getChild(getChild(getChild(getChild(root, "sld"), "cSld"), "spTree"), "nvGrpSpPr"),
        "cNvPr",
      ),
      "id",
    );
    if (rootId !== undefined) used.add(rootId);
  }
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
  if (!CHART_TYPES.has(input.chartType)) throw new Error("addChart: unsupported chartType");
  if (input.legendPosition !== undefined && !LEGEND_POSITIONS.has(input.legendPosition))
    throw new Error("addChart: unsupported legendPosition");
  if (input.radarStyle !== undefined && !RADAR_STYLES.has(input.radarStyle))
    throw new Error("addChart: unsupported radarStyle");
  if (input.displayBlanksAs !== undefined && !BLANK_DISPLAY_MODES.has(input.displayBlanksAs))
    throw new Error("addChart: unsupported displayBlanksAs");
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
  if (input.series.length > XLSX_MAX_SERIES)
    throw new Error(`addChart: series must not exceed ${XLSX_MAX_SERIES}`);
  const count = input.series[0].categories.length;
  if (count === 0) throw new Error("addChart: series categories must not be empty");
  if (count > XLSX_MAX_DATA_POINTS)
    throw new Error(`addChart: data points must not exceed ${XLSX_MAX_DATA_POINTS}`);
  for (const text of [input.name, input.title, input.categoryAxis?.title, input.valueAxis?.title]) {
    if (text !== undefined) assertValidXmlText(text);
  }
  if (input.titleStyle !== undefined && input.title === undefined)
    throw new Error("addChart: titleStyle requires title");
  assertTextStyle(input.titleStyle, "titleStyle");
  assertAreaStyle(input.chartArea, "chartArea");
  assertAreaStyle(input.plotArea, "plotArea");
  assertAxis(input.categoryAxis, "categoryAxis");
  assertAxis(input.valueAxis, "valueAxis");
  input.series.forEach((series, seriesIndex) => {
    if (series.categories.length !== count || series.values.length !== count)
      throw new Error("addChart: every series must have matching category and value counts");
    if (series.categories.some((category, index) => category !== input.series[0].categories[index]))
      throw new Error("addChart: every series must use identical category labels");
    if (series.values.some((value) => !Number.isFinite(value)))
      throw new Error("addChart: values must be finite numbers");
    if (series.name !== undefined) assertValidXmlText(series.name);
    series.categories.forEach(assertValidXmlText);
    if (series.color !== undefined) normalizeColor(series.color);
    assertFill(series.fill, `series[${seriesIndex}].fill`);
    assertOutline(series.outline, `series[${seriesIndex}].outline`);
    if (series.marker !== undefined) {
      if (input.chartType !== "line" && input.chartType !== "radar")
        throw new Error("addChart: marker is only valid for line and radar charts");
      if (series.marker.symbol !== undefined && !MARKER_SYMBOLS.has(series.marker.symbol))
        throw new Error(`addChart: unsupported series[${seriesIndex}].marker.symbol`);
      if (
        series.marker.size !== undefined &&
        (!Number.isInteger(series.marker.size) || series.marker.size < 2 || series.marker.size > 72)
      )
        throw new Error(`addChart: series[${seriesIndex}].marker.size must be from 2 through 72`);
      assertFill(series.marker.fill, `series[${seriesIndex}].marker.fill`);
      assertOutline(series.marker.outline, `series[${seriesIndex}].marker.outline`);
    }
    if (
      series.dataPoints !== undefined &&
      input.chartType !== "pie" &&
      input.chartType !== "doughnut" &&
      !(input.chartType === "bar" && input.series.length === 1)
    )
      throw new Error(
        "addChart: dataPoints are only valid for pie, doughnut, and single-series bar charts",
      );
    const pointIndexes = new Set<number>();
    for (const point of series.dataPoints ?? []) {
      if (!Number.isInteger(point.index) || point.index < 0 || point.index >= count)
        throw new Error(`addChart: series[${seriesIndex}].dataPoints index is out of range`);
      if (pointIndexes.has(point.index))
        throw new Error(`addChart: series[${seriesIndex}].dataPoints indexes must be unique`);
      pointIndexes.add(point.index);
      assertFill(point.fill, `series[${seriesIndex}].dataPoints.fill`);
      assertOutline(point.outline, `series[${seriesIndex}].dataPoints.outline`);
    }
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
  for (const value of [
    input.plotLayout?.x,
    input.plotLayout?.y,
    input.plotLayout?.width,
    input.plotLayout?.height,
  ])
    if (value !== undefined && (!Number.isFinite(value) || value < 0 || value > 1))
      throw new Error("addChart: plotLayout values must be finite numbers from 0 through 1");
  if (
    input.plotLayout?.coordinateMode !== undefined &&
    input.plotLayout.coordinateMode !== "factor" &&
    input.plotLayout.coordinateMode !== "edge"
  )
    throw new Error("addChart: unsupported plotLayout coordinateMode");
}

function assertAxis(axis: AddChartAxisInput | undefined, field: string): void {
  if (axis === undefined) return;
  if (axis.majorTickMark !== undefined && !TICK_MARKS.has(axis.majorTickMark))
    throw new Error(`addChart: unsupported ${field}.majorTickMark`);
  if (axis.labelPosition !== undefined && !LABEL_POSITIONS.has(axis.labelPosition))
    throw new Error(`addChart: unsupported ${field}.labelPosition`);
  if (axis.numberFormat !== undefined) {
    assertValidXmlText(axis.numberFormat.formatCode);
    if (axis.numberFormat.formatCode.length === 0)
      throw new Error(`addChart: ${field}.numberFormat.formatCode must not be empty`);
  }
  assertOutline(axis.line, `${field}.line`);
  assertOutline(axis.majorGridline, `${field}.majorGridline`);
  assertTextStyle(axis.textStyle, `${field}.textStyle`);
}

function assertAreaStyle(style: AddChartAreaStyleInput | undefined, field: string): void {
  if (style === undefined) return;
  assertFill(style.fill, `${field}.fill`);
  assertOutline(style.outline, `${field}.outline`);
}

function assertFill(fill: AddChartFillInput | undefined, field: string): void {
  if (fill === undefined || fill.kind === "none") return;
  if (fill.kind === "solid") {
    assertColor(fill.color, `${field}.color`);
    return;
  }
  throw new Error(`addChart: unsupported ${field}.kind`);
}

function assertOutline(outline: AddChartOutlineInput | undefined, field: string): void {
  if (outline === undefined) return;
  if (
    outline.width !== undefined &&
    (!Number.isFinite(outline.width) || outline.width < 0 || outline.width > MAX_LINE_WIDTH)
  )
    throw new Error(
      `addChart: ${field}.width must be a finite EMU value from 0 through ${MAX_LINE_WIDTH}`,
    );
  if (outline.dash !== undefined && !DASH_STYLES.has(outline.dash))
    throw new Error(`addChart: unsupported ${field}.dash`);
  assertFill(outline.fill, `${field}.fill`);
}

function assertTextStyle(style: AddChartTextStyleInput | undefined, field: string): void {
  if (style === undefined) return;
  if (style.fontFace !== undefined) {
    assertValidXmlText(style.fontFace);
    if (style.fontFace.length === 0)
      throw new Error(`addChart: ${field}.fontFace must not be empty`);
  }
  if (style.fontSize !== undefined) {
    const textSize = Math.round(style.fontSize * 100);
    if (!Number.isFinite(style.fontSize) || textSize < MIN_TEXT_SIZE || textSize > MAX_TEXT_SIZE)
      throw new Error(`addChart: ${field}.fontSize must be from 1 through 4000 points`);
  }
  if (style.color !== undefined) assertColor(style.color, `${field}.color`);
}

function assertColor(color: AddChartColorInput, field: string): void {
  if (color.kind !== "srgb") throw new Error(`addChart: unsupported ${field}.kind`);
  normalizeColor(color.hex);
  for (const transform of color.transforms ?? []) {
    if (
      transform.kind !== "alpha" ||
      !Number.isInteger(transform.value) ||
      transform.value < 0 ||
      transform.value > 100000
    )
      throw new Error(`addChart: ${field} alpha must be an integer from 0 through 100000`);
  }
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

function textElement(tag: string, value: string): string {
  const preserve = value.startsWith(" ") || value.endsWith(" ") ? ` xml:space="preserve"` : "";
  return `<${tag}${preserve}>${escapeXml(value)}</${tag}>`;
}

function assertValidXmlText(value: string): void {
  for (let index = 0; index < value.length; index += 1) {
    const codePoint = value.codePointAt(index);
    if (codePoint === undefined) continue;
    const valid =
      codePoint === 0x9 ||
      codePoint === 0xa ||
      codePoint === 0xd ||
      (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
      (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
      (codePoint >= 0x10000 && codePoint <= 0x10ffff);
    if (!valid) throw new Error("addChart: text contains a character forbidden by XML 1.0");
    if (codePoint > 0xffff) index += 1;
  }
}
