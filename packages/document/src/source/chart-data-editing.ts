import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { unzipSync, zipSync } from "fflate";

import {
  getAttr,
  getChild,
  getChildArray,
  getChildText,
  getNamespacedAttr,
  localName,
  parseXml,
} from "../reader/xml.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import { copyBytes, requireRawBinaryPart } from "./editing-shared.js";
import type {
  PartPath,
  PptxSourceModel,
  PptxSourceModelUpdateBubbleChartDataEdit,
  PptxSourceModelUpdateChartDataEdit,
  PptxSourceModelUpdateScatterChartDataEdit,
  Relationship,
  SourceChart,
  SourceHandle,
} from "./index.js";
import { resolveInternalRelationshipTarget } from "./package-paths.js";
import { findShapeNodeBySourceHandle } from "./shape-editing.js";

export interface UpdateChartSeriesDataInput {
  readonly name: string;
  readonly categories: readonly string[];
  readonly values: readonly number[];
  /** Required for fixed-topology combo charts; identifies the existing OOXML series. */
  readonly source?: UpdateCategoryChartSeriesSource;
}

/** Stable identity of an existing series in a supported category combo chart. */
export interface UpdateCategoryChartSeriesSource {
  readonly chartType: "bar" | "line";
  /** Existing unsigned `c:ser/c:idx@val`, scoped by `chartType`. */
  readonly index: number;
}

export interface UpdateChartDataInput {
  readonly series: readonly UpdateChartSeriesDataInput[];
}

export interface UpdateScatterChartSeriesDataInput {
  readonly name: string;
  readonly xValues: readonly number[];
  readonly yValues: readonly number[];
}

export interface UpdateScatterChartDataInput {
  readonly series: readonly UpdateScatterChartSeriesDataInput[];
}

export interface UpdateBubbleChartSeriesDataInput extends UpdateScatterChartSeriesDataInput {
  readonly bubbleSizes: readonly number[];
}

export interface UpdateBubbleChartDataInput {
  readonly series: readonly UpdateBubbleChartSeriesDataInput[];
}

type OrderedXmlNode = Record<string, unknown>;

const CHART_REL_TYPES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/chart",
]);
const PACKAGE_REL_TYPES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/package",
]);
const SUPPORTED_CHART_ELEMENTS: ReadonlySet<string> = new Set([
  "barChart",
  "lineChart",
  "pieChart",
  "areaChart",
  "doughnutChart",
  "radarChart",
]);
const XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const CHART_2014_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2014/chart";
const SERIES_UNIQUE_ID_EXTENSION_URI = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}";
const XLSX_MAX_SERIES = 16_383;
const XLSX_MAX_DATA_POINTS = 1_048_575;
const encoder = new TextEncoder();
const decoder = new TextDecoder();
const orderedParser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: false,
  trimValues: false,
});
const orderedBuilder = new XMLBuilder({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: true,
});

/**
 * Replaces the data bound to an existing category chart while preserving all non-data chart XML.
 *
 * The supported layout is deliberately narrow: one embedded workbook, one worksheet, names in row
 * 1, categories in column A, and values in columns B onward. Retained series preserve their XML;
 * appended series clone the last existing series as their formatting template.
 */
export function updateChartData(
  source: PptxSourceModel,
  handle: SourceHandle,
  input: UpdateChartDataInput,
): PptxSourceModel {
  assertUpdateInput(input);
  const chart = requireChart(source, handle);
  const chartPartPath = requireChartPartPath(source, chart);
  const chartRawPart = requireRawBinaryPart(source, chartPartPath, "updateChartData");
  const chartRelationships = requireRelationshipGroup(source, chartPartPath);
  const chartXml = decoder.decode(chartRawPart.bytes);
  const orderedRoot = parseOrderedXml(chartXml);
  const binding = inspectAndUpdateChartXml(orderedRoot, input);
  const workbookPartPath = requireWorkbookPartPath(
    source,
    chartRelationships.relationships,
    binding.externalDataRelationshipId,
    chartPartPath,
  );
  const workbookRawPart = requireRawBinaryPart(source, workbookPartPath, "updateChartData");
  assertWorkbookIsNotShared(source, chartPartPath, workbookPartPath);
  const workbook = inspectSupportedEmbeddedWorkbook(
    workbookRawPart.bytes,
    binding.sheetName,
    binding.existingSeriesCount,
    binding.pointCount,
  );

  const updatedChartBytes = encoder.encode(orderedBuilder.build(orderedRoot));
  const updatedWorkbookBytes = updateEmbeddedWorkbook(workbook, { series: binding.series });
  const edit = {
    kind: "updateChartData",
    handle,
    chartPartPath,
    workbookPartPath,
  } satisfies PptxSourceModelUpdateChartDataEdit;

  return {
    ...source,
    packageGraph: {
      ...source.packageGraph,
      rawParts: source.packageGraph.rawParts?.map((part) => {
        if (part.partPath === chartPartPath) {
          if (part.kind !== "binary") {
            throw new Error("updateChartData: chart part is not backed by binary material");
          }
          return { ...part, bytes: copyBytes(updatedChartBytes) };
        }
        if (part.partPath === workbookPartPath) {
          if (part.kind !== "binary") {
            throw new Error("updateChartData: workbook part is not backed by binary material");
          }
          return { ...part, bytes: copyBytes(updatedWorkbookBytes) };
        }
        return part;
      }),
    },
    edits: [...(source.edits ?? []), edit],
  };
}

/**
 * Replaces the data bound to an existing scatter chart while preserving all non-data chart XML.
 *
 * Each series occupies its own two-column table in the embedded worksheet, with X values in
 * column A, Y values in column B, the series name in the first column-B cell, and one empty row
 * between tables. Retained series preserve their XML; appended series clone the last existing
 * series as their formatting template.
 */
export function updateScatterChartData(
  source: PptxSourceModel,
  handle: SourceHandle,
  input: UpdateScatterChartDataInput,
): PptxSourceModel {
  assertScatterUpdateInput(input);
  try {
    const chart = requireChart(source, handle);
    const chartPartPath = requireChartPartPath(source, chart);
    const chartRawPart = requireRawBinaryPart(source, chartPartPath, "updateScatterChartData");
    const chartRelationships = requireRelationshipGroup(source, chartPartPath);
    const chartXml = decoder.decode(chartRawPart.bytes);
    const orderedRoot = parseOrderedXml(chartXml);
    const binding = inspectAndUpdateScatterChartXml(orderedRoot, input);
    const workbookPartPath = requireWorkbookPartPath(
      source,
      chartRelationships.relationships,
      binding.externalDataRelationshipId,
      chartPartPath,
    );
    const workbookRawPart = requireRawBinaryPart(
      source,
      workbookPartPath,
      "updateScatterChartData",
    );
    assertWorkbookIsNotShared(source, chartPartPath, workbookPartPath);
    const workbook = inspectSupportedEmbeddedScatterWorkbook(
      workbookRawPart.bytes,
      binding.sheetName,
      binding.existingPointCounts,
    );

    const updatedChartBytes = encoder.encode(orderedBuilder.build(orderedRoot));
    const updatedWorkbookBytes = updateEmbeddedScatterWorkbook(workbook, input);
    const edit = {
      kind: "updateScatterChartData",
      handle,
      chartPartPath,
      workbookPartPath,
    } satisfies PptxSourceModelUpdateScatterChartDataEdit;

    return {
      ...source,
      packageGraph: {
        ...source.packageGraph,
        rawParts: source.packageGraph.rawParts?.map((part) => {
          if (part.partPath === chartPartPath) {
            if (part.kind !== "binary") {
              throw new Error(
                "updateScatterChartData: chart part is not backed by binary material",
              );
            }
            return { ...part, bytes: copyBytes(updatedChartBytes) };
          }
          if (part.partPath === workbookPartPath) {
            if (part.kind !== "binary") {
              throw new Error(
                "updateScatterChartData: workbook part is not backed by binary material",
              );
            }
            return { ...part, bytes: copyBytes(updatedWorkbookBytes) };
          }
          return part;
        }),
      },
      edits: [...(source.edits ?? []), edit],
    };
  } catch (cause) {
    throw remapChartDataError(cause, "updateScatterChartData");
  }
}

/**
 * Replaces the data bound to an existing bubble chart while preserving all non-data chart XML.
 *
 * Each series occupies a standard three-column table: X in column A, Y/name in column B, bubble
 * size in column C with a `Size` header, and one empty row between series tables.
 */
export function updateBubbleChartData(
  source: PptxSourceModel,
  handle: SourceHandle,
  input: UpdateBubbleChartDataInput,
): PptxSourceModel {
  assertBubbleUpdateInput(input);
  try {
    const chart = requireChart(source, handle);
    const chartPartPath = requireChartPartPath(source, chart);
    const chartRawPart = requireRawBinaryPart(source, chartPartPath, "updateBubbleChartData");
    const chartRelationships = requireRelationshipGroup(source, chartPartPath);
    const orderedRoot = parseOrderedXml(decoder.decode(chartRawPart.bytes));
    const binding = inspectAndUpdateBubbleChartXml(orderedRoot, input);
    const workbookPartPath = requireWorkbookPartPath(
      source,
      chartRelationships.relationships,
      binding.externalDataRelationshipId,
      chartPartPath,
    );
    const workbookRawPart = requireRawBinaryPart(source, workbookPartPath, "updateBubbleChartData");
    assertWorkbookIsNotShared(source, chartPartPath, workbookPartPath);
    const workbook = inspectSupportedEmbeddedBubbleWorkbook(
      workbookRawPart.bytes,
      binding.sheetName,
      binding.existingPointCounts,
    );
    const updatedChartBytes = encoder.encode(orderedBuilder.build(orderedRoot));
    const updatedWorkbookBytes = updateEmbeddedBubbleWorkbook(workbook, input);
    const edit = {
      kind: "updateBubbleChartData",
      handle,
      chartPartPath,
      workbookPartPath,
    } satisfies PptxSourceModelUpdateBubbleChartDataEdit;

    return {
      ...source,
      packageGraph: {
        ...source.packageGraph,
        rawParts: source.packageGraph.rawParts?.map((part) => {
          if (part.partPath === chartPartPath) {
            if (part.kind !== "binary") {
              throw new Error("updateBubbleChartData: chart part is not backed by binary material");
            }
            return { ...part, bytes: copyBytes(updatedChartBytes) };
          }
          if (part.partPath === workbookPartPath) {
            if (part.kind !== "binary") {
              throw new Error(
                "updateBubbleChartData: workbook part is not backed by binary material",
              );
            }
            return { ...part, bytes: copyBytes(updatedWorkbookBytes) };
          }
          return part;
        }),
      },
      edits: [...(source.edits ?? []), edit],
    };
  } catch (cause) {
    throw remapChartDataError(cause, "updateBubbleChartData");
  }
}

function assertScatterUpdateInput(input: UpdateScatterChartDataInput): void {
  if (input.series.length === 0) {
    throw new Error("updateScatterChartData: series must not be empty");
  }
  let nextHeaderRow = 1;
  for (const [seriesIndex, series] of input.series.entries()) {
    if (typeof series.name !== "string") {
      throw new Error(`updateScatterChartData: series[${seriesIndex}].name must be a string`);
    }
    assertScatterXmlText(series.name, `series[${seriesIndex}].name`);
    if (series.xValues.length === 0 || series.xValues.length !== series.yValues.length) {
      throw new Error(
        "updateScatterChartData: every series must have matching non-empty X and Y value counts",
      );
    }
    if (!series.xValues.every(Number.isFinite) || !series.yValues.every(Number.isFinite)) {
      throw new Error("updateScatterChartData: X and Y values must be finite numbers");
    }
    const lastDataRow = nextHeaderRow + series.xValues.length;
    if (lastDataRow > XLSX_MAX_DATA_POINTS + 1) {
      throw new Error(
        `updateScatterChartData: worksheet rows must not exceed ${XLSX_MAX_DATA_POINTS + 1}`,
      );
    }
    nextHeaderRow = lastDataRow + 2;
  }
}

function assertBubbleUpdateInput(input: UpdateBubbleChartDataInput): void {
  if (input.series.length === 0) {
    throw new Error("updateBubbleChartData: series must not be empty");
  }
  let nextHeaderRow = 1;
  for (const [seriesIndex, series] of input.series.entries()) {
    if (typeof series.name !== "string") {
      throw new Error(`updateBubbleChartData: series[${seriesIndex}].name must be a string`);
    }
    try {
      assertXmlText(series.name, `series[${seriesIndex}].name`);
    } catch (cause) {
      throw remapChartDataError(cause, "updateBubbleChartData");
    }
    if (
      series.xValues.length === 0 ||
      series.xValues.length !== series.yValues.length ||
      series.xValues.length !== series.bubbleSizes.length
    ) {
      throw new Error(
        "updateBubbleChartData: every series must have matching non-empty X, Y, and bubble size counts",
      );
    }
    if (
      !series.xValues.every(Number.isFinite) ||
      !series.yValues.every(Number.isFinite) ||
      !series.bubbleSizes.every(Number.isFinite)
    ) {
      throw new Error("updateBubbleChartData: X, Y, and bubble size values must be finite numbers");
    }
    const lastDataRow = nextHeaderRow + series.xValues.length;
    if (lastDataRow > XLSX_MAX_DATA_POINTS + 1) {
      throw new Error(
        `updateBubbleChartData: worksheet rows must not exceed ${XLSX_MAX_DATA_POINTS + 1}`,
      );
    }
    nextHeaderRow = lastDataRow + 2;
  }
}

function assertScatterXmlText(value: string, field: string): void {
  try {
    assertXmlText(value, field);
  } catch (cause) {
    throw remapChartDataError(cause, "updateScatterChartData");
  }
}

function remapChartDataError(cause: unknown, operation: string): unknown {
  if (!(cause instanceof Error) || !cause.message.startsWith("updateChartData:")) return cause;
  return new Error(`${operation}:${cause.message.slice("updateChartData:".length)}`, { cause });
}

function assertUpdateInput(input: UpdateChartDataInput): void {
  if (input.series.length === 0) {
    throw new Error("updateChartData: series must not be empty");
  }
  if (input.series.length > XLSX_MAX_SERIES) {
    throw new Error(`updateChartData: series must not exceed ${XLSX_MAX_SERIES}`);
  }
  const categories = input.series[0]?.categories;
  if (categories === undefined || categories.length === 0) {
    throw new Error("updateChartData: series categories must not be empty");
  }
  if (categories.length > XLSX_MAX_DATA_POINTS) {
    throw new Error(`updateChartData: data points must not exceed ${XLSX_MAX_DATA_POINTS}`);
  }
  for (const [seriesIndex, series] of input.series.entries()) {
    if (typeof series.name !== "string") {
      throw new Error(`updateChartData: series[${seriesIndex}].name must be a string`);
    }
    assertXmlText(series.name, `series[${seriesIndex}].name`);
    if (
      series.categories.length !== categories.length ||
      series.values.length !== categories.length
    ) {
      throw new Error("updateChartData: every series must have matching category and value counts");
    }
    if (!series.categories.every((category, index) => category === categories[index])) {
      throw new Error("updateChartData: every series must use identical category labels");
    }
    for (const [categoryIndex, category] of series.categories.entries()) {
      if (typeof category !== "string") {
        throw new Error(
          `updateChartData: series[${seriesIndex}].categories[${categoryIndex}] must be a string`,
        );
      }
      assertXmlText(category, `series[${seriesIndex}].categories[${categoryIndex}]`);
    }
    if (!series.values.every(Number.isFinite)) {
      throw new Error("updateChartData: values must be finite numbers");
    }
  }
}

function assertXmlText(value: string, field: string): void {
  for (let index = 0; index < value.length; index += 1) {
    const codePoint = value.codePointAt(index);
    if (codePoint === undefined) continue;
    if (codePoint > 0xffff) index += 1;
    const valid =
      codePoint === 0x9 ||
      codePoint === 0xa ||
      codePoint === 0xd ||
      (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
      (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
      (codePoint >= 0x10000 && codePoint <= 0x10ffff);
    if (!valid) throw new Error(`updateChartData: ${field} contains a forbidden XML character`);
  }
}

function requireChart(source: PptxSourceModel, handle: SourceHandle): SourceChart {
  const shape = findShapeNodeBySourceHandle(source, handle);
  if (shape === undefined) {
    throw new Error("updateChartData: chart handle was not found in PptxSourceModel source");
  }
  if (shape.kind !== "chart") {
    throw new Error("updateChartData: shape handle does not reference a chart");
  }
  if (shape.chartRelationshipId === undefined) {
    throw new Error("updateChartData: chart has no chart relationship id");
  }
  return shape;
}

function requireChartPartPath(source: PptxSourceModel, chart: SourceChart): PartPath {
  const ownerPartPath = chart.handle?.partPath;
  if (ownerPartPath === undefined || chart.chartRelationshipId === undefined) {
    throw new Error("updateChartData: chart source relationship cannot be resolved");
  }
  const relationship = requireRelationshipGroup(source, ownerPartPath).relationships.find(
    (candidate) =>
      candidate.id === chart.chartRelationshipId && CHART_REL_TYPES.has(candidate.type),
  );
  if (relationship === undefined) {
    throw new Error("updateChartData: chart relationship was not found");
  }
  const chartPartPath = resolveInternalRelationshipTarget(ownerPartPath, relationship);
  if (chartPartPath === undefined) {
    throw new Error("updateChartData: external chart relationships are not supported");
  }
  return chartPartPath;
}

function requireRelationshipGroup(source: PptxSourceModel, partPath: PartPath) {
  const group = source.packageGraph.relationships.find(
    (candidate) => candidate.sourcePartPath === partPath,
  );
  if (group === undefined) {
    throw new Error(`updateChartData: relationships for part '${partPath}' were not found`);
  }
  return group;
}

function requireWorkbookPartPath(
  source: PptxSourceModel,
  relationships: readonly Relationship[],
  relationshipId: string,
  chartPartPath: PartPath,
): PartPath {
  const relationship = relationships.find(
    (candidate) => candidate.id === relationshipId && PACKAGE_REL_TYPES.has(candidate.type),
  );
  if (relationship === undefined) {
    throw new Error("updateChartData: embedded workbook relationship was not found");
  }
  const workbookPartPath = resolveInternalRelationshipTarget(chartPartPath, relationship);
  if (workbookPartPath === undefined) {
    throw new Error("updateChartData: external workbook data is not supported");
  }
  if (!source.packageGraph.parts.some((part) => part.partPath === workbookPartPath)) {
    throw new Error("updateChartData: embedded workbook part was not found");
  }
  const workbookPart = source.packageGraph.parts.find((part) => part.partPath === workbookPartPath);
  if (workbookPart?.contentType !== XLSX_CONTENT_TYPE) {
    throw new Error("updateChartData: embedded workbook content type is not supported");
  }
  return workbookPartPath;
}

function assertWorkbookIsNotShared(
  source: PptxSourceModel,
  chartPartPath: PartPath,
  workbookPartPath: PartPath,
): void {
  const referencingChartParts = source.packageGraph.relationships.filter((group) =>
    group.relationships.some(
      (relationship) =>
        PACKAGE_REL_TYPES.has(relationship.type) &&
        resolveInternalRelationshipTarget(group.sourcePartPath, relationship) === workbookPartPath,
    ),
  );
  if (
    referencingChartParts.length !== 1 ||
    referencingChartParts[0]?.sourcePartPath !== chartPartPath
  ) {
    throw new Error("updateChartData: embedded workbook is shared by another package part");
  }
}

function parseOrderedXml(xml: string): OrderedXmlNode[] {
  try {
    return unsafeOoxmlBoundaryAssertion<OrderedXmlNode[]>(orderedParser.parse(xml));
  } catch (cause) {
    throw new Error("updateChartData: chart XML could not be parsed", { cause });
  }
}

function inspectAndUpdateChartXml(
  root: OrderedXmlNode[],
  input: UpdateChartDataInput,
): {
  readonly externalDataRelationshipId: string;
  readonly sheetName: string;
  readonly existingSeriesCount: number;
  readonly pointCount: number;
  readonly series: readonly UpdateChartSeriesDataInput[];
} {
  const chartSpace = requireSingleElement(root, "chartSpace", "chart XML has no chartSpace root");
  const chart = requireSingleChild(chartSpace, "chart");
  const plotArea = requireSingleChild(chart, "plotArea");
  const chartGroups = elementChildren(plotArea).filter((entry) =>
    elementLocalName(entry)?.endsWith("Chart"),
  );
  const chartGroupName =
    chartGroups[0] === undefined ? undefined : elementLocalName(chartGroups[0]);
  const isSingleCategoryChart =
    chartGroups.length === 1 &&
    chartGroupName !== undefined &&
    SUPPORTED_CHART_ELEMENTS.has(chartGroupName);
  const isSupportedCombo = isSupportedCategoryCombo(chartGroups);
  if (!isSingleCategoryChart && !isSupportedCombo) {
    throw new Error("updateChartData: chart type or combination is not supported");
  }
  const groupedSeries = chartGroups.map((group) => ({
    chartType: elementLocalName(group),
    group,
    series: elementChildren(group).filter((entry) => elementLocalName(entry) === "ser"),
  }));
  const existingSeries = groupedSeries.flatMap((group) => group.series);
  const comboInputs = isSupportedCombo
    ? matchComboSeriesInputs(groupedSeries, input.series)
    : input.series;

  let sheetName: string | undefined;
  const sheetTokens: string[] = [];
  let pointCount: number | undefined;
  for (const [index, seriesEntry] of existingSeries.entries()) {
    const tx = requireSingleChild(seriesEntry, "tx");
    const txRef = requireSingleChild(tx, "strRef");
    const category = requireSingleChild(seriesEntry, "cat");
    const categoryRef = requireSingleChild(category, "strRef");
    const value = requireSingleChild(seriesEntry, "val");
    const valueRef = requireSingleChild(value, "numRef");
    const txFormula = parseSingleCellFormula(requireFormula(txRef), index + 2, 1);
    const categoryFormula = parseRangeFormula(requireFormula(categoryRef), 1, 2);
    const valueFormula = parseRangeFormula(requireFormula(valueRef), index + 2, 2);
    if (
      categoryFormula.endColumn !== 1 ||
      valueFormula.endColumn !== index + 2 ||
      categoryFormula.endRow !== valueFormula.endRow ||
      txFormula.sheetName !== categoryFormula.sheetName ||
      txFormula.sheetName !== valueFormula.sheetName
    ) {
      throw new Error("updateChartData: chart formulas use an unsupported data layout");
    }
    if (sheetName !== undefined && sheetName !== txFormula.sheetName) {
      throw new Error("updateChartData: chart series refer to multiple worksheets");
    }
    const seriesPointCount = categoryFormula.endRow - 1;
    if (pointCount !== undefined && pointCount !== seriesPointCount) {
      throw new Error("updateChartData: chart series use inconsistent data ranges");
    }
    sheetName = txFormula.sheetName;
    sheetTokens.push(txFormula.sheetToken);
    pointCount = seriesPointCount;
  }

  if (sheetName === undefined || pointCount === undefined) {
    throw new Error("updateChartData: chart has no series");
  }
  if (
    !isSupportedCombo &&
    input.series.length < existingSeries.length &&
    hasExplicitLegendEntries(chart)
  ) {
    throw new Error(
      "updateChartData: removing series with explicit legend entries is not supported",
    );
  }

  const series = isSupportedCombo
    ? existingSeries
    : resizeChartSeries(chartSpace, chartGroups[0], existingSeries, input.series.length);
  const addedSeriesIdentities =
    input.series.length > existingSeries.length
      ? nextAddedSeriesIdentities(existingSeries, input.series.length - existingSeries.length)
      : [];
  for (const [index, seriesEntry] of series.entries()) {
    const sheetToken = sheetTokens[index] ?? sheetTokens.at(-1);
    if (sheetToken === undefined) throw new Error("updateChartData: chart has no series");
    const addedIdentity = addedSeriesIdentities[index - existingSeries.length];
    if (addedIdentity !== undefined) {
      setAttribute(requireSingleChild(seriesEntry, "idx"), "val", String(addedIdentity.index));
      setAttribute(requireSingleChild(seriesEntry, "order"), "val", String(addedIdentity.order));
    }
    const inputSeries = comboInputs[index];
    if (inputSeries === undefined)
      throw new Error("updateChartData: combo series identity is missing");
    const tx = requireSingleChild(seriesEntry, "tx");
    const txRef = requireSingleChild(tx, "strRef");
    const category = requireSingleChild(seriesEntry, "cat");
    const categoryRef = requireSingleChild(category, "strRef");
    const value = requireSingleChild(seriesEntry, "val");
    const valueRef = requireSingleChild(value, "numRef");

    setFormula(txRef, `${sheetToken}!$${spreadsheetColumn(index + 2)}$1`);
    setStringCache(requireSingleChild(txRef, "strCache"), [inputSeries.name]);
    setFormula(categoryRef, `${sheetToken}!$A$2:$A$${inputSeries.categories.length + 1}`);
    setStringCache(requireSingleChild(categoryRef, "strCache"), inputSeries.categories);
    setFormula(
      valueRef,
      `${sheetToken}!$${spreadsheetColumn(index + 2)}$2:$${spreadsheetColumn(index + 2)}$${inputSeries.values.length + 1}`,
    );
    setNumberCache(requireSingleChild(valueRef, "numCache"), inputSeries.values);
  }

  const externalData = requireSingleChild(chartSpace, "externalData");
  const externalDataRelationshipId = namespacedAttribute(externalData, "id");
  if (externalDataRelationshipId === undefined) {
    throw new Error("updateChartData: chart externalData has no relationship id");
  }
  return {
    externalDataRelationshipId,
    sheetName,
    existingSeriesCount: existingSeries.length,
    pointCount,
    series: comboInputs,
  };
}

function isSupportedCategoryCombo(chartGroups: readonly OrderedXmlNode[]): boolean {
  if (chartGroups.length !== 2) return false;
  const names = chartGroups.map(elementLocalName);
  const hasSupportedGroups =
    names.filter((name) => name === "barChart").length === 1 &&
    names.filter((name) => name === "lineChart").length === 1;
  if (!hasSupportedGroups) return false;
  const barGroup = chartGroups.find((group) => elementLocalName(group) === "barChart");
  const barDirection =
    barGroup === undefined
      ? undefined
      : elementChildren(barGroup).find((child) => elementLocalName(child) === "barDir");
  if (barDirection === undefined || attribute(barDirection, "val") !== "col") {
    throw new Error("updateChartData: combo charts require barDir=col");
  }
  return true;
}

function matchComboSeriesInputs(
  groupedSeries: readonly {
    readonly chartType: string | undefined;
    readonly series: readonly OrderedXmlNode[];
  }[],
  inputs: readonly UpdateChartSeriesDataInput[],
): readonly UpdateChartSeriesDataInput[] {
  const identities = groupedSeries.flatMap(({ chartType, series }) => {
    if (chartType !== "barChart" && chartType !== "lineChart") {
      throw new Error("updateChartData: chart type or combination is not supported");
    }
    const sourceChartType = chartType === "barChart" ? "bar" : "line";
    return series.map((entry) => {
      seriesIdentityValue(entry, "order");
      return {
        chartType: sourceChartType,
        index: seriesIdentityValue(entry, "idx"),
      } satisfies UpdateCategoryChartSeriesSource;
    });
  });
  if (identities.length === 0) throw new Error("updateChartData: chart has no series");
  const sourceIdentityKeys = identities.map(
    (identity) => `${identity.chartType}:${identity.index}`,
  );
  if (new Set(sourceIdentityKeys).size !== sourceIdentityKeys.length) {
    throw new Error("updateChartData: combo source series identity is duplicated");
  }
  if (inputs.length !== identities.length) {
    throw new Error("updateChartData: combo chart series count must remain unchanged");
  }
  const byIdentity = new Map<string, UpdateChartSeriesDataInput>();
  for (const input of inputs) {
    const source = input.source;
    if (
      source === undefined ||
      !Number.isSafeInteger(source.index) ||
      source.index < 0 ||
      source.index > 0xffff_ffff
    ) {
      throw new Error("updateChartData: combo series identity is missing or invalid");
    }
    const key = `${source.chartType}:${source.index}`;
    if (byIdentity.has(key))
      throw new Error("updateChartData: combo series identity is duplicated");
    byIdentity.set(key, input);
  }
  const matched = identities.map((identity) =>
    byIdentity.get(`${identity.chartType}:${identity.index}`),
  );
  if (matched.some((input) => input === undefined) || byIdentity.size !== identities.length) {
    throw new Error("updateChartData: combo series identity does not match the source topology");
  }
  return matched.map((input) => {
    if (input === undefined) throw new Error("updateChartData: combo series identity is missing");
    return input;
  });
}

function inspectAndUpdateScatterChartXml(
  root: OrderedXmlNode[],
  input: UpdateScatterChartDataInput,
): {
  readonly externalDataRelationshipId: string;
  readonly sheetName: string;
  readonly existingPointCounts: readonly number[];
} {
  const chartSpace = requireSingleElement(root, "chartSpace", "chart XML has no chartSpace root");
  const chart = requireSingleChild(chartSpace, "chart");
  const plotArea = requireSingleChild(chart, "plotArea");
  const chartGroups = elementChildren(plotArea).filter((entry) =>
    elementLocalName(entry)?.endsWith("Chart"),
  );
  if (chartGroups.length !== 1 || elementLocalName(chartGroups[0]) !== "scatterChart") {
    throw new Error("updateChartData: chart type or combination is not supported");
  }
  const existingSeries = elementChildren(chartGroups[0]).filter(
    (entry) => elementLocalName(entry) === "ser",
  );

  let sheetName: string | undefined;
  const sheetTokens: string[] = [];
  const existingPointCounts: number[] = [];
  let expectedHeaderRow = 1;
  for (const seriesEntry of existingSeries) {
    const txRef = requireSingleChild(requireSingleChild(seriesEntry, "tx"), "strRef");
    const xValueRef = requireSingleChild(requireSingleChild(seriesEntry, "xVal"), "numRef");
    const yValueRef = requireSingleChild(requireSingleChild(seriesEntry, "yVal"), "numRef");
    const txFormula = parseSingleCellFormula(requireFormula(txRef), 2, expectedHeaderRow);
    const xValueFormula = parseRangeFormula(requireFormula(xValueRef), 1, expectedHeaderRow + 1);
    const yValueFormula = parseRangeFormula(requireFormula(yValueRef), 2, expectedHeaderRow + 1);
    if (
      xValueFormula.endColumn !== 1 ||
      yValueFormula.endColumn !== 2 ||
      xValueFormula.endRow !== yValueFormula.endRow ||
      txFormula.sheetName !== xValueFormula.sheetName ||
      txFormula.sheetName !== yValueFormula.sheetName
    ) {
      throw new Error("updateChartData: chart formulas use an unsupported data layout");
    }
    if (sheetName !== undefined && sheetName !== txFormula.sheetName) {
      throw new Error("updateChartData: chart series refer to multiple worksheets");
    }
    const pointCount = xValueFormula.endRow - expectedHeaderRow;
    if (pointCount <= 0) {
      throw new Error("updateChartData: chart formulas use an unsupported data layout");
    }
    sheetName = txFormula.sheetName;
    sheetTokens.push(txFormula.sheetToken);
    existingPointCounts.push(pointCount);
    expectedHeaderRow = xValueFormula.endRow + 2;
  }

  if (sheetName === undefined || existingPointCounts.length === 0) {
    throw new Error("updateChartData: chart has no series");
  }
  if (input.series.length < existingSeries.length && hasExplicitLegendEntries(chart)) {
    throw new Error(
      "updateChartData: removing series with explicit legend entries is not supported",
    );
  }

  const series = resizeChartSeries(chartSpace, chartGroups[0], existingSeries, input.series.length);
  const addedSeriesIdentities =
    input.series.length > existingSeries.length
      ? nextAddedSeriesIdentities(existingSeries, input.series.length - existingSeries.length)
      : [];
  let headerRow = 1;
  for (const [index, seriesEntry] of series.entries()) {
    const sheetToken = sheetTokens[index] ?? sheetTokens.at(-1);
    if (sheetToken === undefined) throw new Error("updateChartData: chart has no series");
    const addedIdentity = addedSeriesIdentities[index - existingSeries.length];
    if (addedIdentity !== undefined) {
      setAttribute(requireSingleChild(seriesEntry, "idx"), "val", String(addedIdentity.index));
      setAttribute(requireSingleChild(seriesEntry, "order"), "val", String(addedIdentity.order));
    }
    const inputSeries = input.series[index];
    const lastDataRow = headerRow + inputSeries.xValues.length;
    const txRef = requireSingleChild(requireSingleChild(seriesEntry, "tx"), "strRef");
    const xValueRef = requireSingleChild(requireSingleChild(seriesEntry, "xVal"), "numRef");
    const yValueRef = requireSingleChild(requireSingleChild(seriesEntry, "yVal"), "numRef");

    setFormula(txRef, `${sheetToken}!$B$${headerRow}`);
    setStringCache(requireSingleChild(txRef, "strCache"), [inputSeries.name]);
    setFormula(xValueRef, `${sheetToken}!$A$${headerRow + 1}:$A$${lastDataRow}`);
    setNumberCache(requireSingleChild(xValueRef, "numCache"), inputSeries.xValues);
    setFormula(yValueRef, `${sheetToken}!$B$${headerRow + 1}:$B$${lastDataRow}`);
    setNumberCache(requireSingleChild(yValueRef, "numCache"), inputSeries.yValues);
    headerRow = lastDataRow + 2;
  }

  const externalData = requireSingleChild(chartSpace, "externalData");
  const externalDataRelationshipId = namespacedAttribute(externalData, "id");
  if (externalDataRelationshipId === undefined) {
    throw new Error("updateChartData: chart externalData has no relationship id");
  }
  return { externalDataRelationshipId, sheetName, existingPointCounts };
}

function inspectAndUpdateBubbleChartXml(
  root: OrderedXmlNode[],
  input: UpdateBubbleChartDataInput,
): {
  readonly externalDataRelationshipId: string;
  readonly sheetName: string;
  readonly existingPointCounts: readonly number[];
} {
  const chartSpace = requireSingleElement(root, "chartSpace", "chart XML has no chartSpace root");
  const chart = requireSingleChild(chartSpace, "chart");
  const plotArea = requireSingleChild(chart, "plotArea");
  const chartGroups = elementChildren(plotArea).filter((entry) =>
    elementLocalName(entry)?.endsWith("Chart"),
  );
  if (chartGroups.length !== 1 || elementLocalName(chartGroups[0]) !== "bubbleChart") {
    throw new Error("updateChartData: chart type or combination is not supported");
  }
  const existingSeries = elementChildren(chartGroups[0]).filter(
    (entry) => elementLocalName(entry) === "ser",
  );

  let sheetName: string | undefined;
  const sheetTokens: string[] = [];
  const existingPointCounts: number[] = [];
  let expectedHeaderRow = 1;
  for (const seriesEntry of existingSeries) {
    const txRef = requireSingleChild(requireSingleChild(seriesEntry, "tx"), "strRef");
    const xValueRef = requireSingleChild(requireSingleChild(seriesEntry, "xVal"), "numRef");
    const yValueRef = requireSingleChild(requireSingleChild(seriesEntry, "yVal"), "numRef");
    const sizeRef = requireSingleChild(requireSingleChild(seriesEntry, "bubbleSize"), "numRef");
    const txFormula = parseSingleCellFormula(requireFormula(txRef), 2, expectedHeaderRow);
    const xValueFormula = parseRangeFormula(requireFormula(xValueRef), 1, expectedHeaderRow + 1);
    const yValueFormula = parseRangeFormula(requireFormula(yValueRef), 2, expectedHeaderRow + 1);
    const sizeFormula = parseRangeFormula(requireFormula(sizeRef), 3, expectedHeaderRow + 1);
    if (
      xValueFormula.endColumn !== 1 ||
      yValueFormula.endColumn !== 2 ||
      sizeFormula.endColumn !== 3 ||
      xValueFormula.endRow !== yValueFormula.endRow ||
      xValueFormula.endRow !== sizeFormula.endRow ||
      txFormula.sheetName !== xValueFormula.sheetName ||
      txFormula.sheetName !== yValueFormula.sheetName ||
      txFormula.sheetName !== sizeFormula.sheetName
    ) {
      throw new Error("updateChartData: chart formulas use an unsupported data layout");
    }
    if (sheetName !== undefined && sheetName !== txFormula.sheetName) {
      throw new Error("updateChartData: chart series refer to multiple worksheets");
    }
    const pointCount = xValueFormula.endRow - expectedHeaderRow;
    if (pointCount <= 0) {
      throw new Error("updateChartData: chart formulas use an unsupported data layout");
    }
    sheetName = txFormula.sheetName;
    sheetTokens.push(txFormula.sheetToken);
    existingPointCounts.push(pointCount);
    expectedHeaderRow = xValueFormula.endRow + 2;
  }

  if (sheetName === undefined || existingPointCounts.length === 0) {
    throw new Error("updateChartData: chart has no series");
  }
  if (input.series.length < existingSeries.length && hasExplicitLegendEntries(chart)) {
    throw new Error(
      "updateChartData: removing series with explicit legend entries is not supported",
    );
  }

  const series = resizeChartSeries(chartSpace, chartGroups[0], existingSeries, input.series.length);
  const addedSeriesIdentities =
    input.series.length > existingSeries.length
      ? nextAddedSeriesIdentities(existingSeries, input.series.length - existingSeries.length)
      : [];
  let headerRow = 1;
  for (const [index, seriesEntry] of series.entries()) {
    const sheetToken = sheetTokens[index] ?? sheetTokens.at(-1);
    if (sheetToken === undefined) throw new Error("updateChartData: chart has no series");
    const addedIdentity = addedSeriesIdentities[index - existingSeries.length];
    if (addedIdentity !== undefined) {
      setAttribute(requireSingleChild(seriesEntry, "idx"), "val", String(addedIdentity.index));
      setAttribute(requireSingleChild(seriesEntry, "order"), "val", String(addedIdentity.order));
    }
    const inputSeries = input.series[index];
    const lastDataRow = headerRow + inputSeries.xValues.length;
    const txRef = requireSingleChild(requireSingleChild(seriesEntry, "tx"), "strRef");
    const xValueRef = requireSingleChild(requireSingleChild(seriesEntry, "xVal"), "numRef");
    const yValueRef = requireSingleChild(requireSingleChild(seriesEntry, "yVal"), "numRef");
    const sizeRef = requireSingleChild(requireSingleChild(seriesEntry, "bubbleSize"), "numRef");

    setFormula(txRef, `${sheetToken}!$B$${headerRow}`);
    setStringCache(requireSingleChild(txRef, "strCache"), [inputSeries.name]);
    setFormula(xValueRef, `${sheetToken}!$A$${headerRow + 1}:$A$${lastDataRow}`);
    setNumberCache(requireSingleChild(xValueRef, "numCache"), inputSeries.xValues);
    setFormula(yValueRef, `${sheetToken}!$B$${headerRow + 1}:$B$${lastDataRow}`);
    setNumberCache(requireSingleChild(yValueRef, "numCache"), inputSeries.yValues);
    setFormula(sizeRef, `${sheetToken}!$C$${headerRow + 1}:$C$${lastDataRow}`);
    setNumberCache(requireSingleChild(sizeRef, "numCache"), inputSeries.bubbleSizes);
    headerRow = lastDataRow + 2;
  }

  const externalData = requireSingleChild(chartSpace, "externalData");
  const externalDataRelationshipId = namespacedAttribute(externalData, "id");
  if (externalDataRelationshipId === undefined) {
    throw new Error("updateChartData: chart externalData has no relationship id");
  }
  return { externalDataRelationshipId, sheetName, existingPointCounts };
}

interface SeriesIdentity {
  readonly index: number;
  readonly order: number;
}

function nextAddedSeriesIdentities(
  existingSeries: readonly OrderedXmlNode[],
  addedCount: number,
): SeriesIdentity[] {
  const identities = existingSeries.map((series) => ({
    index: seriesIdentityValue(series, "idx"),
    order: seriesIdentityValue(series, "order"),
  }));
  if (
    new Set(identities.map((identity) => identity.index)).size !== identities.length ||
    new Set(identities.map((identity) => identity.order)).size !== identities.length
  ) {
    throw new Error("updateChartData: chart series idx/order values must be unique");
  }
  const maxIndex = Math.max(...identities.map((identity) => identity.index));
  const maxOrder = Math.max(...identities.map((identity) => identity.order));
  if (maxIndex + addedCount > 0xffff_ffff || maxOrder + addedCount > 0xffff_ffff) {
    throw new Error("updateChartData: chart series idx/order values cannot be extended");
  }
  return Array.from({ length: addedCount }, (_, offset) => ({
    index: maxIndex + offset + 1,
    order: maxOrder + offset + 1,
  }));
}

function seriesIdentityValue(series: OrderedXmlNode, localName: "idx" | "order"): number {
  const value = attribute(requireSingleChild(series, localName), "val");
  if (value === undefined || !/^\d+$/.test(value)) {
    throw new Error(`updateChartData: chart series ${localName} value is invalid`);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > 0xffff_ffff) {
    throw new Error(`updateChartData: chart series ${localName} value is invalid`);
  }
  return parsed;
}

function resizeChartSeries(
  chartSpace: OrderedXmlNode,
  chartGroup: OrderedXmlNode,
  existingSeries: readonly OrderedXmlNode[],
  desiredCount: number,
): OrderedXmlNode[] {
  const retainedSeries = existingSeries.slice(0, desiredCount);
  const children = mutableElementChildren(chartGroup);
  let seenSeries = 0;
  const retainedChildren = children.filter((entry) => {
    if (elementLocalName(entry) !== "ser") return true;
    const retain = seenSeries < desiredCount;
    seenSeries += 1;
    return retain;
  });
  if (desiredCount > existingSeries.length) {
    const template = existingSeries.at(-1);
    if (template === undefined) throw new Error("updateChartData: chart has no series");
    const rewriteUniqueId = createAddedSeriesUniqueIdRewriter(chartSpace, existingSeries, template);
    const addedSeries = Array.from({ length: desiredCount - existingSeries.length }, () => {
      const clone = structuredClone(template);
      rewriteUniqueId(clone);
      return clone;
    });
    let lastSeriesIndex = -1;
    for (const [index, entry] of retainedChildren.entries()) {
      if (elementLocalName(entry) === "ser") lastSeriesIndex = index;
    }
    retainedChildren.splice(lastSeriesIndex + 1, 0, ...addedSeries);
    retainedSeries.push(...addedSeries);
  }
  setElementChildren(chartGroup, retainedChildren);
  return retainedSeries;
}

function hasExplicitLegendEntries(chart: OrderedXmlNode): boolean {
  const legend = elementChildren(chart).find((entry) => elementLocalName(entry) === "legend");
  return (
    legend !== undefined &&
    elementChildren(legend).some((entry) => elementLocalName(entry) === "legendEntry")
  );
}

function createAddedSeriesUniqueIdRewriter(
  chartSpace: OrderedXmlNode,
  existingSeries: readonly OrderedXmlNode[],
  template: OrderedXmlNode,
): (clone: OrderedXmlNode) => void {
  const chartUniqueIds = inspectChartUniqueIds(chartSpace);
  const existingValues = existingSeries.flatMap((series) =>
    seriesIdentityUniqueIdElements(series, chartUniqueIds.namespaceUris),
  );
  const normalizedSeriesIds = existingValues.map(({ value }) => normalizeGuid(value));
  if (new Set(normalizedSeriesIds.map(({ hex }) => hex)).size !== normalizedSeriesIds.length) {
    throw new Error("updateChartData: chart series uniqueId values must be unique");
  }
  const used = new Set(chartUniqueIds.guids.map(({ hex }) => hex));
  const templateIds = seriesIdentityUniqueIdElements(template, chartUniqueIds.namespaceUris);
  if (templateIds.length === 0) return () => undefined;
  if (templateIds.length !== 1) {
    throw new Error("updateChartData: chart series uniqueId layout is not supported");
  }
  let previous = normalizeGuid(templateIds[0].value);
  return (clone) => {
    const cloneIds = seriesIdentityUniqueIdElements(clone);
    if (cloneIds.length !== 1) {
      throw new Error("updateChartData: chart series uniqueId layout is not supported");
    }
    previous = nextGuid(previous, used);
    used.add(previous.hex);
    setAttribute(cloneIds[0].element, "val", formatGuid(previous));
  };
}

interface SeriesUniqueIdElement {
  readonly element: OrderedXmlNode;
  readonly value: string;
}

function seriesIdentityUniqueIdElements(
  series: OrderedXmlNode,
  namespaceUris?: ReadonlyMap<OrderedXmlNode, string | undefined>,
): SeriesUniqueIdElement[] {
  const extensionLists = elementChildren(series).filter(
    (entry) => elementLocalName(entry) === "extLst",
  );
  if (extensionLists.length > 1) {
    throw new Error("updateChartData: chart series extension layout is not supported");
  }
  const extensionList = extensionLists[0];
  if (extensionList === undefined) return [];
  const identityExtensions = elementChildren(extensionList).filter(
    (entry) =>
      elementLocalName(entry) === "ext" &&
      attribute(entry, "uri")?.toUpperCase() === SERIES_UNIQUE_ID_EXTENSION_URI,
  );
  if (identityExtensions.length > 1) {
    throw new Error("updateChartData: chart series uniqueId layout is not supported");
  }
  const identityExtension = identityExtensions[0];
  if (identityExtension === undefined) return [];
  return elementChildren(identityExtension)
    .filter(
      (entry) =>
        elementLocalName(entry) === "uniqueId" &&
        (namespaceUris === undefined || namespaceUris.get(entry) === CHART_2014_NAMESPACE),
    )
    .map((element) => {
      const value = attribute(element, "val");
      if (value === undefined) {
        throw new Error("updateChartData: chart series uniqueId value is invalid");
      }
      return { element, value };
    });
}

function inspectChartUniqueIds(chartSpace: OrderedXmlNode): {
  readonly namespaceUris: ReadonlyMap<OrderedXmlNode, string | undefined>;
  readonly guids: readonly NormalizedGuid[];
} {
  const namespaceUris = new Map<OrderedXmlNode, string | undefined>();
  const guids: NormalizedGuid[] = [];
  visitElements(chartSpace, {}, (entry, namespaces) => {
    const namespaceUri = namespaceUriForElement(entry, namespaces);
    namespaceUris.set(entry, namespaceUri);
    if (elementLocalName(entry) !== "uniqueId" || namespaceUri !== CHART_2014_NAMESPACE) {
      return;
    }
    const value = attribute(entry, "val");
    if (value !== undefined && isGuid(value)) guids.push(normalizeGuid(value));
  });
  return { namespaceUris, guids };
}

function visitElements(
  entry: OrderedXmlNode,
  inheritedNamespaces: Readonly<Record<string, string>>,
  visit: (entry: OrderedXmlNode, namespaces: Readonly<Record<string, string>>) => void,
): void {
  const namespaces = { ...inheritedNamespaces };
  const attributes = entry[":@"];
  if (typeof attributes === "object" && attributes !== null) {
    for (const [key, value] of Object.entries(
      unsafeOoxmlBoundaryAssertion<Record<string, unknown>>(attributes),
    )) {
      if (typeof value !== "string") continue;
      const name = key.startsWith("@_") ? key.slice(2) : key;
      if (name === "xmlns") namespaces[""] = value;
      else if (name.startsWith("xmlns:")) namespaces[name.slice("xmlns:".length)] = value;
    }
  }
  visit(entry, namespaces);
  for (const child of elementChildren(entry)) visitElements(child, namespaces, visit);
}

function namespaceUriForElement(
  entry: OrderedXmlNode,
  namespaces: Readonly<Record<string, string>>,
): string | undefined {
  const name = elementName(entry);
  if (name === undefined) return undefined;
  const colon = name.indexOf(":");
  return namespaces[colon < 0 ? "" : name.slice(0, colon)];
}

function isGuid(value: string): boolean {
  const match = /^(\{?)[0-9A-Fa-f]{8}(?:-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}(\}?)$/.exec(value);
  return match !== null && (match[1] === "{") === (match[2] === "}");
}

interface NormalizedGuid {
  readonly hex: string;
  readonly braces: boolean;
  readonly lowercase: boolean;
}

function normalizeGuid(value: string): NormalizedGuid {
  const match =
    /^(\{?)([0-9A-Fa-f]{8})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{12})(\}?)$/.exec(
      value,
    );
  if (match === null || (match[1] === "{") !== (match[7] === "}")) {
    throw new Error("updateChartData: chart series uniqueId value is invalid");
  }
  const body = match.slice(2, 7).join("");
  return {
    hex: body.toUpperCase(),
    braces: match[1] === "{",
    lowercase: body === body.toLowerCase(),
  };
}

function nextGuid(previous: NormalizedGuid, used: ReadonlySet<string>): NormalizedGuid {
  const modulus = 1n << 128n;
  let value = (BigInt(`0x${previous.hex}`) + 1n) % modulus;
  while (used.has(value.toString(16).toUpperCase().padStart(32, "0"))) {
    value = (value + 1n) % modulus;
  }
  return { ...previous, hex: value.toString(16).toUpperCase().padStart(32, "0") };
}

function formatGuid(guid: NormalizedGuid): string {
  const body = [
    guid.hex.slice(0, 8),
    guid.hex.slice(8, 12),
    guid.hex.slice(12, 16),
    guid.hex.slice(16, 20),
    guid.hex.slice(20),
  ].join("-");
  const formatted = guid.lowercase ? body.toLowerCase() : body;
  return guid.braces ? `{${formatted}}` : formatted;
}

interface ParsedFormula {
  readonly sheetName: string;
  readonly sheetToken: string;
  readonly endColumn: number;
  readonly endRow: number;
}

function parseSingleCellFormula(formula: string, column: number, row: number): ParsedFormula {
  const parsed = parseRangeFormula(formula, column, row);
  if (parsed.endColumn !== column || parsed.endRow !== row) {
    throw new Error("updateChartData: chart formula uses an unsupported data layout");
  }
  return parsed;
}

function parseRangeFormula(formula: string, column: number, row: number): ParsedFormula {
  const match =
    /^((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*))!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?$/.exec(
      formula,
    );
  if (match === null) {
    throw new Error("updateChartData: chart formula uses an unsupported data layout");
  }
  const startColumn = spreadsheetColumnNumber(match[2]);
  const startRow = Number(match[3]);
  const endColumn = spreadsheetColumnNumber(match[4] ?? match[2]);
  const endRow = Number(match[5] ?? match[3]);
  if (startColumn !== column || startRow !== row || endColumn < startColumn || endRow < startRow) {
    throw new Error("updateChartData: chart formula uses an unsupported data layout");
  }
  const sheetToken = match[1];
  const sheetName = sheetToken.startsWith("'")
    ? sheetToken.slice(1, -1).replaceAll("''", "'")
    : sheetToken;
  return { sheetName, sheetToken, endColumn, endRow };
}

function requireFormula(reference: OrderedXmlNode): string {
  const formula = requireSingleChild(reference, "f");
  const text = elementText(formula);
  if (text === undefined) throw new Error("updateChartData: chart formula is empty");
  return text;
}

function setFormula(reference: OrderedXmlNode, value: string): void {
  const formula = requireSingleChild(reference, "f");
  replaceElementText(formula, value);
}

function setStringCache(cache: OrderedXmlNode, values: readonly string[]): void {
  replaceCachePoints(
    cache,
    values.map((value) => value),
  );
}

function setNumberCache(cache: OrderedXmlNode, values: readonly number[]): void {
  replaceCachePoints(cache, values.map(String));
}

function replaceCachePoints(cache: OrderedXmlNode, values: readonly string[]): void {
  const children = mutableElementChildren(cache);
  const pointCount = requireSingleElement(children, "ptCount", "chart cache has no ptCount");
  setAttribute(pointCount, "val", String(values.length));
  const firstPointIndex = children.findIndex((entry) => elementLocalName(entry) === "pt");
  const retained = children.filter((entry) => elementLocalName(entry) !== "pt");
  const insertionIndex =
    firstPointIndex < 0
      ? retained.findIndex((entry) => elementLocalName(entry) === "ptCount") + 1
      : children.slice(0, firstPointIndex).filter((entry) => elementLocalName(entry) !== "pt")
          .length;
  const prefix = elementPrefix(cache);
  const points = values.map((value, index) =>
    createElement(`${prefix}pt`, [createElement(`${prefix}v`, [{ "#text": value }])], {
      "@_idx": String(index),
    }),
  );
  setElementChildren(cache, [
    ...retained.slice(0, insertionIndex),
    ...points,
    ...retained.slice(insertionIndex),
  ]);
}

interface EditableEmbeddedWorkbook {
  readonly files: Record<string, Uint8Array>;
  readonly worksheetPath: string;
}

function inspectSupportedEmbeddedWorkbook(
  bytes: Uint8Array,
  sheetName: string,
  seriesCount: number,
  pointCount: number,
): EditableEmbeddedWorkbook {
  return inspectEmbeddedWorkbook(bytes, sheetName, (worksheetBytes) =>
    assertSupportedWorksheetLayout(worksheetBytes, seriesCount, pointCount),
  );
}

function inspectSupportedEmbeddedScatterWorkbook(
  bytes: Uint8Array,
  sheetName: string,
  pointCounts: readonly number[],
): EditableEmbeddedWorkbook {
  return inspectEmbeddedWorkbook(bytes, sheetName, (worksheetBytes) =>
    assertSupportedScatterWorksheetLayout(worksheetBytes, pointCounts),
  );
}

function inspectSupportedEmbeddedBubbleWorkbook(
  bytes: Uint8Array,
  sheetName: string,
  pointCounts: readonly number[],
): EditableEmbeddedWorkbook {
  return inspectEmbeddedWorkbook(bytes, sheetName, (worksheetBytes, files) =>
    assertSupportedBubbleWorksheetLayout(
      worksheetBytes,
      pointCounts,
      files["xl/sharedStrings.xml"],
    ),
  );
}

function inspectEmbeddedWorkbook(
  bytes: Uint8Array,
  sheetName: string,
  assertWorksheetLayout: (bytes: Uint8Array, files: Readonly<Record<string, Uint8Array>>) => void,
): EditableEmbeddedWorkbook {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes);
  } catch (cause) {
    throw new Error("updateChartData: embedded workbook is not a readable XLSX package", { cause });
  }
  if (
    Object.keys(files).some(
      (path) =>
        path === "xl/connections.xml" ||
        path.startsWith("xl/externalLinks/") ||
        path.startsWith("xl/queryTables/"),
    )
  ) {
    throw new Error("updateChartData: external workbook links are not supported");
  }
  const workbookBytes = files["xl/workbook.xml"];
  const workbookRelationshipsBytes = files["xl/_rels/workbook.xml.rels"];
  if (workbookBytes === undefined || workbookRelationshipsBytes === undefined) {
    throw new Error("updateChartData: embedded workbook metadata is incomplete");
  }
  const workbook = getChild(parseXml(decoder.decode(workbookBytes)), "workbook");
  if (getChild(workbook, "definedNames") !== undefined) {
    throw new Error("updateChartData: workbook defined names are not supported");
  }
  const sheets = getChildArray(getChild(workbook, "sheets"), "sheet");
  if (sheets.length !== 1 || getAttr(sheets[0], "name") !== sheetName) {
    throw new Error("updateChartData: embedded workbook must contain one matching worksheet");
  }
  const sheetRelationshipId = getNamespacedAttr(sheets[0], "id");
  const relationships = getChildArray(
    getChild(parseXml(decoder.decode(workbookRelationshipsBytes)), "Relationships"),
    "Relationship",
  );
  const worksheetRelationship = relationships.find(
    (relationship) =>
      getAttr(relationship, "Id") === sheetRelationshipId &&
      getAttr(relationship, "Type")?.endsWith("/worksheet") === true &&
      getAttr(relationship, "TargetMode") !== "External",
  );
  const target = getAttr(worksheetRelationship, "Target");
  const worksheetPath = target === undefined ? undefined : normalizeWorkbookTarget(target);
  const worksheetBytes = worksheetPath === undefined ? undefined : files[worksheetPath];
  if (worksheetPath === undefined || worksheetBytes === undefined) {
    throw new Error("updateChartData: embedded worksheet relationship cannot be resolved");
  }
  assertWorksheetLayout(worksheetBytes, files);
  return { files, worksheetPath };
}

function assertSupportedWorksheetLayout(
  bytes: Uint8Array,
  seriesCount: number,
  pointCount: number,
): void {
  const worksheet = getChild(parseXml(decoder.decode(bytes)), "worksheet");
  for (const unsupportedElement of [
    "autoFilter",
    "conditionalFormatting",
    "dataValidations",
    "drawing",
    "hyperlinks",
    "legacyDrawing",
    "mergeCells",
    "oleObjects",
    "tableParts",
  ]) {
    if (getChild(worksheet, unsupportedElement) !== undefined) {
      throw new Error(
        `updateChartData: worksheet ${unsupportedElement} data layout is not supported`,
      );
    }
  }
  const rows = getChildArray(getChild(worksheet, "sheetData"), "row");
  const rowNumbers = new Set<string>();
  for (const row of rows) {
    const rowNumber = getAttr(row, "r");
    if (
      rowNumber === undefined ||
      rowNumbers.has(rowNumber) ||
      !/^[1-9]\d*$/.test(rowNumber) ||
      Number(rowNumber) > pointCount + 1
    ) {
      throw new Error("updateChartData: embedded worksheet uses an unsupported data layout");
    }
    rowNumbers.add(rowNumber);
  }
  const cells = rows.flatMap((row) => getChildArray(row, "c"));
  assertSupportedWorksheetRowAndCellChildren(rows);
  const allowed = new Set<string>(["A1"]);
  for (let row = 1; row <= pointCount + 1; row += 1) {
    for (let column = 1; column <= seriesCount + 1; column += 1) {
      if (row !== 1 || column !== 1) allowed.add(`${spreadsheetColumn(column)}${row}`);
    }
  }
  const actual = new Set<string>();
  for (const cell of cells) {
    const reference = getAttr(cell, "r");
    if (reference === undefined || actual.has(reference) || !allowed.has(reference)) {
      throw new Error("updateChartData: embedded worksheet uses an unsupported data layout");
    }
    if (getChild(cell, "f") !== undefined) {
      throw new Error("updateChartData: formulas in the chart data range are not supported");
    }
    actual.add(reference);
  }
}

function assertSupportedScatterWorksheetLayout(
  bytes: Uint8Array,
  pointCounts: readonly number[],
): void {
  const worksheet = getChild(parseXml(decoder.decode(bytes)), "worksheet");
  assertWorksheetHasNoUnsupportedElements(worksheet);
  const allowed = new Set<string>();
  const allowedRows = new Set<string>();
  let headerRow = 1;
  for (const pointCount of pointCounts) {
    allowedRows.add(String(headerRow));
    allowed.add(`B${headerRow}`);
    for (let row = headerRow + 1; row <= headerRow + pointCount; row += 1) {
      allowedRows.add(String(row));
      allowed.add(`A${row}`);
      allowed.add(`B${row}`);
    }
    headerRow += pointCount + 2;
  }
  assertWorksheetRowsAndCells(worksheet, allowed, allowedRows, headerRow - 2);
}

function assertSupportedBubbleWorksheetLayout(
  bytes: Uint8Array,
  pointCounts: readonly number[],
  sharedStringsBytes: Uint8Array | undefined,
): void {
  const worksheet = getChild(parseXml(decoder.decode(bytes)), "worksheet");
  assertWorksheetHasNoUnsupportedElements(worksheet);
  const allowed = new Set<string>();
  const allowedRows = new Set<string>();
  let headerRow = 1;
  for (const pointCount of pointCounts) {
    allowedRows.add(String(headerRow));
    allowed.add(`B${headerRow}`);
    allowed.add(`C${headerRow}`);
    for (let row = headerRow + 1; row <= headerRow + pointCount; row += 1) {
      allowedRows.add(String(row));
      allowed.add(`A${row}`);
      allowed.add(`B${row}`);
      allowed.add(`C${row}`);
    }
    headerRow += pointCount + 2;
  }
  assertWorksheetRowsAndCells(worksheet, allowed, allowedRows, headerRow - 2);
  const cells = getChildArray(getChild(worksheet, "sheetData"), "row").flatMap((row) =>
    getChildArray(row, "c"),
  );
  let expectedHeaderRow = 1;
  for (const pointCount of pointCounts) {
    const sizeHeader = cells.find((cell) => getAttr(cell, "r") === `C${expectedHeaderRow}`);
    if (workbookStringCellValue(sizeHeader, sharedStringsBytes) !== "Size") {
      throw new Error("updateChartData: bubble size header must be 'Size'");
    }
    expectedHeaderRow += pointCount + 2;
  }
}

function workbookStringCellValue(
  cell: Record<string, unknown> | undefined,
  sharedStringsBytes: Uint8Array | undefined,
): string | undefined {
  if (getAttr(cell, "t") === "inlineStr") return getChildText(getChild(cell, "is"), "t");
  if (getAttr(cell, "t") !== "s" || sharedStringsBytes === undefined) return undefined;
  const indexText = getChildText(cell, "v");
  if (indexText === undefined || !/^\d+$/.test(indexText)) return undefined;
  const sharedStrings = getChild(parseXml(decoder.decode(sharedStringsBytes)), "sst");
  return getChildText(getChildArray(sharedStrings, "si")[Number(indexText)], "t");
}

function assertWorksheetHasNoUnsupportedElements(
  worksheet: Record<string, unknown> | undefined,
): void {
  for (const unsupportedElement of [
    "autoFilter",
    "conditionalFormatting",
    "dataValidations",
    "drawing",
    "hyperlinks",
    "legacyDrawing",
    "mergeCells",
    "oleObjects",
    "tableParts",
  ]) {
    if (getChild(worksheet, unsupportedElement) !== undefined) {
      throw new Error(
        `updateChartData: worksheet ${unsupportedElement} data layout is not supported`,
      );
    }
  }
}

function assertWorksheetRowsAndCells(
  worksheet: Record<string, unknown> | undefined,
  allowed: ReadonlySet<string>,
  allowedRows: ReadonlySet<string>,
  maximumRow: number,
): void {
  const rows = getChildArray(getChild(worksheet, "sheetData"), "row");
  const rowNumbers = new Set<string>();
  for (const row of rows) {
    const rowNumber = getAttr(row, "r");
    if (
      rowNumber === undefined ||
      rowNumbers.has(rowNumber) ||
      !allowedRows.has(rowNumber) ||
      !/^[1-9]\d*$/.test(rowNumber) ||
      Number(rowNumber) > maximumRow
    ) {
      throw new Error("updateChartData: embedded worksheet uses an unsupported data layout");
    }
    rowNumbers.add(rowNumber);
  }
  assertSupportedWorksheetRowAndCellChildren(rows);
  const actual = new Set<string>();
  for (const cell of rows.flatMap((row) => getChildArray(row, "c"))) {
    const reference = getAttr(cell, "r");
    if (reference === undefined || actual.has(reference) || !allowed.has(reference)) {
      throw new Error("updateChartData: embedded worksheet uses an unsupported data layout");
    }
    if (getChild(cell, "f") !== undefined) {
      throw new Error("updateChartData: formulas in the chart data range are not supported");
    }
    actual.add(reference);
  }
}

function assertSupportedWorksheetRowAndCellChildren(
  rows: readonly Record<string, unknown>[],
): void {
  for (const row of rows) {
    assertOnlySupportedWorksheetChildren(row, new Set(["c"]), "row");
    for (const cell of getChildArray(row, "c")) {
      if (getChild(cell, "f") !== undefined) continue;
      assertOnlySupportedWorksheetChildren(cell, new Set(["is", "v"]), "cell");
    }
  }
}

function assertOnlySupportedWorksheetChildren(
  node: Record<string, unknown>,
  allowed: ReadonlySet<string>,
  context: "row" | "cell",
): void {
  const hasUnsupportedChild = Object.keys(node).some(
    (key) => !key.startsWith("@_") && key !== "#text" && !allowed.has(localName(key)),
  );
  if (hasUnsupportedChild) {
    throw new Error(`updateChartData: embedded worksheet ${context} child XML is not supported`);
  }
}

function updateEmbeddedWorkbook(
  workbook: EditableEmbeddedWorkbook,
  input: UpdateChartDataInput,
): Uint8Array {
  const worksheetRoot = parseOrderedXml(decoder.decode(workbook.files[workbook.worksheetPath]));
  const worksheet = requireSingleElement(
    worksheetRoot,
    "worksheet",
    "embedded workbook has no worksheet root",
  );
  const dimension = requireSingleChild(worksheet, "dimension");
  setAttribute(
    dimension,
    "ref",
    `A1:${spreadsheetColumn(input.series.length + 1)}${input.series[0].categories.length + 1}`,
  );
  const sheetData = requireSingleChild(worksheet, "sheetData");
  setElementChildren(sheetData, buildWorksheetRows(input, sheetData));
  return zipSync({
    ...workbook.files,
    [workbook.worksheetPath]: encoder.encode(orderedBuilder.build(worksheetRoot)),
  });
}

function updateEmbeddedScatterWorkbook(
  workbook: EditableEmbeddedWorkbook,
  input: UpdateScatterChartDataInput,
): Uint8Array {
  const worksheetRoot = parseOrderedXml(decoder.decode(workbook.files[workbook.worksheetPath]));
  const worksheet = requireSingleElement(
    worksheetRoot,
    "worksheet",
    "embedded workbook has no worksheet root",
  );
  const lastRow = scatterLastDataRow(input.series.map((series) => series.xValues.length));
  setAttribute(requireSingleChild(worksheet, "dimension"), "ref", `A1:B${lastRow}`);
  const sheetData = requireSingleChild(worksheet, "sheetData");
  setElementChildren(sheetData, buildScatterWorksheetRows(input, sheetData));
  return zipSync({
    ...workbook.files,
    [workbook.worksheetPath]: encoder.encode(orderedBuilder.build(worksheetRoot)),
  });
}

function updateEmbeddedBubbleWorkbook(
  workbook: EditableEmbeddedWorkbook,
  input: UpdateBubbleChartDataInput,
): Uint8Array {
  const worksheetRoot = parseOrderedXml(decoder.decode(workbook.files[workbook.worksheetPath]));
  const worksheet = requireSingleElement(
    worksheetRoot,
    "worksheet",
    "embedded workbook has no worksheet root",
  );
  const lastRow = scatterLastDataRow(input.series.map((series) => series.xValues.length));
  setAttribute(requireSingleChild(worksheet, "dimension"), "ref", `A1:C${lastRow}`);
  const sheetData = requireSingleChild(worksheet, "sheetData");
  setElementChildren(sheetData, buildBubbleWorksheetRows(input, sheetData));
  return zipSync({
    ...workbook.files,
    [workbook.worksheetPath]: encoder.encode(orderedBuilder.build(worksheetRoot)),
  });
}

function buildScatterWorksheetRows(
  input: UpdateScatterChartDataInput,
  existingSheetData: OrderedXmlNode,
): OrderedXmlNode[] {
  const existingRows = new Map(
    elementChildren(existingSheetData)
      .filter((entry) => elementLocalName(entry) === "row")
      .map((entry) => [attribute(entry, "r"), entry]),
  );
  const prefix = elementPrefix(existingSheetData);
  const rows: OrderedXmlNode[] = [];
  let headerRow = 1;
  for (const series of input.series) {
    const headerRowNumber = String(headerRow);
    const existingHeaderRow = existingRows.get(headerRowNumber);
    rows.push(
      createElement(
        `${prefix}row`,
        [
          worksheetStringCell(
            `B${headerRowNumber}`,
            series.name,
            existingCell(existingHeaderRow, `B${headerRowNumber}`),
            prefix,
          ),
        ],
        preservedAttributes(existingHeaderRow, headerRowNumber),
      ),
    );
    for (const [pointIndex, xValue] of series.xValues.entries()) {
      const rowNumber = String(headerRow + pointIndex + 1);
      const existingRow = existingRows.get(rowNumber);
      rows.push(
        createElement(
          `${prefix}row`,
          [
            worksheetNumberCell(
              `A${rowNumber}`,
              xValue,
              existingCell(existingRow, `A${rowNumber}`),
              prefix,
            ),
            worksheetNumberCell(
              `B${rowNumber}`,
              series.yValues[pointIndex],
              existingCell(existingRow, `B${rowNumber}`),
              prefix,
            ),
          ],
          preservedAttributes(existingRow, rowNumber),
        ),
      );
    }
    headerRow += series.xValues.length + 2;
  }
  return rows;
}

function buildBubbleWorksheetRows(
  input: UpdateBubbleChartDataInput,
  existingSheetData: OrderedXmlNode,
): OrderedXmlNode[] {
  const existingRows = new Map(
    elementChildren(existingSheetData)
      .filter((entry) => elementLocalName(entry) === "row")
      .map((entry) => [attribute(entry, "r"), entry]),
  );
  const prefix = elementPrefix(existingSheetData);
  const rows: OrderedXmlNode[] = [];
  let headerRow = 1;
  for (const series of input.series) {
    const headerRowNumber = String(headerRow);
    const existingHeaderRow = existingRows.get(headerRowNumber);
    rows.push(
      createElement(
        `${prefix}row`,
        [
          worksheetStringCell(
            `B${headerRowNumber}`,
            series.name,
            existingCell(existingHeaderRow, `B${headerRowNumber}`),
            prefix,
          ),
          worksheetStringCell(
            `C${headerRowNumber}`,
            "Size",
            existingCell(existingHeaderRow, `C${headerRowNumber}`),
            prefix,
          ),
        ],
        preservedAttributes(existingHeaderRow, headerRowNumber),
      ),
    );
    for (const [pointIndex, xValue] of series.xValues.entries()) {
      const rowNumber = String(headerRow + pointIndex + 1);
      const existingRow = existingRows.get(rowNumber);
      rows.push(
        createElement(
          `${prefix}row`,
          [
            worksheetNumberCell(
              `A${rowNumber}`,
              xValue,
              existingCell(existingRow, `A${rowNumber}`),
              prefix,
            ),
            worksheetNumberCell(
              `B${rowNumber}`,
              series.yValues[pointIndex],
              existingCell(existingRow, `B${rowNumber}`),
              prefix,
            ),
            worksheetNumberCell(
              `C${rowNumber}`,
              series.bubbleSizes[pointIndex],
              existingCell(existingRow, `C${rowNumber}`),
              prefix,
            ),
          ],
          preservedAttributes(existingRow, rowNumber),
        ),
      );
    }
    headerRow += series.xValues.length + 2;
  }
  return rows;
}

function scatterLastDataRow(pointCounts: readonly number[]): number {
  return pointCounts.reduce((row, pointCount) => row + pointCount + 2, -1);
}

function buildWorksheetRows(
  input: UpdateChartDataInput,
  existingSheetData: OrderedXmlNode,
): OrderedXmlNode[] {
  const existingRows = new Map(
    elementChildren(existingSheetData)
      .filter((entry) => elementLocalName(entry) === "row")
      .map((entry) => [attribute(entry, "r"), entry]),
  );
  const prefix = elementPrefix(existingSheetData);
  const headerCells = [
    worksheetStringCell("A1", "Category", existingCell(existingRows.get("1"), "A1"), prefix),
    ...input.series.map((series, index) =>
      worksheetStringCell(
        `${spreadsheetColumn(index + 2)}1`,
        series.name,
        existingCell(existingRows.get("1"), `${spreadsheetColumn(index + 2)}1`),
        prefix,
      ),
    ),
  ];
  const rows = [
    createElement(`${prefix}row`, headerCells, preservedAttributes(existingRows.get("1"), "1")),
  ];
  for (let row = 0; row < input.series[0].categories.length; row += 1) {
    const rowNumber = String(row + 2);
    const existingRow = existingRows.get(rowNumber);
    rows.push(
      createElement(
        `${prefix}row`,
        [
          worksheetStringCell(
            `A${rowNumber}`,
            input.series[0].categories[row],
            existingCell(existingRow, `A${rowNumber}`),
            prefix,
          ),
          ...input.series.map((series, index) => {
            const reference = `${spreadsheetColumn(index + 2)}${rowNumber}`;
            return worksheetNumberCell(
              reference,
              series.values[row],
              existingCell(existingRow, reference),
              prefix,
            );
          }),
        ],
        preservedAttributes(existingRow, rowNumber),
      ),
    );
  }
  return rows;
}

function worksheetStringCell(
  reference: string,
  value: string,
  existing: OrderedXmlNode | undefined,
  prefix: string,
): OrderedXmlNode {
  return createElement(
    `${prefix}c`,
    [createElement(`${prefix}is`, [createTextElement(`${prefix}t`, value)])],
    { ...preservedAttributes(existing, reference), "@_t": "inlineStr" },
  );
}

function worksheetNumberCell(
  reference: string,
  value: number,
  existing: OrderedXmlNode | undefined,
  prefix: string,
): OrderedXmlNode {
  const attributes = preservedAttributes(existing, reference);
  delete attributes["@_t"];
  return createElement(`${prefix}c`, [createElement(`${prefix}v`, [{ "#text": String(value) }])], {
    ...attributes,
  });
}

function createTextElement(name: string, value: string): OrderedXmlNode {
  return createElement(
    name,
    [{ "#text": value }],
    value.startsWith(" ") || value.endsWith(" ") ? { "@_xml:space": "preserve" } : undefined,
  );
}

function existingCell(
  row: OrderedXmlNode | undefined,
  reference: string,
): OrderedXmlNode | undefined {
  return row === undefined
    ? undefined
    : elementChildren(row).find(
        (entry) => elementLocalName(entry) === "c" && attribute(entry, "r") === reference,
      );
}

function preservedAttributes(
  entry: OrderedXmlNode | undefined,
  reference: string,
): Record<string, unknown> {
  const attributes = entry?.[":@"];
  const result =
    typeof attributes === "object" && attributes !== null
      ? { ...unsafeOoxmlBoundaryAssertion<Record<string, unknown>>(attributes) }
      : {};
  const referenceKey =
    Object.keys(result).find((key) => key === "@_r" || key.endsWith(":r")) ?? "@_r";
  result[referenceKey] = reference;
  return result;
}

function normalizeWorkbookTarget(target: string): string {
  const segments = target.startsWith("/")
    ? target.slice(1).split("/")
    : ["xl", ...target.split("/")];
  const normalized: string[] = [];
  for (const segment of segments) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") normalized.pop();
    else normalized.push(segment);
  }
  return normalized.join("/");
}

function requireSingleChild(parent: OrderedXmlNode, localName: string): OrderedXmlNode {
  return requireSingleElement(
    elementChildren(parent),
    localName,
    `chart XML has no ${localName} element`,
  );
}

function requireSingleElement(
  nodes: readonly OrderedXmlNode[],
  localName: string,
  missingMessage: string,
): OrderedXmlNode {
  const matches = nodes.filter((entry) => elementLocalName(entry) === localName);
  if (matches.length !== 1) throw new Error(`updateChartData: ${missingMessage}`);
  return matches[0];
}

function elementChildren(entry: OrderedXmlNode): readonly OrderedXmlNode[] {
  return mutableElementChildren(entry);
}

function mutableElementChildren(entry: OrderedXmlNode): OrderedXmlNode[] {
  const name = elementName(entry);
  const value = name === undefined ? undefined : entry[name];
  if (!Array.isArray(value)) return [];
  return unsafeOoxmlBoundaryAssertion<OrderedXmlNode[]>(value);
}

function setElementChildren(entry: OrderedXmlNode, children: OrderedXmlNode[]): void {
  const name = elementName(entry);
  if (name === undefined) throw new Error("updateChartData: malformed chart XML element");
  entry[name] = children;
}

function elementName(entry: OrderedXmlNode): string | undefined {
  return Object.keys(entry).find((key) => key !== ":@");
}

function elementLocalName(entry: OrderedXmlNode): string | undefined {
  const name = elementName(entry);
  if (name === undefined || name === "#text" || name === "?xml") return undefined;
  const colon = name.indexOf(":");
  return colon < 0 ? name : name.slice(colon + 1);
}

function elementPrefix(entry: OrderedXmlNode): string {
  const name = elementName(entry);
  if (name === undefined) return "c:";
  const colon = name.indexOf(":");
  return colon < 0 ? "" : name.slice(0, colon + 1);
}

function elementText(entry: OrderedXmlNode): string | undefined {
  const text = mutableElementChildren(entry).find((child) => "#text" in child)?.["#text"];
  return typeof text === "string" ? text : undefined;
}

function replaceElementText(entry: OrderedXmlNode, value: string): void {
  const children = mutableElementChildren(entry);
  const nonText = children.filter((child) => !("#text" in child));
  setElementChildren(entry, [{ "#text": value }, ...nonText]);
}

function setAttribute(entry: OrderedXmlNode, localName: string, value: string): void {
  const attributes = entry[":@"];
  const record =
    typeof attributes === "object" && attributes !== null
      ? unsafeOoxmlBoundaryAssertion<Record<string, unknown>>(attributes)
      : {};
  const existingKey = Object.keys(record).find((key) => {
    const name = key.startsWith("@_") ? key.slice(2) : key;
    const colon = name.indexOf(":");
    return (colon < 0 ? name : name.slice(colon + 1)) === localName;
  });
  record[existingKey ?? `@_${localName}`] = value;
  entry[":@"] = record;
}

function attribute(entry: OrderedXmlNode, localName: string): string | undefined {
  const attributes = entry[":@"];
  if (typeof attributes !== "object" || attributes === null) return undefined;
  const record = unsafeOoxmlBoundaryAssertion<Record<string, unknown>>(attributes);
  const value = Object.entries(record).find(([key]) => {
    const name = key.startsWith("@_") ? key.slice(2) : key;
    const colon = name.indexOf(":");
    return (colon < 0 ? name : name.slice(colon + 1)) === localName;
  })?.[1];
  return typeof value === "string" ? value : undefined;
}

function namespacedAttribute(entry: OrderedXmlNode, localName: string): string | undefined {
  const attributes = entry[":@"];
  if (typeof attributes !== "object" || attributes === null) return undefined;
  const record = unsafeOoxmlBoundaryAssertion<Record<string, unknown>>(attributes);
  for (const [key, value] of Object.entries(record)) {
    const name = key.startsWith("@_") ? key.slice(2) : key;
    const colon = name.indexOf(":");
    if (colon >= 0 && name.slice(colon + 1) === localName && typeof value === "string")
      return value;
  }
  return undefined;
}

function createElement(
  name: string,
  children: OrderedXmlNode[],
  attributes?: Record<string, unknown>,
): OrderedXmlNode {
  return { [name]: children, ...(attributes === undefined ? {} : { ":@": attributes }) };
}

function spreadsheetColumn(index: number): string {
  let value = index;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function spreadsheetColumnNumber(value: string): number {
  let result = 0;
  for (const char of value) result = result * 26 + (char.charCodeAt(0) - 64);
  return result;
}
