import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { unzipSync, zipSync } from "fflate";

import { getAttr, getChild, getChildArray, getNamespacedAttr, parseXml } from "../reader/xml.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import { copyBytes, requireRawBinaryPart } from "./editing-shared.js";
import type {
  PartPath,
  PptxSourceModel,
  PptxSourceModelUpdateChartDataEdit,
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
}

export interface UpdateChartDataInput {
  readonly series: readonly UpdateChartSeriesDataInput[];
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
 * The supported layout is deliberately narrow: one embedded workbook, one worksheet, the existing
 * series count, names in row 1, categories in column A, and values in columns B onward.
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
    binding.seriesCount,
    binding.pointCount,
  );

  const updatedChartBytes = encoder.encode(orderedBuilder.build(orderedRoot));
  const updatedWorkbookBytes = updateEmbeddedWorkbook(workbook, input);
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
  readonly seriesCount: number;
  readonly pointCount: number;
} {
  const chartSpace = requireSingleElement(root, "chartSpace", "chart XML has no chartSpace root");
  const chart = requireSingleChild(chartSpace, "chart");
  const plotArea = requireSingleChild(chart, "plotArea");
  const chartGroups = elementChildren(plotArea).filter((entry) =>
    elementLocalName(entry)?.endsWith("Chart"),
  );
  const chartGroupName =
    chartGroups[0] === undefined ? undefined : elementLocalName(chartGroups[0]);
  if (
    chartGroups.length !== 1 ||
    chartGroupName === undefined ||
    !SUPPORTED_CHART_ELEMENTS.has(chartGroupName)
  ) {
    throw new Error("updateChartData: chart type or combination is not supported");
  }
  const series = elementChildren(chartGroups[0]).filter(
    (entry) => elementLocalName(entry) === "ser",
  );
  if (series.length !== input.series.length) {
    throw new Error("updateChartData: changing the existing series count is not supported");
  }

  let sheetName: string | undefined;
  let pointCount: number | undefined;
  for (const [index, seriesEntry] of series.entries()) {
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
    pointCount = seriesPointCount;

    setFormula(txRef, `${txFormula.sheetToken}!$${spreadsheetColumn(index + 2)}$1`);
    setStringCache(requireSingleChild(txRef, "strCache"), [input.series[index].name]);
    setFormula(
      categoryRef,
      `${txFormula.sheetToken}!$A$2:$A$${input.series[index].categories.length + 1}`,
    );
    setStringCache(requireSingleChild(categoryRef, "strCache"), input.series[index].categories);
    setFormula(
      valueRef,
      `${txFormula.sheetToken}!$${spreadsheetColumn(index + 2)}$2:$${spreadsheetColumn(index + 2)}$${input.series[index].values.length + 1}`,
    );
    setNumberCache(requireSingleChild(valueRef, "numCache"), input.series[index].values);
  }

  const externalData = requireSingleChild(chartSpace, "externalData");
  const externalDataRelationshipId = namespacedAttribute(externalData, "id");
  if (externalDataRelationshipId === undefined) {
    throw new Error("updateChartData: chart externalData has no relationship id");
  }
  if (sheetName === undefined || pointCount === undefined) {
    throw new Error("updateChartData: chart has no series");
  }
  return {
    externalDataRelationshipId,
    sheetName,
    seriesCount: series.length,
    pointCount,
  };
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
  assertSupportedWorksheetLayout(worksheetBytes, seriesCount, pointCount);
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
  const cells = getChildArray(getChild(worksheet, "sheetData"), "row").flatMap((row) =>
    getChildArray(row, "c"),
  );
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
