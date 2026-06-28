import { parseColorElement } from "../reader/drawing.js";
import {
  getAttr,
  getChild,
  getChildArray,
  getChildText,
  parseXml,
  type XmlNode,
} from "../reader/xml.js";
import type { SourceColor, SourceTheme } from "../source/index.js";
import { resolveColor } from "./color.js";
import type {
  ComputedChartData,
  ComputedChartLegend,
  ComputedChartSeries,
  ComputedChartType,
  ComputedColor,
} from "./pptx-computed-view.js";

interface ChartColorContext {
  readonly theme?: SourceTheme;
  readonly colorMap: Readonly<Record<string, string>>;
}

const ACCENT_KEYS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const;

const CHART_TYPE_MAP: [string, ComputedChartType][] = [
  ["barChart", "bar"],
  ["bar3DChart", "bar"],
  ["lineChart", "line"],
  ["line3DChart", "line"],
  ["pieChart", "pie"],
  ["pie3DChart", "pie"],
  ["doughnutChart", "doughnut"],
  ["scatterChart", "scatter"],
  ["bubbleChart", "bubble"],
  ["areaChart", "area"],
  ["area3DChart", "area"],
  ["radarChart", "radar"],
  ["stockChart", "stock"],
  ["surfaceChart", "surface"],
  ["surface3DChart", "surface"],
  ["ofPieChart", "ofPie"],
];

export function parseComputedChartData(
  chartXml: string,
  colorContext: ChartColorContext,
): ComputedChartData | undefined {
  const parsed = parseXml(chartXml);
  const chartSpace = getChild(parsed, "chartSpace");
  if (chartSpace === undefined) return undefined;

  const chart = getChild(chartSpace, "chart");
  if (chart === undefined) return undefined;

  const plotArea = getChild(chart, "plotArea");
  if (plotArea === undefined) return undefined;

  const title = parseChartTitle(getChild(chart, "title"));
  const parsedChart = parseChartTypeAndData(plotArea, colorContext);
  if (parsedChart.chartType === undefined) return undefined;

  return {
    chartType: parsedChart.chartType,
    title,
    series: parsedChart.series,
    categories: parsedChart.categories,
    ...(parsedChart.barDirection !== undefined ? { barDirection: parsedChart.barDirection } : {}),
    ...(parsedChart.holeSize !== undefined ? { holeSize: parsedChart.holeSize } : {}),
    ...(parsedChart.radarStyle !== undefined ? { radarStyle: parsedChart.radarStyle } : {}),
    ...(parsedChart.ofPieType !== undefined ? { ofPieType: parsedChart.ofPieType } : {}),
    ...(parsedChart.secondPieSize !== undefined
      ? { secondPieSize: parsedChart.secondPieSize }
      : {}),
    ...(parsedChart.splitPos !== undefined ? { splitPos: parsedChart.splitPos } : {}),
    legend: parseLegend(getChild(chart, "legend")),
  };
}

function parseChartTypeAndData(
  plotArea: XmlNode,
  colorContext: ChartColorContext,
): {
  readonly chartType?: ComputedChartType;
  readonly series: ComputedChartSeries[];
  readonly categories: string[];
  readonly barDirection?: "col" | "bar";
  readonly holeSize?: number;
  readonly radarStyle?: "standard" | "marker" | "filled";
  readonly ofPieType?: "pie" | "bar";
  readonly secondPieSize?: number;
  readonly splitPos?: number;
} {
  for (const [xmlTag, chartType] of CHART_TYPE_MAP) {
    const chartNode = getChild(plotArea, xmlTag);
    if (chartNode === undefined) continue;

    const serList = getChildArray(chartNode, "ser");
    const series = serList.map((ser, index) => parseSeries(ser, chartType, index, colorContext));
    const categories = extractCategories(serList);
    const barDirection =
      chartType === "bar"
        ? parseBarDirection(getAttr(getChild(chartNode, "barDir"), "val"))
        : undefined;

    const holeSize =
      chartType === "doughnut"
        ? parseNumberAttribute(getChild(chartNode, "holeSize"), "val", 50)
        : undefined;

    const radarStyle =
      chartType === "radar"
        ? parseRadarStyle(getAttr(getChild(chartNode, "radarStyle"), "val"))
        : undefined;

    const ofPieType =
      chartType === "ofPie"
        ? parseOfPieType(getAttr(getChild(chartNode, "ofPieType"), "val"))
        : undefined;

    const secondPieSize =
      chartType === "ofPie"
        ? parseNumberAttribute(getChild(chartNode, "secondPieSize"), "val", 75)
        : undefined;

    const splitPos =
      chartType === "ofPie"
        ? parseNumberAttribute(getChild(chartNode, "splitPos"), "val", 2)
        : undefined;

    return {
      chartType,
      series,
      categories,
      ...(barDirection !== undefined ? { barDirection } : {}),
      ...(holeSize !== undefined ? { holeSize } : {}),
      ...(radarStyle !== undefined ? { radarStyle } : {}),
      ...(ofPieType !== undefined ? { ofPieType } : {}),
      ...(secondPieSize !== undefined ? { secondPieSize } : {}),
      ...(splitPos !== undefined ? { splitPos } : {}),
    };
  }

  return { series: [], categories: [] };
}

function parseSeries(
  ser: XmlNode,
  chartType: ComputedChartType,
  seriesIndex: number,
  colorContext: ChartColorContext,
): ComputedChartSeries {
  const name = parseSeriesName(getChild(ser, "tx"));
  const usesXY = chartType === "scatter" || chartType === "bubble";
  const values = parseNumericData(usesXY ? getChild(ser, "yVal") : getChild(ser, "val"));
  const xValues = usesXY ? parseNumericData(getChild(ser, "xVal")) : undefined;
  const bubbleSizes =
    chartType === "bubble" ? parseNumericData(getChild(ser, "bubbleSize")) : undefined;
  const color = resolveSeriesColor(getChild(ser, "spPr"), seriesIndex, colorContext);

  return {
    name,
    values,
    ...(xValues !== undefined ? { xValues } : {}),
    ...(bubbleSizes !== undefined ? { bubbleSizes } : {}),
    color,
  };
}

function parseSeriesName(tx: XmlNode | undefined): string | null {
  if (tx === undefined) return null;
  const strCache = getChild(getChild(tx, "strRef"), "strCache");
  const cachedValue = firstIndexedPointText(strCache);
  if (cachedValue !== undefined) return cachedValue;
  return getChildText(tx, "v") ?? null;
}

function parseNumericData(valNode: XmlNode | undefined): number[] {
  const numCache = getChild(getChild(valNode, "numRef"), "numCache") ?? getChild(valNode, "numLit");
  return indexedPointTexts(numCache).map(parseNumericPointValue);
}

function extractCategories(serList: readonly XmlNode[]): string[] {
  for (const ser of serList) {
    const cat = getChild(ser, "cat");
    if (cat === undefined) continue;

    const categoryCache =
      getChild(getChild(cat, "strRef"), "strCache") ??
      getChild(cat, "strLit") ??
      getChild(getChild(cat, "numRef"), "numCache") ??
      getChild(cat, "numLit");
    const values = indexedPointTexts(categoryCache).map((value) => String(value ?? ""));
    if (values.length > 0) return values;

    const multiCache =
      getChild(getChild(cat, "multiLvlStrRef"), "multiLvlStrCache") ??
      getChild(cat, "multiLvlStrLit");
    const firstLevel = getChildArray(multiCache, "lvl")[0];
    const multiValues = indexedPointTexts(firstLevel).map((value) => String(value ?? ""));
    if (multiValues.length > 0) return multiValues;
  }
  return [];
}

function indexedPointTexts(cache: XmlNode | undefined): string[] {
  return getChildArray(cache, "pt")
    .slice()
    .sort((a, b) => Number(getAttr(a, "idx")) - Number(getAttr(b, "idx")))
    .map((point) => getChildText(point, "v") ?? "");
}

function firstIndexedPointText(cache: XmlNode | undefined): string | undefined {
  return indexedPointTexts(cache)[0];
}

function parseNumericPointValue(value: string | undefined): number {
  const parsed = Number(value ?? 0);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseNumberAttribute(
  node: XmlNode | undefined,
  attrName: string,
  fallback: number,
): number {
  const parsed = Number(getAttr(node, attrName) ?? fallback);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function resolveSeriesColor(
  spPr: XmlNode | undefined,
  seriesIndex: number,
  colorContext: ChartColorContext,
): ComputedColor {
  const sourceColor = parseSeriesSourceColor(spPr);
  if (sourceColor !== undefined) {
    const resolved = resolveColor(colorContext, sourceColor);
    if (resolved !== undefined) return resolved;
  }

  const accentKey = ACCENT_KEYS[seriesIndex % ACCENT_KEYS.length];
  return (
    resolveColor(colorContext, { kind: "scheme", scheme: accentKey }) ?? {
      hex: "#4472c4",
      alpha: 1,
    }
  );
}

function parseSeriesSourceColor(spPr: XmlNode | undefined): SourceColor | undefined {
  return parseColorElement(getChild(spPr, "solidFill")) ?? parseColorElement(spPr);
}

function parseChartTitle(titleNode: XmlNode | undefined): string | null {
  const rich = getChild(getChild(titleNode, "tx"), "rich");
  if (rich === undefined) return null;

  const texts: string[] = [];
  for (const paragraph of getChildArray(rich, "p")) {
    for (const run of getChildArray(paragraph, "r")) {
      const text = getChildText(run, "t");
      if (text !== undefined) texts.push(text);
    }
  }
  return texts.length > 0 ? texts.join("") : null;
}

function parseLegend(legendNode: XmlNode | undefined): ComputedChartLegend | null {
  if (legendNode === undefined) return null;
  return { position: parseLegendPosition(getAttr(getChild(legendNode, "legendPos"), "val")) };
}

function parseBarDirection(value: string | undefined): "col" | "bar" {
  return value === "bar" ? "bar" : "col";
}

function parseRadarStyle(value: string | undefined): "standard" | "marker" | "filled" {
  if (value === "marker" || value === "filled") return value;
  return "standard";
}

function parseOfPieType(value: string | undefined): "pie" | "bar" {
  return value === "bar" ? "bar" : "pie";
}

function parseLegendPosition(value: string | undefined): ComputedChartLegend["position"] {
  if (value === "b" || value === "t" || value === "l" || value === "tr") return value;
  return "r";
}
