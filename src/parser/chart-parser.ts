import type { ChartData, ChartType, ChartSeries, ChartLegend } from "../model/chart.js";
import type { ResolvedColor } from "../model/fill.js";
import type { ColorResolver } from "../color/color-resolver.js";
import { parseXml, type XmlNode } from "./xml-parser.js";

const ACCENT_KEYS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const;

const CHART_TYPE_MAP: [string, ChartType][] = [
  ["barChart", "bar"],
  ["bar3DChart", "bar"],
  ["lineChart", "line"],
  ["line3DChart", "line"],
  ["pieChart", "pie"],
  ["pie3DChart", "pie"],
  ["doughnutChart", "pie"],
  ["scatterChart", "scatter"],
];

export function parseChart(chartXml: string, colorResolver: ColorResolver): ChartData | null {
  const parsed = parseXml(chartXml);
  const chartSpace = parsed.chartSpace as XmlNode | undefined;
  if (!chartSpace) return null;

  const chart = chartSpace.chart as XmlNode | undefined;
  if (!chart) return null;

  const plotArea = chart.plotArea as XmlNode | undefined;
  if (!plotArea) return null;

  const title = parseChartTitle(chart.title as XmlNode);
  const { chartType, series, categories, barDirection } = parseChartTypeAndData(
    plotArea,
    colorResolver,
  );
  if (!chartType) return null;

  const legend = parseLegend(chart.legend as XmlNode);

  return {
    chartType,
    title,
    series,
    categories,
    ...(barDirection !== undefined && { barDirection }),
    legend,
  };
}

function parseChartTypeAndData(
  plotArea: XmlNode,
  colorResolver: ColorResolver,
): {
  chartType: ChartType | null;
  series: ChartSeries[];
  categories: string[];
  barDirection?: "col" | "bar";
} {
  for (const [xmlTag, chartType] of CHART_TYPE_MAP) {
    const chartNode = plotArea[xmlTag] as XmlNode | undefined;
    if (!chartNode) continue;

    const serList = (chartNode.ser as XmlNode[] | undefined) ?? [];
    const series = serList.map((ser: XmlNode, index: number) =>
      parseSeries(ser, chartType, index, colorResolver),
    );
    const categories = extractCategories(serList);
    const barDirNode = chartNode.barDir as XmlNode | undefined;
    const barDirection =
      chartType === "bar"
        ? (((barDirNode?.["@_val"] as string | undefined) ?? "col") as "col" | "bar")
        : undefined;

    return { chartType, series, categories, barDirection };
  }

  return { chartType: null, series: [], categories: [] };
}

function parseSeries(
  ser: XmlNode,
  chartType: ChartType,
  seriesIndex: number,
  colorResolver: ColorResolver,
): ChartSeries {
  const name = parseSeriesName(ser.tx as XmlNode);
  const values = parseNumericData(
    chartType === "scatter" ? (ser.yVal as XmlNode) : (ser.val as XmlNode),
  );
  const xValues = chartType === "scatter" ? parseNumericData(ser.xVal as XmlNode) : undefined;
  const color = resolveSeriesColor(ser.spPr as XmlNode, seriesIndex, colorResolver);

  return {
    name,
    values,
    ...(xValues !== undefined && { xValues }),
    color,
  };
}

function parseSeriesName(tx: XmlNode): string | null {
  if (!tx) return null;
  const strRef = tx.strRef as XmlNode | undefined;
  const strCache = strRef?.strCache as XmlNode | undefined;
  if (strCache?.pt) {
    const pts = strCache.pt as XmlNode[];
    return (pts[0]?.v as string | undefined) ?? null;
  }
  if (typeof tx.v === "string") return tx.v;
  return null;
}

function parseNumericData(valNode: XmlNode): number[] {
  if (!valNode) return [];
  const numRef = valNode.numRef as XmlNode | undefined;
  const numCache = numRef?.numCache as XmlNode | undefined;
  if (!numCache?.pt) return [];
  const pts = numCache.pt as XmlNode[];

  return pts
    .slice()
    .sort((a: XmlNode, b: XmlNode) => Number(a["@_idx"]) - Number(b["@_idx"]))
    .map((pt: XmlNode) => Number(pt.v ?? 0));
}

function extractCategories(serList: XmlNode[]): string[] {
  for (const ser of serList) {
    const cat = ser.cat as XmlNode | undefined;
    if (!cat) continue;
    const strRef = cat.strRef as XmlNode | undefined;
    const numRef = cat.numRef as XmlNode | undefined;
    const strCache =
      (strRef?.strCache as XmlNode | undefined) ?? (numRef?.numCache as XmlNode | undefined);
    if (strCache?.pt) {
      return (strCache.pt as XmlNode[])
        .slice()
        .sort((a: XmlNode, b: XmlNode) => Number(a["@_idx"]) - Number(b["@_idx"]))
        .map((pt: XmlNode) => String((pt.v as string | number | undefined) ?? ""));
    }
  }
  return [];
}

function resolveSeriesColor(
  spPr: XmlNode,
  seriesIndex: number,
  colorResolver: ColorResolver,
): ResolvedColor {
  if (spPr) {
    const resolved = colorResolver.resolve((spPr.solidFill as XmlNode | undefined) ?? spPr);
    if (resolved) return resolved;
  }
  // Default: use theme accent colors
  const accentKey = ACCENT_KEYS[seriesIndex % ACCENT_KEYS.length];
  const resolved = colorResolver.resolve({ schemeClr: { "@_val": accentKey } });
  return resolved ?? { hex: "#4472C4", alpha: 1 };
}

function parseChartTitle(titleNode: XmlNode): string | null {
  if (!titleNode) return null;
  const tx = titleNode.tx as XmlNode | undefined;
  const rich = tx?.rich as XmlNode | undefined;
  if (!rich?.p) return null;
  const pList = Array.isArray(rich.p) ? (rich.p as XmlNode[]) : [rich.p as XmlNode];
  const texts: string[] = [];
  for (const p of pList) {
    const rList = (p.r as XmlNode[] | undefined) ?? [];
    for (const r of rList) {
      const t = r.t;
      if (typeof t === "string") texts.push(t);
      else if (t && (t as XmlNode)["#text"]) texts.push(String((t as XmlNode)["#text"]));
    }
  }
  return texts.length > 0 ? texts.join("") : null;
}

function parseLegend(legendNode: XmlNode): ChartLegend | null {
  if (!legendNode) return null;
  const legendPos = legendNode.legendPos as XmlNode | undefined;
  const pos = ((legendPos?.["@_val"] as string | undefined) ?? "b") as ChartLegend["position"];
  return { position: pos };
}
