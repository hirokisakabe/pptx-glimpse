import type { ChartData, ChartType, ChartSeries, ChartLegend } from "../model/chart.js";
import type { ResolvedColor } from "../model/fill.js";
import type { ColorResolver } from "../color/color-resolver.js";
import { parseXml } from "./xml-parser.js";

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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(chartXml) as any;
  const chartSpace = parsed.chartSpace;
  if (!chartSpace) return null;

  const chart = chartSpace.chart;
  if (!chart) return null;

  const plotArea = chart.plotArea;
  if (!plotArea) return null;

  const title = parseChartTitle(chart.title);
  const { chartType, series, categories, barDirection } = parseChartTypeAndData(
    plotArea,
    colorResolver,
  );
  if (!chartType) return null;

  const legend = parseLegend(chart.legend);

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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  plotArea: any,
  colorResolver: ColorResolver,
): {
  chartType: ChartType | null;
  series: ChartSeries[];
  categories: string[];
  barDirection?: "col" | "bar";
} {
  for (const [xmlTag, chartType] of CHART_TYPE_MAP) {
    const chartNode = plotArea[xmlTag];
    if (!chartNode) continue;

    const serList = chartNode.ser ?? [];
    const series = serList.map(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (ser: any, index: number) => parseSeries(ser, chartType, index, colorResolver),
    );
    const categories = extractCategories(serList);
    const barDirection = chartType === "bar" ? (chartNode.barDir?.["@_val"] ?? "col") : undefined;

    return { chartType, series, categories, barDirection };
  }

  return { chartType: null, series: [], categories: [] };
}

function parseSeries(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ser: any,
  chartType: ChartType,
  seriesIndex: number,
  colorResolver: ColorResolver,
): ChartSeries {
  const name = parseSeriesName(ser.tx);
  const values = parseNumericData(chartType === "scatter" ? ser.yVal : ser.val);
  const xValues = chartType === "scatter" ? parseNumericData(ser.xVal) : undefined;
  const color = resolveSeriesColor(ser.spPr, seriesIndex, colorResolver);

  return {
    name,
    values,
    ...(xValues !== undefined && { xValues }),
    color,
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseSeriesName(tx: any): string | null {
  if (!tx) return null;
  const strCache = tx.strRef?.strCache;
  if (strCache?.pt) {
    const pts = strCache.pt;
    return pts[0]?.v ?? null;
  }
  if (typeof tx.v === "string") return tx.v;
  return null;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseNumericData(valNode: any): number[] {
  if (!valNode) return [];
  const numCache = valNode.numRef?.numCache;
  if (!numCache?.pt) return [];
  const pts = numCache.pt;

  return (
    pts
      .slice()
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .sort((a: any, b: any) => Number(a["@_idx"]) - Number(b["@_idx"]))
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .map((pt: any) => Number(pt.v ?? 0))
  );
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractCategories(serList: any[]): string[] {
  for (const ser of serList) {
    const cat = ser.cat;
    if (!cat) continue;
    const strCache = cat.strRef?.strCache ?? cat.numRef?.numCache;
    if (strCache?.pt) {
      return (
        strCache.pt
          .slice()
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          .sort((a: any, b: any) => Number(a["@_idx"]) - Number(b["@_idx"]))
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          .map((pt: any) => String(pt.v ?? ""))
      );
    }
  }
  return [];
}

function resolveSeriesColor(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  spPr: any,
  seriesIndex: number,
  colorResolver: ColorResolver,
): ResolvedColor {
  if (spPr) {
    const resolved = colorResolver.resolve(spPr.solidFill ?? spPr);
    if (resolved) return resolved;
  }
  // Default: use theme accent colors
  const accentKey = ACCENT_KEYS[seriesIndex % ACCENT_KEYS.length];
  const resolved = colorResolver.resolve({ schemeClr: { "@_val": accentKey } });
  return resolved ?? { hex: "#4472C4", alpha: 1 };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseChartTitle(titleNode: any): string | null {
  if (!titleNode) return null;
  const rich = titleNode.tx?.rich;
  if (!rich?.p) return null;
  const pList = Array.isArray(rich.p) ? rich.p : [rich.p];
  const texts: string[] = [];
  for (const p of pList) {
    const rList = p.r ?? [];
    for (const r of rList) {
      const t = r.t;
      if (typeof t === "string") texts.push(t);
      else if (t?.["#text"]) texts.push(String(t["#text"]));
    }
  }
  return texts.length > 0 ? texts.join("") : null;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseLegend(legendNode: any): ChartLegend | null {
  if (!legendNode) return null;
  const pos = legendNode.legendPos?.["@_val"] ?? "b";
  return { position: pos };
}
