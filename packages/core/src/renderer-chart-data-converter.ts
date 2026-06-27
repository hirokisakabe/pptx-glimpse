import type { ChartData, ChartLegend, ChartSeries, ChartType } from "@pptx-glimpse/renderer";
import type { ResolvedColor } from "@pptx-glimpse/renderer";

import type { ColorResolver } from "./color/color-resolver.js";
import { parseXml, type XmlNode } from "./parser/xml-parser.js";
import { unsafeTypeAssertion } from "./unsafe-type-assertion.js";

const ACCENT_KEYS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const;

const CHART_TYPE_MAP: [string, ChartType][] = [
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

/**
 * Renderer-specific chart conversion contract.
 *
 * Input: resolved chart part XML from the computed view plus the slide color resolver.
 * Output: the current renderer ChartData model, or null when the XML is unsupported.
 *
 * This stays in core because ChartData is a renderer compatibility target.
 * Future document work should own chart source/computed semantics only after it
 * covers chart data, style/color parts, embedded workbook references, and
 * compatibility expectations. The renderer ChartData adapter must stay outside
 * document or be replaced by another renderer-owned contract.
 */
export function convertChartXmlToRendererChartData(
  chartXml: string,
  colorResolver: ColorResolver,
): ChartData | null {
  const parsed = parseXml(chartXml);
  const chartSpace = unsafeTypeAssertion<XmlNode | undefined>(parsed.chartSpace);
  if (!chartSpace) return null;

  const chart = unsafeTypeAssertion<XmlNode | undefined>(chartSpace.chart);
  if (!chart) return null;

  const plotArea = unsafeTypeAssertion<XmlNode | undefined>(chart.plotArea);
  if (!plotArea) return null;

  const title = parseChartTitle(unsafeTypeAssertion<XmlNode>(chart.title));
  const {
    chartType,
    series,
    categories,
    barDirection,
    holeSize,
    radarStyle,
    ofPieType,
    secondPieSize,
    splitPos,
  } = parseChartTypeAndData(plotArea, colorResolver);
  if (!chartType) return null;

  const legend = parseLegend(unsafeTypeAssertion<XmlNode>(chart.legend));

  return {
    chartType,
    title,
    series,
    categories,
    ...(barDirection !== undefined && { barDirection }),
    ...(holeSize !== undefined && { holeSize }),
    ...(radarStyle !== undefined && { radarStyle }),
    ...(ofPieType !== undefined && { ofPieType }),
    ...(secondPieSize !== undefined && { secondPieSize }),
    ...(splitPos !== undefined && { splitPos }),
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
  holeSize?: number;
  radarStyle?: "standard" | "marker" | "filled";
  ofPieType?: "pie" | "bar";
  secondPieSize?: number;
  splitPos?: number;
} {
  for (const [xmlTag, chartType] of CHART_TYPE_MAP) {
    const chartNode = unsafeTypeAssertion<XmlNode | undefined>(plotArea[xmlTag]);
    if (!chartNode) continue;

    const serList = unsafeTypeAssertion<XmlNode[] | undefined>(chartNode.ser) ?? [];
    const series = serList.map((ser: XmlNode, index: number) =>
      parseSeries(ser, chartType, index, colorResolver),
    );
    const categories = extractCategories(serList);
    const barDirNode = unsafeTypeAssertion<XmlNode | undefined>(chartNode.barDir);
    const barDirection =
      chartType === "bar"
        ? unsafeTypeAssertion<"col" | "bar">(
            unsafeTypeAssertion<string | undefined>(barDirNode?.["@_val"]) ?? "col",
          )
        : undefined;

    const holeSizeNode = unsafeTypeAssertion<XmlNode | undefined>(chartNode.holeSize);
    const holeSize = chartType === "doughnut" ? Number(holeSizeNode?.["@_val"] ?? 50) : undefined;

    const radarStyleNode = unsafeTypeAssertion<XmlNode | undefined>(chartNode.radarStyle);
    const radarStyle =
      chartType === "radar"
        ? unsafeTypeAssertion<"standard" | "marker" | "filled">(
            unsafeTypeAssertion<string | undefined>(radarStyleNode?.["@_val"]) ?? "standard",
          )
        : undefined;

    const ofPieTypeNode = unsafeTypeAssertion<XmlNode | undefined>(chartNode.ofPieType);
    const ofPieType =
      chartType === "ofPie"
        ? unsafeTypeAssertion<"pie" | "bar">(
            unsafeTypeAssertion<string | undefined>(ofPieTypeNode?.["@_val"]) ?? "pie",
          )
        : undefined;

    const secondPieSizeNode = unsafeTypeAssertion<XmlNode | undefined>(chartNode.secondPieSize);
    const secondPieSize =
      chartType === "ofPie" ? Number(secondPieSizeNode?.["@_val"] ?? 75) : undefined;

    const splitPosNode = unsafeTypeAssertion<XmlNode | undefined>(chartNode.splitPos);
    const splitPos = chartType === "ofPie" ? Number(splitPosNode?.["@_val"] ?? 2) : undefined;

    return {
      chartType,
      series,
      categories,
      barDirection,
      holeSize,
      radarStyle,
      ofPieType,
      secondPieSize,
      splitPos,
    };
  }

  return { chartType: null, series: [], categories: [] };
}

function parseSeries(
  ser: XmlNode,
  chartType: ChartType,
  seriesIndex: number,
  colorResolver: ColorResolver,
): ChartSeries {
  const name = parseSeriesName(unsafeTypeAssertion<XmlNode>(ser.tx));
  const usesXY = chartType === "scatter" || chartType === "bubble";
  const values = parseNumericData(
    usesXY ? unsafeTypeAssertion<XmlNode>(ser.yVal) : unsafeTypeAssertion<XmlNode>(ser.val),
  );
  const xValues = usesXY ? parseNumericData(unsafeTypeAssertion<XmlNode>(ser.xVal)) : undefined;
  const bubbleSizes =
    chartType === "bubble"
      ? parseNumericData(unsafeTypeAssertion<XmlNode>(ser.bubbleSize))
      : undefined;
  const color = resolveSeriesColor(
    unsafeTypeAssertion<XmlNode>(ser.spPr),
    seriesIndex,
    colorResolver,
  );

  return {
    name,
    values,
    ...(xValues !== undefined && { xValues }),
    ...(bubbleSizes !== undefined && { bubbleSizes }),
    color,
  };
}

function parseSeriesName(tx: XmlNode): string | null {
  if (!tx) return null;
  const strRef = unsafeTypeAssertion<XmlNode | undefined>(tx.strRef);
  const strCache = unsafeTypeAssertion<XmlNode | undefined>(strRef?.strCache);
  if (strCache?.pt) {
    const pts = unsafeTypeAssertion<XmlNode[]>(strCache.pt);
    return unsafeTypeAssertion<string | undefined>(pts[0]?.v) ?? null;
  }
  if (typeof tx.v === "string") return tx.v;
  return null;
}

function parseNumericData(valNode: XmlNode): number[] {
  if (!valNode) return [];
  const numRef = unsafeTypeAssertion<XmlNode | undefined>(valNode.numRef);
  const numCache = unsafeTypeAssertion<XmlNode | undefined>(numRef?.numCache);
  if (!numCache?.pt) return [];
  const pts = unsafeTypeAssertion<XmlNode[]>(numCache.pt);

  return pts
    .slice()
    .sort((a: XmlNode, b: XmlNode) => Number(a["@_idx"]) - Number(b["@_idx"]))
    .map((pt: XmlNode) => Number(pt.v ?? 0));
}

function extractCategories(serList: XmlNode[]): string[] {
  for (const ser of serList) {
    const cat = unsafeTypeAssertion<XmlNode | undefined>(ser.cat);
    if (!cat) continue;
    const strRef = unsafeTypeAssertion<XmlNode | undefined>(cat.strRef);
    const numRef = unsafeTypeAssertion<XmlNode | undefined>(cat.numRef);
    const multiLvlStrRef = unsafeTypeAssertion<XmlNode | undefined>(cat.multiLvlStrRef);
    const strCache =
      unsafeTypeAssertion<XmlNode | undefined>(strRef?.strCache) ??
      unsafeTypeAssertion<XmlNode | undefined>(numRef?.numCache);
    if (strCache?.pt) {
      return unsafeTypeAssertion<XmlNode[]>(strCache.pt)
        .slice()
        .sort((a: XmlNode, b: XmlNode) => Number(a["@_idx"]) - Number(b["@_idx"]))
        .map((pt: XmlNode) => String(unsafeTypeAssertion<string | number | undefined>(pt.v) ?? ""));
    }
    // multiLvlStrRef: categories stored in lvl > pt structure (uses first level only)
    const multiCache = unsafeTypeAssertion<XmlNode | undefined>(multiLvlStrRef?.multiLvlStrCache);
    if (multiCache?.lvl) {
      const lvls = Array.isArray(multiCache.lvl)
        ? unsafeTypeAssertion<XmlNode[]>(multiCache.lvl)
        : [unsafeTypeAssertion<XmlNode>(multiCache.lvl)];
      const firstLvl = lvls[0];
      if (firstLvl?.pt) {
        return unsafeTypeAssertion<XmlNode[]>(firstLvl.pt)
          .slice()
          .sort((a: XmlNode, b: XmlNode) => Number(a["@_idx"]) - Number(b["@_idx"]))
          .map((pt: XmlNode) =>
            String(unsafeTypeAssertion<string | number | undefined>(pt.v) ?? ""),
          );
      }
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
    const resolved = colorResolver.resolve(
      unsafeTypeAssertion<XmlNode | undefined>(spPr.solidFill) ?? spPr,
    );
    if (resolved) return resolved;
  }
  // Default: use theme accent colors
  const accentKey = ACCENT_KEYS[seriesIndex % ACCENT_KEYS.length];
  const resolved = colorResolver.resolve({ schemeClr: { "@_val": accentKey } });
  return resolved ?? { hex: "#4472C4", alpha: 1 };
}

function parseChartTitle(titleNode: XmlNode): string | null {
  if (!titleNode) return null;
  const tx = unsafeTypeAssertion<XmlNode | undefined>(titleNode.tx);
  const rich = unsafeTypeAssertion<XmlNode | undefined>(tx?.rich);
  if (!rich?.p) return null;
  const pList = Array.isArray(rich.p)
    ? unsafeTypeAssertion<XmlNode[]>(rich.p)
    : [unsafeTypeAssertion<XmlNode>(rich.p)];
  const texts: string[] = [];
  for (const p of pList) {
    const rList = unsafeTypeAssertion<XmlNode[] | undefined>(p.r) ?? [];
    for (const r of rList) {
      const t = r.t;
      if (typeof t === "string") texts.push(t);
      else if (t && unsafeTypeAssertion<XmlNode>(t)["#text"])
        texts.push(String(unsafeTypeAssertion<XmlNode>(t)["#text"]));
    }
  }
  return texts.length > 0 ? texts.join("") : null;
}

function parseLegend(legendNode: XmlNode): ChartLegend | null {
  if (!legendNode) return null;
  const legendPos = unsafeTypeAssertion<XmlNode | undefined>(legendNode.legendPos);
  const pos = unsafeTypeAssertion<ChartLegend["position"]>(
    unsafeTypeAssertion<string | undefined>(legendPos?.["@_val"]) ?? "r",
  );
  return { position: pos };
}
