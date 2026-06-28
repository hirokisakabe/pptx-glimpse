import type { ChartData, ChartLegend, ChartSeries, ChartType } from "@pptx-glimpse/renderer";
import type { ResolvedColor } from "@pptx-glimpse/renderer";

import type { ColorResolver } from "./color/color-resolver.js";
import { parseXml, type XmlNode } from "./ooxml/xml-parser.js";
import { unsafeXmlBoundaryAssertion } from "./unsafe-type-assertion.js";

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
  const chartSpace = unsafeXmlBoundaryAssertion<XmlNode | undefined>(parsed.chartSpace);
  if (!chartSpace) return null;

  const chart = unsafeXmlBoundaryAssertion<XmlNode | undefined>(chartSpace.chart);
  if (!chart) return null;

  const plotArea = unsafeXmlBoundaryAssertion<XmlNode | undefined>(chart.plotArea);
  if (!plotArea) return null;

  const title = parseChartTitle(unsafeXmlBoundaryAssertion<XmlNode>(chart.title));
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

  const legend = parseLegend(unsafeXmlBoundaryAssertion<XmlNode>(chart.legend));

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
    const chartNode = unsafeXmlBoundaryAssertion<XmlNode | undefined>(plotArea[xmlTag]);
    if (!chartNode) continue;

    const serList = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(chartNode.ser) ?? [];
    const series = serList.map((ser: XmlNode, index: number) =>
      parseSeries(ser, chartType, index, colorResolver),
    );
    const categories = extractCategories(serList);
    const barDirNode = unsafeXmlBoundaryAssertion<XmlNode | undefined>(chartNode.barDir);
    const barDirection =
      chartType === "bar"
        ? unsafeXmlBoundaryAssertion<"col" | "bar">(
            unsafeXmlBoundaryAssertion<string | undefined>(barDirNode?.["@_val"]) ?? "col",
          )
        : undefined;

    const holeSizeNode = unsafeXmlBoundaryAssertion<XmlNode | undefined>(chartNode.holeSize);
    const holeSize = chartType === "doughnut" ? Number(holeSizeNode?.["@_val"] ?? 50) : undefined;

    const radarStyleNode = unsafeXmlBoundaryAssertion<XmlNode | undefined>(chartNode.radarStyle);
    const radarStyle =
      chartType === "radar"
        ? unsafeXmlBoundaryAssertion<"standard" | "marker" | "filled">(
            unsafeXmlBoundaryAssertion<string | undefined>(radarStyleNode?.["@_val"]) ?? "standard",
          )
        : undefined;

    const ofPieTypeNode = unsafeXmlBoundaryAssertion<XmlNode | undefined>(chartNode.ofPieType);
    const ofPieType =
      chartType === "ofPie"
        ? unsafeXmlBoundaryAssertion<"pie" | "bar">(
            unsafeXmlBoundaryAssertion<string | undefined>(ofPieTypeNode?.["@_val"]) ?? "pie",
          )
        : undefined;

    const secondPieSizeNode = unsafeXmlBoundaryAssertion<XmlNode | undefined>(
      chartNode.secondPieSize,
    );
    const secondPieSize =
      chartType === "ofPie" ? Number(secondPieSizeNode?.["@_val"] ?? 75) : undefined;

    const splitPosNode = unsafeXmlBoundaryAssertion<XmlNode | undefined>(chartNode.splitPos);
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
  const name = parseSeriesName(unsafeXmlBoundaryAssertion<XmlNode>(ser.tx));
  const usesXY = chartType === "scatter" || chartType === "bubble";
  const values = parseNumericData(
    usesXY
      ? unsafeXmlBoundaryAssertion<XmlNode>(ser.yVal)
      : unsafeXmlBoundaryAssertion<XmlNode>(ser.val),
  );
  const xValues = usesXY
    ? parseNumericData(unsafeXmlBoundaryAssertion<XmlNode>(ser.xVal))
    : undefined;
  const bubbleSizes =
    chartType === "bubble"
      ? parseNumericData(unsafeXmlBoundaryAssertion<XmlNode>(ser.bubbleSize))
      : undefined;
  const color = resolveSeriesColor(
    unsafeXmlBoundaryAssertion<XmlNode>(ser.spPr),
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
  const strRef = unsafeXmlBoundaryAssertion<XmlNode | undefined>(tx.strRef);
  const strCache = unsafeXmlBoundaryAssertion<XmlNode | undefined>(strRef?.strCache);
  if (strCache?.pt) {
    const pts = unsafeXmlBoundaryAssertion<XmlNode[]>(strCache.pt);
    return unsafeXmlBoundaryAssertion<string | undefined>(pts[0]?.v) ?? null;
  }
  if (typeof tx.v === "string") return tx.v;
  return null;
}

function parseNumericData(valNode: XmlNode): number[] {
  if (!valNode) return [];
  const numRef = unsafeXmlBoundaryAssertion<XmlNode | undefined>(valNode.numRef);
  const numCache = unsafeXmlBoundaryAssertion<XmlNode | undefined>(numRef?.numCache);
  if (!numCache?.pt) return [];
  const pts = unsafeXmlBoundaryAssertion<XmlNode[]>(numCache.pt);

  return pts
    .slice()
    .sort((a: XmlNode, b: XmlNode) => Number(a["@_idx"]) - Number(b["@_idx"]))
    .map((pt: XmlNode) => Number(pt.v ?? 0));
}

function extractCategories(serList: XmlNode[]): string[] {
  for (const ser of serList) {
    const cat = unsafeXmlBoundaryAssertion<XmlNode | undefined>(ser.cat);
    if (!cat) continue;
    const strRef = unsafeXmlBoundaryAssertion<XmlNode | undefined>(cat.strRef);
    const numRef = unsafeXmlBoundaryAssertion<XmlNode | undefined>(cat.numRef);
    const multiLvlStrRef = unsafeXmlBoundaryAssertion<XmlNode | undefined>(cat.multiLvlStrRef);
    const strCache =
      unsafeXmlBoundaryAssertion<XmlNode | undefined>(strRef?.strCache) ??
      unsafeXmlBoundaryAssertion<XmlNode | undefined>(numRef?.numCache);
    if (strCache?.pt) {
      return unsafeXmlBoundaryAssertion<XmlNode[]>(strCache.pt)
        .slice()
        .sort((a: XmlNode, b: XmlNode) => Number(a["@_idx"]) - Number(b["@_idx"]))
        .map((pt: XmlNode) =>
          String(unsafeXmlBoundaryAssertion<string | number | undefined>(pt.v) ?? ""),
        );
    }
    // multiLvlStrRef: categories stored in lvl > pt structure (uses first level only)
    const multiCache = unsafeXmlBoundaryAssertion<XmlNode | undefined>(
      multiLvlStrRef?.multiLvlStrCache,
    );
    if (multiCache?.lvl) {
      const lvls = Array.isArray(multiCache.lvl)
        ? unsafeXmlBoundaryAssertion<XmlNode[]>(multiCache.lvl)
        : [unsafeXmlBoundaryAssertion<XmlNode>(multiCache.lvl)];
      const firstLvl = lvls[0];
      if (firstLvl?.pt) {
        return unsafeXmlBoundaryAssertion<XmlNode[]>(firstLvl.pt)
          .slice()
          .sort((a: XmlNode, b: XmlNode) => Number(a["@_idx"]) - Number(b["@_idx"]))
          .map((pt: XmlNode) =>
            String(unsafeXmlBoundaryAssertion<string | number | undefined>(pt.v) ?? ""),
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
      unsafeXmlBoundaryAssertion<XmlNode | undefined>(spPr.solidFill) ?? spPr,
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
  const tx = unsafeXmlBoundaryAssertion<XmlNode | undefined>(titleNode.tx);
  const rich = unsafeXmlBoundaryAssertion<XmlNode | undefined>(tx?.rich);
  if (!rich?.p) return null;
  const pList = Array.isArray(rich.p)
    ? unsafeXmlBoundaryAssertion<XmlNode[]>(rich.p)
    : [unsafeXmlBoundaryAssertion<XmlNode>(rich.p)];
  const texts: string[] = [];
  for (const p of pList) {
    const rList = unsafeXmlBoundaryAssertion<XmlNode[] | undefined>(p.r) ?? [];
    for (const r of rList) {
      const t = r.t;
      if (typeof t === "string") texts.push(t);
      else if (t && unsafeXmlBoundaryAssertion<XmlNode>(t)["#text"])
        texts.push(String(unsafeXmlBoundaryAssertion<XmlNode>(t)["#text"]));
    }
  }
  return texts.length > 0 ? texts.join("") : null;
}

function parseLegend(legendNode: XmlNode): ChartLegend | null {
  if (!legendNode) return null;
  const legendPos = unsafeXmlBoundaryAssertion<XmlNode | undefined>(legendNode.legendPos);
  const pos = unsafeXmlBoundaryAssertion<ChartLegend["position"]>(
    unsafeXmlBoundaryAssertion<string | undefined>(legendPos?.["@_val"]) ?? "r",
  );
  return { position: pos };
}
