import type { ResolvedColor } from "./fill.js";
import type { Transform } from "./shape.js";

export interface ChartElement {
  type: "chart";
  transform: Transform;
  chart: ChartData;
}

export type ChartType =
  | "combo"
  | "bar"
  | "line"
  | "pie"
  | "doughnut"
  | "scatter"
  | "bubble"
  | "area"
  | "radar"
  | "stock"
  | "surface"
  | "ofPie";

export interface ChartData {
  chartType: ChartType;
  title: string | null;
  series: ChartSeries[];
  categories: string[];
  barDirection?: "col" | "bar";
  holeSize?: number;
  radarStyle?: "standard" | "marker" | "filled";
  ofPieType?: "pie" | "bar";
  secondPieSize?: number;
  splitPos?: number;
  legend: ChartLegend | null;
}

export interface ChartSeries {
  name: string | null;
  values: number[];
  xValues?: number[];
  bubbleSizes?: number[];
  color: ResolvedColor;
  /** Plot type for a series in a category combo chart. */
  chartType?: "bar" | "line";
}

export interface ChartLegend {
  position: "b" | "t" | "l" | "r" | "tr";
}
