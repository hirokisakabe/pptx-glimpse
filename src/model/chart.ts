import type { Transform } from "./shape.js";
import type { ResolvedColor } from "./fill.js";

export interface ChartElement {
  type: "chart";
  transform: Transform;
  chart: ChartData;
}

export type ChartType =
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
}

export interface ChartLegend {
  position: "b" | "t" | "l" | "r" | "tr";
}
