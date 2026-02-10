import type { Transform } from "./shape.js";
import type { ResolvedColor } from "./fill.js";

export interface ChartElement {
  type: "chart";
  transform: Transform;
  chart: ChartData;
}

export type ChartType = "bar" | "line" | "pie" | "scatter";

export interface ChartData {
  chartType: ChartType;
  title: string | null;
  series: ChartSeries[];
  categories: string[];
  barDirection?: "col" | "bar";
  legend: ChartLegend | null;
}

export interface ChartSeries {
  name: string | null;
  values: number[];
  xValues?: number[];
  color: ResolvedColor;
}

export interface ChartLegend {
  position: "b" | "t" | "l" | "r" | "tr";
}
