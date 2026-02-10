import type { Fill } from "./fill.js";
import type { Outline } from "./line.js";
import type { TextBody } from "./text.js";

export interface Transform {
  offsetX: number;
  offsetY: number;
  extentWidth: number;
  extentHeight: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
}

export type Geometry = PresetGeometry | CustomGeometry;

export interface PresetGeometry {
  type: "preset";
  preset: string;
  adjustValues: Record<string, number>;
}

export interface CustomGeometry {
  type: "custom";
  pathData: string;
}

export interface ShapeElement {
  type: "shape";
  transform: Transform;
  geometry: Geometry;
  fill: Fill | null;
  outline: Outline | null;
  textBody: TextBody | null;
  placeholderType?: string;
  placeholderIdx?: number;
}

export interface ConnectorElement {
  type: "connector";
  transform: Transform;
  outline: Outline | null;
}

export interface GroupElement {
  type: "group";
  transform: Transform;
  childTransform: Transform;
  children: SlideElement[];
}

export type SlideElement =
  | ShapeElement
  | ImageElement
  | ConnectorElement
  | GroupElement
  | ChartElement;

import type { ImageElement } from "./image.js";
import type { ChartElement } from "./chart.js";
