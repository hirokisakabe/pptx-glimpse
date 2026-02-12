import type { Fill } from "./fill.js";
import type { Outline } from "./line.js";
import type { TextBody } from "./text.js";
import type { EffectList } from "./effect.js";

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

export interface CustomGeometryPath {
  width: number;
  height: number;
  commands: string;
}

export interface CustomGeometry {
  type: "custom";
  paths: CustomGeometryPath[];
}

export interface ShapeElement {
  type: "shape";
  transform: Transform;
  geometry: Geometry;
  fill: Fill | null;
  outline: Outline | null;
  textBody: TextBody | null;
  effects: EffectList | null;
  placeholderType?: string;
  placeholderIdx?: number;
}

export interface ConnectorElement {
  type: "connector";
  transform: Transform;
  geometry: Geometry;
  outline: Outline | null;
  effects: EffectList | null;
}

export interface GroupElement {
  type: "group";
  transform: Transform;
  childTransform: Transform;
  children: SlideElement[];
  effects: EffectList | null;
}

export type SlideElement =
  | ShapeElement
  | ImageElement
  | ConnectorElement
  | GroupElement
  | ChartElement
  | TableElement;

import type { ImageElement } from "./image.js";
import type { ChartElement } from "./chart.js";
import type { TableElement } from "./table.js";
