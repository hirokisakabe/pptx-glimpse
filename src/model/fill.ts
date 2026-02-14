import type { Emu } from "../utils/unit-types.js";

export type Fill = SolidFill | GradientFill | ImageFill | PatternFill | NoFill;

export interface SolidFill {
  type: "solid";
  color: ResolvedColor;
}

export interface GradientFill {
  type: "gradient";
  stops: GradientStop[];
  angle: number;
  gradientType: "linear" | "radial";
  centerX?: number;
  centerY?: number;
}

export interface GradientStop {
  position: number;
  color: ResolvedColor;
}

export interface ImageFill {
  type: "image";
  imageData: string;
  mimeType: string;
  tile: ImageFillTile | null;
}

export interface ImageFillTile {
  tx: Emu;
  ty: Emu;
  sx: number;
  sy: number;
  flip: "none" | "x" | "y" | "xy";
  align: string;
}

export interface PatternFill {
  type: "pattern";
  preset: string;
  foregroundColor: ResolvedColor;
  backgroundColor: ResolvedColor;
}

export interface NoFill {
  type: "none";
}

export interface ResolvedColor {
  hex: string;
  alpha: number;
}
