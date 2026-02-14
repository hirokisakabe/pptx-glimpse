import type { GradientFill, SolidFill } from "./fill.js";

export type ArrowType = "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow";
export type ArrowSize = "sm" | "med" | "lg";

export interface ArrowEndpoint {
  type: ArrowType;
  width: ArrowSize;
  length: ArrowSize;
}

export type LineCap = "butt" | "round" | "square";
export type LineJoin = "miter" | "round" | "bevel";

export interface Outline {
  width: number;
  fill: SolidFill | GradientFill | null;
  dashStyle: DashStyle;
  customDash?: number[];
  lineCap?: LineCap;
  lineJoin?: LineJoin;
  headEnd: ArrowEndpoint | null;
  tailEnd: ArrowEndpoint | null;
}

export type DashStyle =
  | "solid"
  | "dash"
  | "dot"
  | "dashDot"
  | "lgDash"
  | "lgDashDot"
  | "sysDash"
  | "sysDot";
