import type { SolidFill } from "./fill.js";

export type ArrowType = "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow";
export type ArrowSize = "sm" | "med" | "lg";

export interface ArrowEndpoint {
  type: ArrowType;
  width: ArrowSize;
  length: ArrowSize;
}

export interface Outline {
  width: number;
  fill: SolidFill | null;
  dashStyle: DashStyle;
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
