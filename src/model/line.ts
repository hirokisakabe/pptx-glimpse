import type { SolidFill } from "./fill.js";

export interface Outline {
  width: number;
  fill: SolidFill | null;
  dashStyle: DashStyle;
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
