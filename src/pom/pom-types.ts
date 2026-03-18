/**
 * Self-contained pom output types.
 * These mirror @hirokisakabe/pom's POMNode types so that pptx-glimpse
 * does not take a runtime dependency on pom.
 */

// ---------------------------------------------------------------------------
// Common
// ---------------------------------------------------------------------------

interface PomPadding {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
}

interface PomBorder {
  color?: string;
  width?: number;
  dashType?: PomDashType;
}

export type PomDashType =
  | "solid"
  | "dash"
  | "dashDot"
  | "lgDash"
  | "lgDashDot"
  | "lgDashDotDot"
  | "sysDash"
  | "sysDot";

export interface PomShadow {
  type?: "outer" | "inner";
  opacity?: number;
  blur?: number;
  angle?: number;
  offset?: number;
  color?: string;
}

export interface PomFill {
  color?: string;
  transparency?: number;
}

export type PomAlignText = "left" | "center" | "right";

export type PomAlignItems = "start" | "center" | "end" | "stretch";
export type PomJustifyContent =
  | "start"
  | "center"
  | "end"
  | "spaceBetween"
  | "spaceAround"
  | "spaceEvenly";

// ---------------------------------------------------------------------------
// Base properties shared by all nodes
// ---------------------------------------------------------------------------

interface PomBaseNode {
  w?: number | string;
  h?: number | string;
  minW?: number;
  maxW?: number;
  minH?: number;
  maxH?: number;
  padding?: number | PomPadding;
  backgroundColor?: string;
  border?: PomBorder;
  borderRadius?: number;
  opacity?: number;
}

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

export interface PomTextNode extends PomBaseNode {
  type: "text";
  text: string;
  fontPx?: number;
  color?: string;
  alignText?: PomAlignText;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  highlight?: string;
  fontFamily?: string;
  lineSpacingMultiple?: number;
}

// ---------------------------------------------------------------------------
// List
// ---------------------------------------------------------------------------

export interface PomLiNode {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  highlight?: string;
  color?: string;
  fontPx?: number;
  fontFamily?: string;
}

export interface PomUlNode extends PomBaseNode {
  type: "ul";
  items: PomLiNode[];
  fontPx?: number;
  color?: string;
  alignText?: PomAlignText;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  highlight?: string;
  fontFamily?: string;
  lineSpacingMultiple?: number;
}

export interface PomOlNode extends PomBaseNode {
  type: "ol";
  items: PomLiNode[];
  fontPx?: number;
  color?: string;
  alignText?: PomAlignText;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  highlight?: string;
  fontFamily?: string;
  lineSpacingMultiple?: number;
  numberType?: string;
  numberStartAt?: number;
}

// ---------------------------------------------------------------------------
// Image
// ---------------------------------------------------------------------------

export interface PomImageNode extends PomBaseNode {
  type: "image";
  src: string;
  sizing?: {
    type: "cover" | "contain" | "crop";
    w?: number;
    h?: number;
    x?: number;
    y?: number;
  };
  shadow?: PomShadow;
}

// ---------------------------------------------------------------------------
// Table
// ---------------------------------------------------------------------------

export interface PomTableCell {
  text: string;
  fontPx?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  highlight?: string;
  alignText?: PomAlignText;
  backgroundColor?: string;
  colspan?: number;
  rowspan?: number;
}

export interface PomTableRow {
  cells: PomTableCell[];
  height?: number;
}

export interface PomTableColumn {
  width?: number;
}

export interface PomTableNode extends PomBaseNode {
  type: "table";
  columns: PomTableColumn[];
  rows: PomTableRow[];
  defaultRowHeight?: number;
}

// ---------------------------------------------------------------------------
// Shape
// ---------------------------------------------------------------------------

export interface PomShapeNode extends PomBaseNode {
  type: "shape";
  shapeType: string;
  text?: string;
  fill?: PomFill;
  line?: {
    color?: string;
    width?: number;
    dashType?: PomDashType;
  };
  shadow?: PomShadow;
  fontPx?: number;
  color?: string;
  alignText?: PomAlignText;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  highlight?: string;
  fontFamily?: string;
  lineSpacingMultiple?: number;
}

// ---------------------------------------------------------------------------
// Chart
// ---------------------------------------------------------------------------

export type PomChartType = "bar" | "line" | "pie" | "area" | "doughnut" | "radar";

export interface PomChartData {
  name?: string;
  labels: string[];
  values: number[];
}

export interface PomChartNode extends PomBaseNode {
  type: "chart";
  chartType: PomChartType;
  data: PomChartData[];
  showLegend?: boolean;
  showTitle?: boolean;
  title?: string;
  chartColors?: string[];
  radarStyle?: "standard" | "marker" | "filled";
}

// ---------------------------------------------------------------------------
// Line
// ---------------------------------------------------------------------------

export type PomLineArrow =
  | boolean
  | { type?: "none" | "diamond" | "triangle" | "arrow" | "oval" | "stealth" };

export interface PomLineNode extends PomBaseNode {
  type: "line";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color?: string;
  lineWidth?: number;
  dashType?: PomDashType;
  beginArrow?: PomLineArrow;
  endArrow?: PomLineArrow;
}

// ---------------------------------------------------------------------------
// Layout containers
// ---------------------------------------------------------------------------

export interface PomBoxNode extends PomBaseNode {
  type: "box";
  children: PomNode;
  shadow?: PomShadow;
}

export interface PomVStackNode extends PomBaseNode {
  type: "vstack";
  children: PomNode[];
  gap?: number;
  alignItems?: PomAlignItems;
  justifyContent?: PomJustifyContent;
}

export interface PomHStackNode extends PomBaseNode {
  type: "hstack";
  children: PomNode[];
  gap?: number;
  alignItems?: PomAlignItems;
  justifyContent?: PomJustifyContent;
}

export type PomLayerChild = PomNode & {
  x: number;
  y: number;
};

export interface PomLayerNode extends PomBaseNode {
  type: "layer";
  children: PomLayerChild[];
}

// ---------------------------------------------------------------------------
// Union
// ---------------------------------------------------------------------------

export type PomNode =
  | PomTextNode
  | PomUlNode
  | PomOlNode
  | PomImageNode
  | PomTableNode
  | PomBoxNode
  | PomVStackNode
  | PomHStackNode
  | PomShapeNode
  | PomChartNode
  | PomLineNode
  | PomLayerNode;
