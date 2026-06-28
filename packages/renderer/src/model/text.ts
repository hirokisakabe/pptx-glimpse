import type { Emu, HundredthPt, Pt } from "../utils/unit-types.js";
import type { ResolvedColor } from "./fill.js";
import type { Geometry, Transform } from "./shape.js";

/** Default run properties corresponding to defRPr */
export interface DefaultRunProperties {
  fontSize?: Pt;
  fontFamily?: string | null;
  fontFamilyEa?: string | null;
  fontFamilyCs?: string | null;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: ResolvedColor;
}

/** Default paragraph properties for each level of defaultTextStyle */
export interface DefaultParagraphLevelProperties {
  alignment?: "l" | "ctr" | "r" | "just";
  marginLeft?: Emu;
  indent?: Emu;
  bullet?: BulletType;
  bulletFont?: string;
  bulletColor?: ResolvedColor;
  bulletSizePct?: number;
  defaultRunProperties?: DefaultRunProperties;
}

/** defaultTextStyle of presentation.xml / titleStyle, bodyStyle, otherStyle of slideMaster */
export interface DefaultTextStyle {
  defaultParagraph?: DefaultParagraphLevelProperties;
  levels: (DefaultParagraphLevelProperties | undefined)[]; // index 0 = lvl1pPr, ... index 8 = lvl9pPr
}

/** slideMaster's txStyles */
export interface TxStyles {
  titleStyle?: DefaultTextStyle;
  bodyStyle?: DefaultTextStyle;
  otherStyle?: DefaultTextStyle;
}

/** Style information associated with a placeholder(text, position, and geometry) */
export interface PlaceholderStyleInfo {
  placeholderType: string;
  placeholderIdx?: number;
  lstStyle?: DefaultTextStyle;
  transform?: Transform;
  geometry?: Geometry;
}

export interface TextBody {
  paragraphs: Paragraph[];
  bodyProperties: BodyProperties;
}

export type TextVerticalType =
  | "horz"
  | "vert"
  | "vert270"
  | "eaVert"
  | "wordArtVert"
  | "mongolianVert";

export interface BodyProperties {
  anchor: "t" | "ctr" | "b";
  marginLeft: Emu;
  marginRight: Emu;
  marginTop: Emu;
  marginBottom: Emu;
  wrap: "square" | "none";
  autoFit: "noAutofit" | "normAutofit" | "spAutofit";
  fontScale: number;
  lnSpcReduction: number;
  numCol: number;
  vert: TextVerticalType;
}

export interface Paragraph {
  runs: TextRun[];
  properties: ParagraphProperties;
  endParaRunProperties?: RunProperties;
}

export type AutoNumScheme =
  | "arabicPeriod"
  | "arabicParenR"
  | "romanUcPeriod"
  | "romanLcPeriod"
  | "alphaUcPeriod"
  | "alphaLcPeriod"
  | "alphaLcParenR"
  | "alphaUcParenR"
  | "arabicPlain";

export type BulletType =
  | { type: "none" }
  | { type: "char"; char: string }
  | { type: "autoNum"; scheme: AutoNumScheme; startAt: number };

/** Paragraph spacing value (points or percentage) */
export type SpacingValue =
  | { type: "pts"; value: HundredthPt } // 1/100 point unit (spcPts)
  | { type: "pct"; value: number }; // 1/1000 percent (spcPct, 50000 = 50%)

/** Tab stop definition */
export interface TabStop {
  position: Emu;
  alignment: "l" | "ctr" | "r" | "dec";
}

export interface ParagraphProperties {
  alignment: "l" | "ctr" | "r" | "just" | null;
  lineSpacing: SpacingValue | null;
  spaceBefore: SpacingValue;
  spaceAfter: SpacingValue;
  level: number;
  bullet: BulletType | null;
  bulletFont: string | null;
  bulletColor: ResolvedColor | null;
  bulletSizePct: number | null;
  marginLeft: Emu | null;
  indent: Emu | null;
  tabStops: TabStop[];
}

export interface TextRun {
  text: string;
  properties: RunProperties;
}

export interface Hyperlink {
  url: string;
  tooltip?: string;
}

export interface TextOutline {
  width: Emu;
  color: ResolvedColor;
}

export interface RunProperties {
  fontSize: Pt | null;
  fontFamily: string | null;
  fontFamilyEa: string | null;
  fontFamilyCs: string | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  color: ResolvedColor | null;
  baseline: number;
  hyperlink: Hyperlink | null;
  outline: TextOutline | null;
}
