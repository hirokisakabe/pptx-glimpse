import type { ResolvedColor } from "./fill.js";

export interface TextBody {
  paragraphs: Paragraph[];
  bodyProperties: BodyProperties;
}

export interface BodyProperties {
  anchor: "t" | "ctr" | "b";
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  wrap: "square" | "none";
  autoFit: "noAutofit" | "normAutofit" | "spAutofit";
  fontScale: number;
  lnSpcReduction: number;
}

export interface Paragraph {
  runs: TextRun[];
  properties: ParagraphProperties;
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

export interface ParagraphProperties {
  alignment: "l" | "ctr" | "r" | "just";
  lineSpacing: number | null;
  spaceBefore: number;
  spaceAfter: number;
  level: number;
  bullet: BulletType | null;
  bulletFont: string | null;
  bulletColor: ResolvedColor | null;
  bulletSizePct: number | null;
  marginLeft: number;
  indent: number;
}

export interface TextRun {
  text: string;
  properties: RunProperties;
}

export interface RunProperties {
  fontSize: number | null;
  fontFamily: string | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  color: ResolvedColor | null;
  baseline: number;
}
