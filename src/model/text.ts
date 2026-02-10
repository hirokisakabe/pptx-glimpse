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

export interface ParagraphProperties {
  alignment: "l" | "ctr" | "r" | "just";
  lineSpacing: number | null;
  spaceBefore: number;
  spaceAfter: number;
  level: number;
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
}
