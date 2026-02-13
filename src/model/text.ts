import type { ResolvedColor } from "./fill.js";

/** defRPr に対応するデフォルトランプロパティ */
export interface DefaultRunProperties {
  fontSize?: number;
  fontFamily?: string | null;
  fontFamilyEa?: string | null;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
}

/** defaultTextStyle の各レベルに対応するデフォルト段落プロパティ */
export interface DefaultParagraphLevelProperties {
  alignment?: "l" | "ctr" | "r" | "just";
  marginLeft?: number;
  indent?: number;
  defaultRunProperties?: DefaultRunProperties;
}

/** presentation.xml の defaultTextStyle / slideMaster の titleStyle・bodyStyle・otherStyle */
export interface DefaultTextStyle {
  defaultParagraph?: DefaultParagraphLevelProperties;
  levels: (DefaultParagraphLevelProperties | undefined)[]; // index 0 = lvl1pPr, ... index 8 = lvl9pPr
}

/** slideMaster の txStyles */
export interface TxStyles {
  titleStyle?: DefaultTextStyle;
  bodyStyle?: DefaultTextStyle;
  otherStyle?: DefaultTextStyle;
}

/** プレースホルダーに紐づくテキストスタイル情報 */
export interface PlaceholderStyleInfo {
  placeholderType: string;
  placeholderIdx?: number;
  lstStyle?: DefaultTextStyle;
}

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
  numCol: number;
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

/** 段落間隔の値（ポイント指定またはパーセント指定） */
export type SpacingValue =
  | { type: "pts"; value: number } // 1/100 ポイント単位 (spcPts)
  | { type: "pct"; value: number }; // 1/1000 パーセント単位 (spcPct, 50000 = 50%)

/** タブストップ定義 */
export interface TabStop {
  position: number; // EMU
  alignment: "l" | "ctr" | "r" | "dec";
}

export interface ParagraphProperties {
  alignment: "l" | "ctr" | "r" | "just";
  lineSpacing: number | null;
  spaceBefore: SpacingValue;
  spaceAfter: SpacingValue;
  level: number;
  bullet: BulletType | null;
  bulletFont: string | null;
  bulletColor: ResolvedColor | null;
  bulletSizePct: number | null;
  marginLeft: number;
  indent: number;
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

export interface RunProperties {
  fontSize: number | null;
  fontFamily: string | null;
  fontFamilyEa: string | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  color: ResolvedColor | null;
  baseline: number;
  hyperlink: Hyperlink | null;
}
