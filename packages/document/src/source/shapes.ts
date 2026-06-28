/**
 * Simple shape / text run / image  source node types.
 *
 * Drawings centered around the current supported subset (simple shapes / text / images
 * nodes) to typed. theme color and relationship are unresolved references in source
 * cascade resolution is the responsibility of the computed view. Unsupported nodes are raw
 * Save with escape hatch.
 */

import type { RelationshipId, SourceHandle, SourceNodeId } from "./handles.js";
import type { RawSidecar } from "./raw.js";
import type { Emu, HundredthPt, OoxmlAngle, OoxmlPercent, Pt } from "./units.js";

/** Source transform from `a:xfrm`. Keep the coordinates and size as EMU. */
export interface SourceTransform {
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly rotation?: OoxmlAngle;
  readonly flipHorizontal?: boolean;
  readonly flipVertical?: boolean;
}

/** Source reference for preset geometry (`a:prstGeom`). Keep only the preset name. */
export interface SourcePresetGeometry {
  /** Preset name (Example: `rect`, `roundRect`). */
  readonly preset: string;
  readonly adjustValues?: Readonly<Record<string, number>>;
}

export interface SourceCustomGeometryPath {
  readonly width: number;
  readonly height: number;
  readonly commands: string;
}

export interface SourceCustomGeometry {
  readonly kind: "custom";
  readonly paths: readonly SourceCustomGeometryPath[];
}

export type SourceGeometry = SourcePresetGeometry | SourceCustomGeometry;

/**
 * Unresolved source color reference. `schemeClr` is kept as the scheme name and computed
 * Resolve to concrete color via `clrMap` / `colorScheme` in view. `srgbClr` /
 * `sysClr` is a concrete value, but any transformations (lumMod, etc.) are retained before being applied.
 */
export type SourceColor =
  | {
      readonly kind: "srgb";
      readonly hex: string;
      readonly transforms?: readonly SourceColorTransform[];
    }
  | {
      readonly kind: "scheme";
      readonly scheme: string;
      readonly transforms?: readonly SourceColorTransform[];
    }
  | {
      readonly kind: "system";
      /** `a:sysClr@val` (Example: `windowText`). */
      readonly value: string;
      /** Resolved fallback hex for `a:sysClr@lastClr`, if it has one. */
      readonly lastColor?: string;
      readonly transforms?: readonly SourceColorTransform[];
    };

/** Color conversions such as lumMod / tint / shade (values are OOXML percentages). */
export interface SourceColorTransform {
  readonly kind: "lumMod" | "lumOff" | "tint" | "shade" | "alpha";
  readonly value: OoxmlPercent;
}

/** shape fill (source level, minimum). Unsupported fills are saved as raw. */
export type SourceFill =
  | { readonly kind: "none" }
  | { readonly kind: "solid"; readonly color: SourceColor }
  | ({ readonly kind: "gradient"; readonly stops: readonly SourceGradientStop[] } & SourceGradient)
  | {
      readonly kind: "pattern";
      readonly preset: string;
      readonly foregroundColor: SourceColor;
      readonly backgroundColor: SourceColor;
    }
  | {
      readonly kind: "image";
      readonly blipRelationshipId?: RelationshipId;
      readonly tile?: SourceImageFillTile;
    }
  | { readonly kind: "raw"; readonly raw: RawSidecar };

export interface SourceGradientStop {
  readonly position: number;
  readonly color: SourceColor;
}

export type SourceGradient =
  | {
      readonly gradientType: "linear";
      readonly angle: OoxmlAngle;
    }
  | {
      readonly gradientType: "radial";
      readonly centerX: number;
      readonly centerY: number;
    };

export type SourceRectangleAlignment = "tl" | "t" | "tr" | "l" | "ctr" | "r" | "bl" | "b" | "br";

export interface SourceImageFillTile {
  readonly tx: Emu;
  readonly ty: Emu;
  readonly sx: number;
  readonly sy: number;
  readonly flip: "none" | "x" | "y" | "xy";
  readonly align: SourceRectangleAlignment;
}

export interface SourceStyleReference {
  readonly index: number;
  readonly color?: SourceColor;
}

export interface SourceShapeStyle {
  readonly fillRef?: SourceStyleReference;
  readonly lineRef?: SourceStyleReference;
  readonly effectRef?: SourceStyleReference;
}

export interface SourceEffectList {
  readonly outerShadow?: SourceOuterShadow;
  readonly innerShadow?: SourceInnerShadow;
  readonly glow?: SourceGlow;
  readonly softEdge?: SourceSoftEdge;
}

export interface SourceOuterShadow {
  readonly blurRadius: Emu;
  readonly distance: Emu;
  readonly direction: OoxmlAngle;
  readonly color: SourceColor;
  readonly alignment: SourceRectangleAlignment;
  readonly rotateWithShape: boolean;
}

export interface SourceInnerShadow {
  readonly blurRadius: Emu;
  readonly distance: Emu;
  readonly direction: OoxmlAngle;
  readonly color: SourceColor;
}

export interface SourceGlow {
  readonly radius: Emu;
  readonly color: SourceColor;
}

export interface SourceSoftEdge {
  readonly radius: Emu;
}

export interface SourceBlipEffects {
  readonly grayscale: boolean;
  readonly biLevel?: SourceBiLevelEffect;
  readonly blur?: SourceBlurEffect;
  readonly lum?: SourceLumEffect;
  readonly duotone?: SourceDuotoneEffect;
  readonly clrChange?: SourceColorChangeEffect;
}

export interface SourceBiLevelEffect {
  readonly threshold: number;
}

export interface SourceBlurEffect {
  readonly radius: Emu;
  readonly grow: boolean;
}

export interface SourceLumEffect {
  readonly brightness: number;
  readonly contrast: number;
}

export interface SourceDuotoneEffect {
  readonly color1: SourceColor;
  readonly color2: SourceColor;
}

export interface SourceColorChangeEffect {
  readonly from: SourceColor;
  readonly to: SourceColor;
}

/** simple solid line outline (`a:ln`). Minimal representation of color and width only. */
export interface SourceOutline {
  readonly width?: Emu;
  readonly fill?: SourceFill;
  readonly dashStyle?: SourceDashStyle;
  readonly customDash?: readonly number[];
  readonly lineCap?: SourceLineCap;
  readonly lineJoin?: SourceLineJoin;
  readonly headEnd?: SourceArrowEndpoint;
  readonly tailEnd?: SourceArrowEndpoint;
}

export type SourceArrowType = "triangle" | "stealth" | "diamond" | "oval" | "arrow";
export type SourceArrowSize = "sm" | "med" | "lg";
export interface SourceArrowEndpoint {
  readonly type: SourceArrowType;
  readonly width: SourceArrowSize;
  readonly length: SourceArrowSize;
}
export type SourceDashStyle =
  | "solid"
  | "dash"
  | "dot"
  | "dashDot"
  | "lgDash"
  | "lgDashDot"
  | "sysDash"
  | "sysDot";
export type SourceLineCap = "butt" | "round" | "square";
export type SourceLineJoin = "miter" | "round" | "bevel";

/** placeholder declaration (`p:ph`). Keep type / idx unresolved. */
export interface SourcePlaceholder {
  readonly type?: string;
  readonly index?: number;
}

/** Minimal subset of run properties (`a:rPr`) of text run (`a:r`). */
export interface SourceRunProperties {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strikethrough?: boolean;
  readonly baseline?: number;
  /** font size. In OOXML-domain, it is stored as pt. */
  readonly fontSize?: Pt;
  /** latin typeface name (kept unresolved including theme token). */
  readonly typeface?: string;
  /** East Asian typeface name (kept unresolved including theme token). */
  readonly typefaceEa?: string;
  /** Complex script typeface name (kept unresolved including theme token). */
  readonly typefaceCs?: string;
  readonly color?: SourceColor;
}

export type SourceAutoNumScheme =
  | "arabicPeriod"
  | "arabicParenR"
  | "romanUcPeriod"
  | "romanLcPeriod"
  | "alphaUcPeriod"
  | "alphaLcPeriod"
  | "alphaLcParenR"
  | "alphaUcParenR"
  | "arabicPlain";

export type SourceBulletType =
  | { readonly type: "none" }
  | { readonly type: "char"; readonly char: string }
  | { readonly type: "autoNum"; readonly scheme: SourceAutoNumScheme; readonly startAt: number };

export type SourceSpacingValue =
  | { readonly type: "pts"; readonly value: HundredthPt }
  | { readonly type: "pct"; readonly value: number };

/** text run (`a:r`). */
export interface SourceTextRun {
  readonly kind: "textRun";
  readonly text: string;
  readonly properties?: SourceRunProperties;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

export type SourceTextAlign = "left" | "center" | "right" | "justify";

export interface SourceTabStop {
  readonly position: Emu;
  readonly alignment: "l" | "ctr" | "r" | "dec";
}

/** Minimum subset of paragraph properties (`a:pPr`) of paragraph (`a:p`). */
export interface SourceParagraphProperties {
  readonly align?: SourceTextAlign;
  /** indentation level (`a:pPr@lvl`). */
  readonly level?: number;
  readonly lineSpacing?: SourceSpacingValue;
  readonly spaceBefore?: SourceSpacingValue;
  readonly spaceAfter?: SourceSpacingValue;
  readonly marginLeft?: Emu;
  readonly indent?: Emu;
  readonly bullet?: SourceBulletType;
  readonly bulletFont?: string;
  readonly bulletColor?: SourceColor;
  readonly bulletSizePct?: number;
  readonly tabStops?: readonly SourceTabStop[];
  readonly defaultRunProperties?: SourceRunProperties;
}

/** paragraph (`a:p`). */
export interface SourceParagraph {
  readonly runs: readonly SourceTextRun[];
  readonly properties?: SourceParagraphProperties;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/** vertical anchor (`a:bodyPr@anchor`). */
export type SourceVerticalAnchor = "top" | "middle" | "bottom";

export type SourceTextWrap = "square" | "none";

export type SourceTextVerticalType =
  | "horz"
  | "vert"
  | "vert270"
  | "eaVert"
  | "wordArtVert"
  | "mongolianVert";

export type SourceTextAutoFit = "noAutofit" | "normAutofit" | "spAutofit";

/** Minimal subset of text body properties (`a:bodyPr`). inset and vertical anchor. */
export interface SourceTextBodyProperties {
  readonly marginLeft?: Emu;
  readonly marginRight?: Emu;
  readonly marginTop?: Emu;
  readonly marginBottom?: Emu;
  readonly anchor?: SourceVerticalAnchor;
  readonly wrap?: SourceTextWrap;
  readonly autoFit?: SourceTextAutoFit;
  readonly fontScale?: number;
  readonly lnSpcReduction?: number;
  readonly numCol?: number;
  readonly vert?: SourceTextVerticalType;
}

/** text body (`p:txBody`). */
export interface SourceTextBody {
  readonly paragraphs: readonly SourceParagraph[];
  readonly properties?: SourceTextBodyProperties;
  readonly listStyle?: SourceTextStyle;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

export interface SourceTextStyle {
  readonly defaultParagraph?: SourceParagraphProperties;
  readonly levels: readonly (SourceParagraphProperties | undefined)[];
}

/** simple shape (`p:sp`). */
export interface SourceShape {
  readonly kind: "shape";
  readonly nodeId?: SourceNodeId;
  readonly name?: string;
  readonly transform?: SourceTransform;
  readonly geometry?: SourceGeometry;
  readonly fill?: SourceFill;
  readonly outline?: SourceOutline;
  readonly effects?: SourceEffectList;
  readonly style?: SourceShapeStyle;
  readonly textBody?: SourceTextBody;
  readonly placeholder?: SourcePlaceholder;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

export interface SourceConnector {
  readonly kind: "connector";
  readonly nodeId?: SourceNodeId;
  readonly name?: string;
  readonly transform?: SourceTransform;
  readonly geometry?: SourceGeometry;
  readonly outline?: SourceOutline;
  readonly effects?: SourceEffectList;
  readonly style?: SourceShapeStyle;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

export interface SourceGroup {
  readonly kind: "group";
  readonly nodeId?: SourceNodeId;
  readonly name?: string;
  readonly transform?: SourceTransform;
  readonly childTransform?: SourceTransform;
  readonly fill?: SourceFill;
  readonly effects?: SourceEffectList;
  readonly children: readonly SourceShapeNode[];
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/** crop (`a:srcRect`) of image. Preserve insets on each edge as OOXML percentages. */
export interface SourceImageCrop {
  readonly left?: OoxmlPercent;
  readonly top?: OoxmlPercent;
  readonly right?: OoxmlPercent;
  readonly bottom?: OoxmlPercent;
}

export interface SourceImageStretch {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/** image (`p:pic`). The blip keeps the relationship id (`r:embed`) unresolved. */
export interface SourceImage {
  readonly kind: "image";
  readonly nodeId?: SourceNodeId;
  readonly name?: string;
  readonly transform?: SourceTransform;
  /** Relationship id of `a:blip@r:embed`. media part is solved by computed view. */
  readonly blipRelationshipId?: RelationshipId;
  readonly crop?: SourceImageCrop;
  readonly stretch?: SourceImageStretch;
  readonly tile?: SourceImageFillTile;
  readonly effects?: SourceEffectList;
  readonly blipEffects?: SourceBlipEffects;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

export interface SourceTable {
  readonly kind: "table";
  readonly nodeId?: SourceNodeId;
  readonly name?: string;
  readonly transform?: SourceTransform;
  readonly table: SourceTableData;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

export interface SourceTableData {
  readonly columns: readonly SourceTableColumn[];
  readonly rows: readonly SourceTableRow[];
  readonly tableStyleId?: string;
}

export interface SourceTableColumn {
  readonly width: Emu;
}

export interface SourceTableRow {
  readonly height: Emu;
  readonly cells: readonly SourceTableCell[];
}

export interface SourceTableCell {
  readonly textBody?: SourceTextBody;
  readonly fill?: SourceFill;
  readonly borders?: SourceCellBorders;
  readonly gridSpan: number;
  readonly rowSpan: number;
  readonly hMerge: boolean;
  readonly vMerge: boolean;
  readonly rawSidecars?: readonly RawSidecar[];
}

export interface SourceCellBorders {
  readonly top?: SourceOutline;
  readonly bottom?: SourceOutline;
  readonly left?: SourceOutline;
  readonly right?: SourceOutline;
}

export interface SourceChart {
  readonly kind: "chart";
  readonly nodeId?: SourceNodeId;
  readonly name?: string;
  readonly transform?: SourceTransform;
  /** Relationship id of `c:chart@r:id`. Chart parts are resolved with computed views. */
  readonly chartRelationshipId?: RelationshipId;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

export interface SourceSmartArt {
  readonly kind: "smartArt";
  readonly nodeId?: SourceNodeId;
  readonly name?: string;
  readonly transform?: SourceTransform;
  /** Relationship id of `dgm:relIds@r:dm`. Diagram data/drawing part is solved by computed view. */
  readonly dataRelationshipId?: RelationshipId;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/**
 * Raw escape hatch for shape tree nodes that are not represented as typed nodes. Unsupported shapes,
 * groups, connectors, and similar nodes are saved for round trips.
 */
export interface SourceRawShapeNode {
  readonly kind: "raw";
  readonly nodeId?: SourceNodeId;
  readonly raw: RawSidecar;
  readonly handle?: SourceHandle;
}

/** Shape tree source node union. */
export type SourceShapeNode =
  | SourceShape
  | SourceConnector
  | SourceGroup
  | SourceImage
  | SourceTable
  | SourceChart
  | SourceSmartArt
  | SourceRawShapeNode;
