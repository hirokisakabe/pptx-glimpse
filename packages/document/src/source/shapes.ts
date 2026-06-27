/**
 * Simple shape / text run / image の source node 型。
 *
 * PoC scope の最小 subset (simple shapes / text / images) のみを typed に表す。
 * theme color や relationship は source では未解決の参照のまま保持し
 * (`docs/cleandoc-source-computed-view.md`)、cascade 解決は computed view の
 * 責務とする。未対応ノードは raw escape hatch で保存する。
 */

import type { RelationshipId, SourceHandle, SourceNodeId } from "./handles.js";
import type { RawSidecar } from "./raw.js";
import type { Emu, HundredthPt, OoxmlAngle, OoxmlPercent, Pt } from "./units.js";

/** `a:xfrm` 由来の source transform。座標・サイズは EMU のまま保持する。 */
export interface SourceTransform {
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
  readonly rotation?: OoxmlAngle;
  readonly flipHorizontal?: boolean;
  readonly flipVertical?: boolean;
}

/** preset geometry (`a:prstGeom`) の source 参照。preset 名のみ保持する。 */
export interface SourcePresetGeometry {
  /** preset 名 (例: `rect`, `roundRect`)。 */
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
 * 未解決の source color 参照。`schemeClr` は scheme 名のまま保持し、computed
 * view で `clrMap` / `colorScheme` を経由して具体色へ解決する。`srgbClr` /
 * `sysClr` は具体値だが、変換 (lumMod 等) は適用前のまま保持する。
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
      /** `a:sysClr@val` (例: `windowText`)。 */
      readonly value: string;
      /** `a:sysClr@lastClr` の解決済みフォールバック hex (持つ場合)。 */
      readonly lastColor?: string;
      readonly transforms?: readonly SourceColorTransform[];
    };

/** lumMod / tint / shade 等の色変換 (値は OOXML パーセンテージ)。 */
export interface SourceColorTransform {
  readonly kind: "lumMod" | "lumOff" | "tint" | "shade" | "alpha";
  readonly value: OoxmlPercent;
}

/** shape の fill (source レベル、最小)。未対応 fill は raw で保存する。 */
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

export interface SourceImageFillTile {
  readonly tx: Emu;
  readonly ty: Emu;
  readonly sx: number;
  readonly sy: number;
  readonly flip: "none" | "x" | "y" | "xy";
  readonly align: string;
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
  readonly alignment: string;
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

/** simple solid line の outline (`a:ln`)。色と幅のみの最小表現。 */
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

/** placeholder 宣言 (`p:ph`)。type / idx を未解決のまま保持する。 */
export interface SourcePlaceholder {
  readonly type?: string;
  readonly index?: number;
}

/** text run (`a:r`) の run プロパティ (`a:rPr`) の最小 subset。 */
export interface SourceRunProperties {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strikethrough?: boolean;
  readonly baseline?: number;
  /** フォントサイズ。OOXML-domain では pt で保持する。 */
  readonly fontSize?: Pt;
  /** latin typeface 名 (theme token も含めて未解決のまま保持)。 */
  readonly typeface?: string;
  /** East Asian typeface 名 (theme token も含めて未解決のまま保持)。 */
  readonly typefaceEa?: string;
  /** Complex script typeface 名 (theme token も含めて未解決のまま保持)。 */
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

/** text run (`a:r`)。 */
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

/** paragraph (`a:p`) の段落プロパティ (`a:pPr`) の最小 subset。 */
export interface SourceParagraphProperties {
  readonly align?: SourceTextAlign;
  /** インデントレベル (`a:pPr@lvl`)。 */
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

/** paragraph (`a:p`)。 */
export interface SourceParagraph {
  readonly runs: readonly SourceTextRun[];
  readonly properties?: SourceParagraphProperties;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/** vertical anchor (`a:bodyPr@anchor`)。 */
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

/** text body properties (`a:bodyPr`) の最小 subset。inset と vertical anchor。 */
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

/** text body (`p:txBody`)。 */
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

/** simple shape (`p:sp`)。 */
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
  readonly effects?: SourceEffectList;
  readonly children: readonly SourceShapeNode[];
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/** image の crop (`a:srcRect`)。各辺の inset を OOXML パーセンテージで保持する。 */
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

/** image (`p:pic`)。blip は relationship id (`r:embed`) を未解決のまま保持する。 */
export interface SourceImage {
  readonly kind: "image";
  readonly nodeId?: SourceNodeId;
  readonly name?: string;
  readonly transform?: SourceTransform;
  /** `a:blip@r:embed` の relationship id。media part は computed view で解決する。 */
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
  /** `c:chart@r:id` の relationship id。chart part は computed view で解決する。 */
  readonly chartRelationshipId?: RelationshipId;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

export interface SourceSmartArt {
  readonly kind: "smartArt";
  readonly nodeId?: SourceNodeId;
  readonly name?: string;
  readonly transform?: SourceTransform;
  /** `dgm:relIds@r:dm` の relationship id。diagram data/drawing part は computed view で解決する。 */
  readonly dataRelationshipId?: RelationshipId;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/**
 * typed に表現しない shape tree ノードの raw escape hatch。未対応の図形・
 * グループ・コネクタ等を round-trip のために保存する。
 */
export interface SourceRawShapeNode {
  readonly kind: "raw";
  readonly nodeId?: SourceNodeId;
  readonly raw: RawSidecar;
  readonly handle?: SourceHandle;
}

/** shape tree の source node 共用体。 */
export type SourceShapeNode =
  | SourceShape
  | SourceConnector
  | SourceGroup
  | SourceImage
  | SourceTable
  | SourceChart
  | SourceSmartArt
  | SourceRawShapeNode;
