/**
 * PptxSourceModel computed view の型。
 *
 * source model を mutation せず、slide / layout / master / theme cascade から
 * 消費しやすい effective value を派生させる。renderer 固有の pixel conversion
 * や font fallback はここに含めない。
 */

import type {
  Emu,
  MediaPart,
  PartPath,
  RawPackagePart,
  Relationship,
  RelationshipId,
  SourceBackground,
  SourceBlipEffects,
  SourceCellBorders,
  SourceChart,
  SourceConnector,
  SourceEffectList,
  SourceFill,
  SourceGroup,
  SourceImage,
  SourceOutline,
  SourceParagraphProperties,
  SourceRawShapeNode,
  SourceRunProperties,
  SourceShape,
  SourceShapeNode,
  SourceSmartArt,
  SourceTable,
  SourceTableCell,
  SourceTableColumn,
  SourceTableRow,
  SourceTextBodyProperties,
  SourceTransform,
} from "../source/index.js";

export interface CreateComputedViewOptions {
  /** 1-based slide numbers. 未指定時は presentation order の全 slide。 */
  readonly slides?: readonly number[];
  /** `p:sld@showMasterSp` / `p:sldLayout@showMasterSp` を反映する。既定 true。 */
  readonly applyMasterVisibility?: boolean;
}

export interface PptxComputedView {
  readonly slideSize?: ComputedSlideSize;
  readonly slides: readonly ComputedSlide[];
}

export interface ComputedSlideSize {
  readonly width: Emu;
  readonly height: Emu;
}

export interface ComputedSlide {
  readonly slideNumber: number;
  readonly partPath: PartPath;
  readonly layoutPartPath?: PartPath;
  readonly masterPartPath?: PartPath;
  readonly themePartPath?: PartPath;
  readonly slideSize?: ComputedSlideSize;
  readonly relationships: readonly ComputedRelationship[];
  readonly colorMap: Readonly<Record<string, string>>;
  readonly colorScheme: Readonly<Record<string, string>>;
  readonly background?: ComputedBackground;
  readonly showMasterShapes: boolean;
  readonly layoutShowMasterShapes: boolean;
  readonly elements: readonly ComputedElement[];
}

export interface ComputedRelationship {
  readonly id: RelationshipId;
  readonly type: string;
  readonly source: Relationship;
  readonly target: string;
  readonly targetMode?: "Internal" | "External";
  readonly targetPartPath?: PartPath;
  readonly media?: MediaPart;
}

export type ComputedElement =
  | ComputedShapeElement
  | ComputedConnectorElement
  | ComputedGroupElement
  | ComputedImageElement
  | ComputedTableElement
  | ComputedChartElement
  | ComputedSmartArtElement
  | ComputedRawElement;

export type ComputedElementLayer = "master" | "layout" | "slide";

interface ComputedElementBase {
  readonly sourceLayer: ComputedElementLayer;
  readonly sourcePartPath: PartPath;
  readonly sourceNode: SourceShapeNode;
}

export interface ComputedShapeElement extends ComputedElementBase {
  readonly kind: "shape";
  readonly sourceNode: SourceShape;
  readonly transform?: SourceTransform;
  readonly geometry?: SourceShape["geometry"];
  readonly fill?: ComputedFill;
  readonly outline?: ComputedOutline;
  readonly effects?: ComputedEffectList;
  readonly textBody?: ComputedTextBody;
  readonly placeholderMatch?: ComputedPlaceholderMatch;
}

export interface ComputedImageElement extends ComputedElementBase {
  readonly kind: "image";
  readonly sourceNode: SourceImage;
  readonly transform?: SourceTransform;
  readonly relationship?: ComputedRelationship;
  readonly media?: MediaPart;
  readonly effects?: ComputedEffectList;
  readonly blipEffects?: ComputedBlipEffects;
}

export interface ComputedConnectorElement extends ComputedElementBase {
  readonly kind: "connector";
  readonly sourceNode: SourceConnector;
  readonly transform?: SourceTransform;
  readonly geometry?: SourceConnector["geometry"];
  readonly outline?: ComputedOutline;
  readonly effects?: ComputedEffectList;
}

export interface ComputedGroupElement extends ComputedElementBase {
  readonly kind: "group";
  readonly sourceNode: SourceGroup;
  readonly transform?: SourceTransform;
  readonly childTransform?: SourceTransform;
  readonly effects?: ComputedEffectList;
  readonly children: readonly ComputedElement[];
}

export interface ComputedTableElement extends ComputedElementBase {
  readonly kind: "table";
  readonly sourceNode: SourceTable;
  readonly transform?: SourceTransform;
  readonly table: ComputedTableData;
}

export interface ComputedChartElement extends ComputedElementBase {
  readonly kind: "chart";
  readonly sourceNode: SourceChart;
  readonly transform?: SourceTransform;
  readonly relationship?: ComputedRelationship;
  readonly chartXml?: string;
}

export interface ComputedSmartArtElement extends ComputedElementBase {
  readonly kind: "smartArt";
  readonly sourceNode: SourceSmartArt;
  readonly transform?: SourceTransform;
  readonly dataRelationship?: ComputedRelationship;
  readonly drawingRelationship?: ComputedRelationship;
  readonly drawingPartPath?: PartPath;
  readonly drawingXml?: string;
  readonly drawingRelationships: readonly ComputedRelationship[];
  readonly media: readonly MediaPart[];
  readonly diagramDrawing?: ComputedDiagramDrawing;
}

export interface ComputedDiagramDrawing {
  readonly sourcePartPath: PartPath;
  readonly rawXml: string;
  readonly rawPart?: RawPackagePart;
  readonly rawHandle: { readonly partPath: PartPath };
  readonly relationships: readonly ComputedRelationship[];
  readonly media: readonly MediaPart[];
  readonly childTransform?: SourceTransform;
  readonly children: readonly ComputedElement[];
  readonly diagnostics: readonly ComputedDiagramDrawingDiagnostic[];
}

export interface ComputedDiagramDrawingDiagnostic {
  readonly severity: "warning";
  readonly code: "diagram-drawing-shape-tree-missing";
  readonly message: string;
  readonly sourcePartPath: PartPath;
}

export interface ComputedTableData {
  readonly columns: readonly SourceTableColumn[];
  readonly rows: readonly ComputedTableRow[];
}

export interface ComputedTableRow {
  readonly source: SourceTableRow;
  readonly height: SourceTableRow["height"];
  readonly cells: readonly ComputedTableCell[];
}

export interface ComputedTableCell {
  readonly source: SourceTableCell;
  readonly textBody?: ComputedTextBody;
  readonly fill?: ComputedFill;
  readonly borders?: ComputedCellBorders;
  readonly gridSpan: number;
  readonly rowSpan: number;
  readonly hMerge: boolean;
  readonly vMerge: boolean;
}

export type ComputedCellBorders = {
  readonly [K in keyof SourceCellBorders]?: ComputedOutline;
};

export interface ComputedRawElement extends ComputedElementBase {
  readonly kind: "raw";
  readonly sourceNode: SourceRawShapeNode;
}

export interface ComputedPlaceholderMatch {
  readonly layout?: SourceShape;
  readonly master?: SourceShape;
}

export type ComputedBackground =
  | {
      readonly kind: "fill";
      readonly source: SourceBackground;
      readonly fill: ComputedFill;
      readonly sourceLayer: ComputedElementLayer;
    }
  | {
      readonly kind: "styleReference";
      readonly source: SourceBackground;
      readonly index: number;
      readonly color?: ComputedColor;
      readonly sourceLayer: ComputedElementLayer;
    }
  | {
      readonly kind: "raw";
      readonly source: SourceBackground;
      readonly sourceLayer: ComputedElementLayer;
    };

export type ComputedFill =
  | { readonly kind: "none"; readonly source: SourceFill }
  | { readonly kind: "solid"; readonly source: SourceFill; readonly color: ComputedColor }
  | {
      readonly kind: "gradient";
      readonly source: SourceFill;
      readonly stops: readonly ComputedGradientStop[];
      readonly gradientType: "linear" | "radial";
      readonly angle?: number;
      readonly centerX?: number;
      readonly centerY?: number;
    }
  | {
      readonly kind: "pattern";
      readonly source: SourceFill;
      readonly preset: string;
      readonly foregroundColor: ComputedColor;
      readonly backgroundColor: ComputedColor;
    }
  | {
      readonly kind: "image";
      readonly source: SourceFill;
      readonly relationship?: ComputedRelationship;
      readonly media?: MediaPart;
      readonly tile?: Extract<SourceFill, { readonly kind: "image" }>["tile"];
    }
  | { readonly kind: "raw"; readonly source: SourceFill };

export interface ComputedGradientStop {
  readonly position: number;
  readonly color: ComputedColor;
}

export interface ComputedOutline {
  readonly width?: Emu;
  readonly fill?: ComputedFill;
  readonly source: SourceOutline;
}

export interface ComputedEffectList {
  readonly source: SourceEffectList;
  readonly outerShadow?: ComputedOuterShadow;
  readonly innerShadow?: ComputedInnerShadow;
  readonly glow?: ComputedGlow;
  readonly softEdge?: SourceEffectList["softEdge"];
}

export interface ComputedOuterShadow {
  readonly blurRadius: Emu;
  readonly distance: Emu;
  readonly direction: number;
  readonly color: ComputedColor;
  readonly alignment: string;
  readonly rotateWithShape: boolean;
}

export interface ComputedInnerShadow {
  readonly blurRadius: Emu;
  readonly distance: Emu;
  readonly direction: number;
  readonly color: ComputedColor;
}

export interface ComputedGlow {
  readonly radius: Emu;
  readonly color: ComputedColor;
}

export interface ComputedBlipEffects {
  readonly source: SourceBlipEffects;
  readonly grayscale: boolean;
  readonly biLevel?: SourceBlipEffects["biLevel"];
  readonly blur?: SourceBlipEffects["blur"];
  readonly lum?: SourceBlipEffects["lum"];
  readonly duotone?: ComputedDuotoneEffect;
  readonly clrChange?: ComputedColorChangeEffect;
}

export interface ComputedDuotoneEffect {
  readonly color1: ComputedColor;
  readonly color2: ComputedColor;
}

export interface ComputedColorChangeEffect {
  readonly from: ComputedColor;
  readonly to: ComputedColor;
}

export interface ComputedColor {
  /** `#rrggbb`。 */
  readonly hex: string;
  /** 0-1 の正規化 opacity。 */
  readonly alpha: number;
}

export interface ComputedTextBody {
  readonly properties?: SourceTextBodyProperties;
  readonly paragraphs: readonly ComputedParagraph[];
}

export interface ComputedParagraph {
  readonly properties?: ComputedParagraphProperties;
  readonly runs: readonly ComputedTextRun[];
}

export interface ComputedParagraphProperties extends Omit<
  SourceParagraphProperties,
  "bulletColor" | "defaultRunProperties"
> {
  readonly bulletColor?: ComputedColor;
}

export interface ComputedTextRun {
  readonly text: string;
  readonly properties?: ComputedRunProperties;
}

export type ComputedRunProperties = Omit<SourceRunProperties, "color"> & {
  readonly color?: ComputedColor;
};
