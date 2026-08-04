/**
 * PptxSourceModel computed view types.
 *
 * This view derives effective document values from the slide / layout / master / theme
 * cascade without mutating the source model. The source keeps each package part's
 * authored state and raw preservation hooks, while this view deterministically resolves
 * presentation order, slide size, the per-slide layout/master/theme chain, relationship
 * targets, background fallback, theme color maps and schemes, placeholder matches, text
 * style cascades, showMasterSp visibility, and computed element ordering.
 *
 * The computed view is a projection of document semantics, not an editable source of
 * truth. It keeps the part path, source node, and source layer needed for source
 * provenance and diagnostics, but it does not materialize renderer-contract-friendly
 * required defaults or `null` fallbacks here. Chart XML type, series, categories,
 * legend, and similar data are projected into parsed chart data as OOXML document
 * semantics, while renderer-specific drawing completion and pixel layout decisions
 * remain in the adapter / renderer.
 *
 * Renderer-specific pixel conversion, system font discovery, font fallback, text
 * measurement / wrapping, text-to-path, SVG/PNG output decisions, and renderer warning
 * policy remain the responsibility of the core adapter or renderer. Raw elements, raw
 * fills, and raw backgrounds are kept for preservation and diagnostics, and the adapter
 * decides direct rendering policy.
 *
 * Repository-level package and dependency boundaries are documented in the
 * [architecture overview](../../../../docs/architecture/overview.md).
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
  SourceRectangleAlignment,
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
  /** 1-based slide numbers. When omitted, all slides in presentation order are used. */
  readonly slides?: readonly number[];
  /** Applies `p:sld@showMasterSp` / `p:sldLayout@showMasterSp`. Defaults to true. */
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
  readonly placeholder?: ComputedPlaceholder;
}

export interface ComputedPlaceholder {
  readonly type: string;
  readonly index: number;
  readonly orientation: string;
  readonly size: string;
  readonly hasCustomPrompt: boolean;
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
  readonly placeholderMatch?: ComputedPlaceholderMatch;
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
  /** Non-destructive projection of the group's parent-local authored transform. */
  readonly transform?: SourceTransform;
  /**
   * Non-destructive projection of the authored child coordinate space. Effective
   * affine composition belongs to the core adapter / renderer boundary.
   */
  readonly childTransform?: SourceTransform;
  readonly fill?: ComputedFill;
  readonly effects?: ComputedEffectList;
  readonly children: readonly ComputedElement[];
}

export interface ComputedTableElement extends ComputedElementBase {
  readonly kind: "table";
  readonly sourceNode: SourceTable;
  readonly transform?: SourceTransform;
  readonly table: ComputedTableData;
  readonly placeholderMatch?: ComputedPlaceholderMatch;
}

export interface ComputedChartElement extends ComputedElementBase {
  readonly kind: "chart";
  readonly sourceNode: SourceChart;
  readonly transform?: SourceTransform;
  readonly relationship?: ComputedRelationship;
  readonly chartXml?: string;
  readonly chartData?: ComputedChartData;
  readonly placeholderMatch?: ComputedPlaceholderMatch;
}

export type ComputedChartType =
  | "combo"
  | "bar"
  | "line"
  | "pie"
  | "doughnut"
  | "scatter"
  | "bubble"
  | "area"
  | "radar"
  | "stock"
  | "surface"
  | "ofPie";

export interface ComputedChartData {
  readonly chartType: ComputedChartType;
  readonly title: string | null;
  readonly series: readonly ComputedChartSeries[];
  readonly categories: readonly string[];
  /** Ordered OOXML plot groups for a supported category combo chart. */
  readonly plotGroups?: readonly ComputedChartPlotGroup[];
  /** Value axes referenced by combo plot groups, keyed through `axId`. */
  readonly valueAxes?: readonly ComputedChartValueAxis[];
  readonly barDirection?: "col" | "bar";
  readonly holeSize?: number;
  readonly radarStyle?: "standard" | "marker" | "filled";
  readonly ofPieType?: "pie" | "bar";
  readonly secondPieSize?: number;
  readonly splitPos?: number;
  readonly legend: ComputedChartLegend | null;
}

export interface ComputedChartPlotGroup {
  readonly chartType: "bar" | "line";
  /** Authored `c:grouping@val`, when present. */
  readonly grouping?: string;
  /** Positions in `ComputedChartData.series`, retained in OOXML plot-group order. */
  readonly seriesIndexes: readonly number[];
  /** Ordered `c:axId` references authored on the plot group. */
  readonly axisIds: readonly string[];
  /** Group references that resolve to exactly one supported `c:catAx` definition. */
  readonly categoryAxisIds?: readonly string[];
  /** Group references that resolve to exactly one `c:valAx` definition. */
  readonly valueAxisIds?: readonly string[];
  /** The referenced `c:valAx/c:axId`, when it can be resolved. */
  readonly valueAxisId?: string;
}

export interface ComputedChartValueAxis {
  /** Unsigned `c:valAx/c:axId@val`. */
  readonly id: string;
  readonly position?: "l" | "r" | "t" | "b";
}

export interface ComputedChartSeries {
  readonly name: string | null;
  readonly values: readonly number[];
  readonly xValues?: readonly number[];
  readonly bubbleSizes?: readonly number[];
  readonly color: ComputedColor;
  /** Existing OOXML identity, available for editable bar/line category series. */
  readonly source?: ComputedCategoryChartSeriesSource;
}

export interface ComputedCategoryChartSeriesSource {
  readonly chartType: "bar" | "line";
  /** Unsigned `c:ser/c:idx@val`, scoped by `chartType`. */
  readonly index: number;
}

export interface ComputedChartLegend {
  readonly position: "b" | "t" | "l" | "r" | "tr";
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
  readonly placeholderMatch?: ComputedPlaceholderMatch;
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
  readonly placeholderMatch?: ComputedPlaceholderMatch;
}

export interface ComputedPlaceholderMatch {
  /** Compatible `p:sp` projection retained for existing shape consumers. */
  readonly layout?: SourceShape;
  /** Compatible `p:sp` projection retained for existing shape consumers. */
  readonly master?: SourceShape;
  /** Exact placeholder-capable layout source node, including picture/graphic-frame nodes. */
  readonly layoutNode: SourceShapeNode;
  /** Exact placeholder-capable master source node when the type category has one unique match. */
  readonly masterNode?: SourceShapeNode;
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

interface ComputedGradientStop {
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

interface ComputedOuterShadow {
  readonly blurRadius: Emu;
  readonly distance: Emu;
  readonly direction: number;
  readonly color: ComputedColor;
  readonly alignment: SourceRectangleAlignment;
  readonly rotateWithShape: boolean;
}

interface ComputedInnerShadow {
  readonly blurRadius: Emu;
  readonly distance: Emu;
  readonly direction: number;
  readonly color: ComputedColor;
}

interface ComputedGlow {
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
  /** `#rrggbb`. */
  readonly hex: string;
  /** Normalized opacity in the 0-1 range. */
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

interface ComputedParagraphProperties extends Omit<
  SourceParagraphProperties,
  "bulletColor" | "defaultRunProperties"
> {
  readonly bulletColor?: ComputedColor;
}

export interface ComputedTextRun {
  readonly text: string;
  readonly properties?: ComputedRunProperties;
}

export type ComputedRunProperties = Omit<
  SourceRunProperties,
  "color" | "underlineColor" | "highlight"
> & {
  readonly color?: ComputedColor;
  readonly underlineColor?: ComputedColor;
  readonly highlight?: ComputedColor;
};
