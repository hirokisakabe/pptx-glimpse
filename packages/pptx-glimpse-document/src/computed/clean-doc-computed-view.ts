/**
 * CleanDoc computed view の型。
 *
 * source model を mutation せず、slide / layout / master / theme cascade から
 * 消費しやすい effective value を派生させる。renderer 固有の pixel conversion
 * や font fallback はここに含めない。
 */

import type {
  Emu,
  MediaPart,
  PartPath,
  Relationship,
  RelationshipId,
  SourceBackground,
  SourceCellBorders,
  SourceFill,
  SourceImage,
  SourceOutline,
  SourceParagraphProperties,
  SourceRawShapeNode,
  SourceRunProperties,
  SourceShape,
  SourceShapeNode,
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

export interface CleanDocComputedView {
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
  | ComputedImageElement
  | ComputedTableElement
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
  readonly textBody?: ComputedTextBody;
  readonly placeholderMatch?: ComputedPlaceholderMatch;
}

export interface ComputedImageElement extends ComputedElementBase {
  readonly kind: "image";
  readonly sourceNode: SourceImage;
  readonly transform?: SourceTransform;
  readonly relationship?: ComputedRelationship;
  readonly media?: MediaPart;
}

export interface ComputedTableElement extends ComputedElementBase {
  readonly kind: "table";
  readonly sourceNode: SourceTable;
  readonly transform?: SourceTransform;
  readonly table: ComputedTableData;
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
  | { readonly kind: "raw"; readonly source: SourceFill };

export interface ComputedOutline {
  readonly width?: Emu;
  readonly fill?: ComputedFill;
  readonly source: SourceOutline;
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
  readonly properties?: SourceParagraphProperties;
  readonly runs: readonly ComputedTextRun[];
}

export interface ComputedTextRun {
  readonly text: string;
  readonly properties?: ComputedRunProperties;
}

export type ComputedRunProperties = Omit<SourceRunProperties, "color"> & {
  readonly color?: ComputedColor;
};
