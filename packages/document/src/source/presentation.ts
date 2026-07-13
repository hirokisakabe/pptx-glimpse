/**
 * Source types for the presentation hierarchy.
 *
 * Slide, layout, master, and theme relationships are represented by part paths. Each
 * part stays separate; resolving the master -> layout -> slide cascade is the computed
 * view's responsibility. Materials needed for that resolution, such as background,
 * clrMap, clrMapOvr, theme scheme, and showMasterSp, are kept as source-local values so
 * the source model can retain them without falling back to raw XML.
 */

import type { PartPath, SourceHandle } from "./handles.js";
import type { RawSidecar } from "./raw.js";
import type {
  SourceColor,
  SourceEffectList,
  SourceFill,
  SourceOutline,
  SourceShapeNode,
  SourceTextBodyProperties,
  SourceTextStyle,
} from "./shapes.js";
import type { Emu } from "./units.js";

/** Slide size (`p:sldSz`), kept in EMUs. */
export interface SlideSize {
  readonly width: Emu;
  readonly height: Emu;
}

/**
 * Color map (`p:clrMap`) / color map override (`p:clrMapOvr`). Keeps the mapping from
 * logical names (`bg1` / `tx1` / `accent1`, etc.) to theme scheme slots (`dk1` / `lt1`,
 * etc.) unresolved.
 */
export interface SourceColorMap {
  readonly mapping: Readonly<Record<string, string>>;
}

/**
 * Slide / layout / master background (`p:bg`). This is either a direct fill
 * (`p:bgPr`), a theme style matrix reference (`p:bgRef`), or raw content.
 */
export type SourceBackground =
  | { readonly kind: "fill"; readonly fill: SourceFill }
  | { readonly kind: "styleReference"; readonly index: number; readonly color: SourceColor }
  | { readonly kind: "raw"; readonly raw: RawSidecar };

/** Theme color scheme (`a:clrScheme`). Maps slot names to unconverted concrete colors. */
export interface SourceThemeColorScheme {
  /** `dk1` / `lt1` / `dk2` / `lt2` / `accent1`..`accent6` / `hlink` / `folHlink`. */
  readonly colors: Readonly<Record<string, SourceColor>>;
}

/** Minimal subset of the theme font scheme (`a:fontScheme`). */
export interface SourceThemeFontScheme {
  readonly majorLatin?: string;
  readonly minorLatin?: string;
  readonly majorEastAsian?: string;
  readonly minorEastAsian?: string;
  readonly majorComplexScript?: string;
  readonly minorComplexScript?: string;
  readonly majorJapanese?: string;
  readonly minorJapanese?: string;
}

/** Subset of the theme format scheme (`a:fmtScheme`) used for shape style ref resolution. */
export interface SourceThemeFormatScheme {
  readonly fillStyles: readonly SourceFill[];
  readonly lineStyles: readonly SourceOutline[];
  readonly effectStyles: readonly (SourceEffectList | undefined)[];
  readonly backgroundFillStyles: readonly SourceFill[];
}

/** theme part (`ppt/theme/themeN.xml`). */
export interface SourceTheme {
  readonly partPath: PartPath;
  readonly name?: string;
  readonly colorScheme?: SourceThemeColorScheme;
  readonly fontScheme?: SourceThemeFontScheme;
  readonly formatScheme?: SourceThemeFormatScheme;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/** Slide master part. References the theme and child layouts. */
export interface SourceSlideMaster {
  readonly partPath: PartPath;
  /** Author-visible name from `p:cSld@name`. */
  readonly name?: string;
  /** Theme part referenced by this master. */
  readonly themePartPath?: PartPath;
  /** Layout parts belonging to this master. */
  readonly layoutPartPaths: readonly PartPath[];
  readonly background?: SourceBackground;
  /** Master `p:clrMap`. */
  readonly colorMap?: SourceColorMap;
  readonly txStyles?: SourceMasterTextStyles;
  readonly shapes: readonly SourceShapeNode[];
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

export interface SourceMasterTextStyles {
  readonly titleStyle?: SourceTextStyle;
  readonly bodyStyle?: SourceTextStyle;
  readonly otherStyle?: SourceTextStyle;
}

/** Slide layout part. References the parent master. */
export interface SourceSlideLayout {
  readonly partPath: PartPath;
  /** Author-visible name from `p:cSld@name`. */
  readonly name?: string;
  /** Parent slide master part for this layout. */
  readonly masterPartPath: PartPath;
  /** layout type (`p:sldLayout@type`). */
  readonly type?: string;
  readonly background?: SourceBackground;
  /** Layout `p:clrMapOvr` (when overridden). */
  readonly colorMapOverride?: SourceColorMap;
  /** `p:sldLayout@show` (layout visibility). When omitted, the effective value is visible (`show ?? true`). */
  readonly show?: boolean;
  /** `p:sldLayout@showMasterSp` (master shape visibility). */
  readonly showMasterShapes?: boolean;
  readonly shapes: readonly SourceShapeNode[];
  /**
   * From-scratch authoring default applied to text-bearing shapes added to slides that
   * reference this layout. PowerPoint has no package-level layout margin property, so
   * the resolved values are materialized into each authored shape's `a:bodyPr`.
   */
  readonly defaultTextBodyProperties?: SourceTextBodyProperties;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/** Slide part. References the applied layout. */
export interface SourceSlide {
  readonly partPath: PartPath;
  /** Layout part referenced by this slide. */
  readonly layoutPartPath: PartPath;
  readonly background?: SourceBackground;
  /** Slide `p:clrMapOvr` (when overridden). */
  readonly colorMapOverride?: SourceColorMap;
  /** `p:sld@showMasterSp` (master shape visibility). */
  readonly showMasterShapes?: boolean;
  readonly shapes: readonly SourceShapeNode[];
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/** Presentation part (`ppt/presentation.xml`). Maintains slide order. */
export interface SourcePresentation {
  readonly partPath: PartPath;
  readonly slideSize?: SlideSize;
  readonly defaultTextStyle?: SourceTextStyle;
  /** Slide part paths reflecting `p:sldIdLst` order. */
  readonly slidePartPaths: readonly PartPath[];
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}
