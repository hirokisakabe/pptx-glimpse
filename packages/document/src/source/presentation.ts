/**
 * Source type of Presentation hierarchy. Reference slide / layout / master / theme
 * Represent the relationship as a part path. Keep each part separate and cascade (master ->
 * **resolution** of layout -> slide) is the responsibility of the computed view. However, it is necessary to solve
 * The material (background / clrMap / clrMapOvr / theme scheme / showMasterSp) is source
 * are kept source-local so they can be retained without returning to raw XML.
 */

import type { PartPath, SourceHandle } from "./handles.js";
import type { RawSidecar } from "./raw.js";
import type {
  SourceColor,
  SourceEffectList,
  SourceFill,
  SourceOutline,
  SourceShapeNode,
  SourceTextStyle,
} from "./shapes.js";
import type { Emu } from "./units.js";

/** Slide size (`p:sldSz`).kept in EMUs. */
export interface SlideSize {
  readonly width: Emu;
  readonly height: Emu;
}

/**
 * color map (`p:clrMap`) / color map override (`p:clrMapOvr`).logical name
 * (`bg1` / `tx1` / `accent1` etc.) to theme scheme slot (`dk1` / `lt1` etc.)
 * Keep the mapping unresolved.
 */
export interface SourceColorMap {
  readonly mapping: Readonly<Record<string, string>>;
}

/**
 * slide / layout / master background (`p:bg`). directly fill (`p:bgPr`) or
 * Either the theme's style matrix reference (`p:bgRef`) or raw.
 */
export type SourceBackground =
  | { readonly kind: "fill"; readonly fill: SourceFill }
  | { readonly kind: "styleReference"; readonly index: number; readonly color: SourceColor }
  | { readonly kind: "raw"; readonly raw: RawSidecar };

/** The theme's color scheme (`a:clrScheme`). slot name -> unconverted map of concrete color. */
export interface SourceThemeColorScheme {
  /** `dk1` / `lt1` / `dk2` / `lt2` / `accent1`..`accent6` / `hlink` / `folHlink`. */
  readonly colors: Readonly<Record<string, SourceColor>>;
}

/** The minimal subset of theme's font schemes (`a:fontScheme`). */
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

/** A subset of theme's format scheme (`a:fmtScheme`) to use for shape style ref resolution. */
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
  /** Parent slide master part for this layout. */
  readonly masterPartPath: PartPath;
  /** layout type (`p:sldLayout@type`). */
  readonly type?: string;
  readonly background?: SourceBackground;
  /** Layout `p:clrMapOvr` (when overridden). */
  readonly colorMapOverride?: SourceColorMap;
  /** `p:sldLayout@showMasterSp` (master shape visibility). */
  readonly showMasterShapes?: boolean;
  readonly shapes: readonly SourceShapeNode[];
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

/** presentation part (`ppt/presentation.xml`). Maintain the order of slides. */
export interface SourcePresentation {
  readonly partPath: PartPath;
  readonly slideSize?: SlideSize;
  readonly defaultTextStyle?: SourceTextStyle;
  /** `p:sldIdLst` slide part paths reflecting the order of. */
  readonly slidePartPaths: readonly PartPath[];
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}
