/**
 * Presentation hierarchy の source 型。slide / layout / master / theme の参照
 * 関係を part path で表す。各 part は分離したまま保持し、cascade (master →
 * layout → slide) の **解決** は computed view の責務とする
 * (`docs/cleandoc-source-computed-view.md`)。ただし解決に必要な素材
 * (background / clrMap / clrMapOvr / theme scheme / showMasterSp) は、source が
 * raw XML に戻らず保持できるよう source-local に置く
 * (`docs/cleandoc-minimal-poc-scope.md` の Included PPTX Subset)。
 */

import type { PartPath, SourceHandle } from "./handles.js";
import type { RawSidecar } from "./raw.js";
import type {
  SourceColor,
  SourceEffectList,
  SourceFill,
  SourceOutline,
  SourceShapeNode,
} from "./shapes.js";
import type { Emu } from "./units.js";

/** スライドサイズ (`p:sldSz`)。EMU で保持する。 */
export interface SlideSize {
  readonly width: Emu;
  readonly height: Emu;
}

/**
 * color map (`p:clrMap`) / color map override (`p:clrMapOvr`)。logical name
 * (`bg1` / `tx1` / `accent1` 等) から theme scheme slot (`dk1` / `lt1` 等) への
 * マッピングを未解決のまま保持する。
 */
export interface SourceColorMap {
  readonly mapping: Readonly<Record<string, string>>;
}

/**
 * slide / layout / master の background (`p:bg`)。直接 fill (`p:bgPr`) か、
 * theme の style matrix 参照 (`p:bgRef`)、あるいは raw のいずれか。
 */
export type SourceBackground =
  | { readonly kind: "fill"; readonly fill: SourceFill }
  | { readonly kind: "styleReference"; readonly index: number; readonly color: SourceColor }
  | { readonly kind: "raw"; readonly raw: RawSidecar };

/** theme の color scheme (`a:clrScheme`)。slot 名 → 具体色の未変換マップ。 */
export interface SourceThemeColorScheme {
  /** `dk1` / `lt1` / `dk2` / `lt2` / `accent1`..`accent6` / `hlink` / `folHlink`。 */
  readonly colors: Readonly<Record<string, SourceColor>>;
}

/** theme の font scheme (`a:fontScheme`) の最小 subset。 */
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

/** theme の format scheme (`a:fmtScheme`) のうち shape style ref 解決に使う subset。 */
export interface SourceThemeFormatScheme {
  readonly fillStyles: readonly SourceFill[];
  readonly lineStyles: readonly SourceOutline[];
  readonly effectStyles: readonly (SourceEffectList | undefined)[];
  readonly backgroundFillStyles: readonly SourceFill[];
}

/** theme part (`ppt/theme/themeN.xml`)。 */
export interface SourceTheme {
  readonly partPath: PartPath;
  readonly name?: string;
  readonly colorScheme?: SourceThemeColorScheme;
  readonly fontScheme?: SourceThemeFontScheme;
  readonly formatScheme?: SourceThemeFormatScheme;
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/** slide master part。theme と配下 layout を参照する。 */
export interface SourceSlideMaster {
  readonly partPath: PartPath;
  /** この master が参照する theme part。 */
  readonly themePartPath?: PartPath;
  /** この master に属する layout part 群。 */
  readonly layoutPartPaths: readonly PartPath[];
  readonly background?: SourceBackground;
  /** master の `p:clrMap`。 */
  readonly colorMap?: SourceColorMap;
  readonly shapes: readonly SourceShapeNode[];
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/** slide layout part。親 master を参照する。 */
export interface SourceSlideLayout {
  readonly partPath: PartPath;
  /** この layout の親 slide master part。 */
  readonly masterPartPath: PartPath;
  /** layout type (`p:sldLayout@type`)。 */
  readonly type?: string;
  readonly background?: SourceBackground;
  /** layout の `p:clrMapOvr` (上書きする場合)。 */
  readonly colorMapOverride?: SourceColorMap;
  /** `p:sldLayout@showMasterSp` (master 図形の表示可否)。 */
  readonly showMasterShapes?: boolean;
  readonly shapes: readonly SourceShapeNode[];
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/** slide part。適用される layout を参照する。 */
export interface SourceSlide {
  readonly partPath: PartPath;
  /** この slide が参照する layout part。 */
  readonly layoutPartPath: PartPath;
  readonly background?: SourceBackground;
  /** slide の `p:clrMapOvr` (上書きする場合)。 */
  readonly colorMapOverride?: SourceColorMap;
  /** `p:sld@showMasterSp` (master 図形の表示可否)。 */
  readonly showMasterShapes?: boolean;
  readonly shapes: readonly SourceShapeNode[];
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/** presentation part (`ppt/presentation.xml`)。slide の並び順を保持する。 */
export interface SourcePresentation {
  readonly partPath: PartPath;
  readonly slideSize?: SlideSize;
  /** `p:sldIdLst` の順序を反映した slide part path 列。 */
  readonly slidePartPaths: readonly PartPath[];
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}
