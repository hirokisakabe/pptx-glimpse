/**
 * Presentation hierarchy の source 型。slide / layout / master / theme の参照
 * 関係を part path で表す。各 part は分離したまま保持し、cascade (master →
 * layout → slide) の解決は computed view の責務とする
 * (`docs/cleandoc-source-computed-view.md`)。
 */

import type { PartPath, SourceHandle } from "./handles.js";
import type { RawSidecar } from "./raw.js";
import type { SourceShapeNode } from "./shapes.js";
import type { Emu } from "./units.js";

/** スライドサイズ (`p:sldSz`)。EMU で保持する。 */
export interface SlideSize {
  readonly width: Emu;
  readonly height: Emu;
}

/** theme part (`ppt/theme/themeN.xml`)。 */
export interface SourceTheme {
  readonly partPath: PartPath;
  readonly name?: string;
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
  readonly shapes: readonly SourceShapeNode[];
  readonly handle?: SourceHandle;
  readonly rawSidecars?: readonly RawSidecar[];
}

/** slide part。適用される layout を参照する。 */
export interface SourceSlide {
  readonly partPath: PartPath;
  /** この slide が参照する layout part。 */
  readonly layoutPartPath: PartPath;
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
