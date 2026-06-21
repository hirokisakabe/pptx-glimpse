/**
 * CleanDoc source model のトップレベル型。
 *
 * `@pptx-glimpse/document` が所有する canonical な document 表現で、writer /
 * editor / round-trip の source of truth となる。renderer-specific fallback や
 * pixel output 固有の値は含まない (`docs/document-boundaries.md` /
 * `docs/cleandoc-source-computed-view.md`)。computed view はこの source から
 * 派生して生成する derived projection であり、本型には含めない。
 */

import type { Diagnostic } from "./diagnostics.js";
import type { PackageGraph } from "./package-graph.js";
import type {
  SourcePresentation,
  SourceSlide,
  SourceSlideLayout,
  SourceSlideMaster,
  SourceTheme,
} from "./presentation.js";

export interface CleanDocSource {
  /** package part / relationship / content type / media の構造。 */
  readonly packageGraph: PackageGraph;
  readonly presentation: SourcePresentation;
  readonly slides: readonly SourceSlide[];
  readonly slideLayouts: readonly SourceSlideLayout[];
  readonly slideMasters: readonly SourceSlideMaster[];
  readonly themes: readonly SourceTheme[];
  /** document 正しさに関する診断。 */
  readonly diagnostics: readonly Diagnostic[];
}
