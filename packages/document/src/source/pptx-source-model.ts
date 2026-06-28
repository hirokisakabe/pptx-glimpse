/**
 * PptxSourceModel source model のトップレベル型。
 *
 * `@pptx-glimpse/document` が所有する canonical な PPTX document 表現で、OOXML
 * package parts をそのまま公開 API にする代わりに、presentation / slides /
 * layouts / masters / themes / relationships / media / content types を source
 * semantics として束ねる。core、editor-core、pom などの上位層はこの package を
 * consume できるが、document からそれらへ依存してはいけない。renderer は core
 * adapter の出力を consume し、PptxSourceModel を直接知らない。
 *
 * この model は writer / editor / round-trip の source of truth であり、source
 * local な値、relationship id、part path、element ordering、typed PPTX-domain
 * units、stable source handles、diagnostics、raw preservation hooks を保持する。
 * Unsupported OOXML、vendor extension、mc:AlternateContent、未対応 DrawingML は
 * typed operation の primary API に混ぜず、raw sidecar / raw package part として
 * structural preservation のために残す。
 *
 * PptxSourceModel には renderer-specific fallback、environment-specific font
 * substitution、SVG/PNG output、pixel output 固有の値、pom authoring primitive を
 * 入れない。slide/layout/master/theme cascade、relationship resolution、theme
 * color resolution、placeholder / text style resolution などの effective value は
 * source を mutate しない computed view として派生させる。
 */

import type { Diagnostic } from "./diagnostics.js";
import type { SourceHandle } from "./handles.js";
import type { PackageGraph } from "./package-graph.js";
import type {
  SourcePresentation,
  SourceSlide,
  SourceSlideLayout,
  SourceSlideMaster,
  SourceTheme,
} from "./presentation.js";

export interface PptxSourceModel {
  /** package part / relationship / content type / media の構造。 */
  readonly packageGraph: PackageGraph;
  readonly presentation: SourcePresentation;
  readonly slides: readonly SourceSlide[];
  readonly slideLayouts: readonly SourceSlideLayout[];
  readonly slideMasters: readonly SourceSlideMaster[];
  readonly themes: readonly SourceTheme[];
  /** document 正しさに関する診断。 */
  readonly diagnostics: readonly Diagnostic[];
  /** typed PptxSourceModel operation と dirty scope。writer はここから最小更新範囲を判断する。 */
  readonly edits?: readonly PptxSourceModelEdit[];
}

export type PptxSourceModelEdit = PptxSourceModelTextRunEdit;

export interface PptxSourceModelTextRunEdit {
  readonly kind: "replaceTextRunPlainText";
  readonly handle: SourceHandle;
  readonly text: string;
}
