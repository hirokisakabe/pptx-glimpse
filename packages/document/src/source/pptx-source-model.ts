/**
 * PptxSourceModel source model  top-level types.
 *
 * Internal note.
 * instead of exposing package parts directly as the public API、presentation / slides /
 * Internal note.
 * Internal note.
 * Internal note.
 * adapter output and、PptxSourceModel does not know it directly。
 *
 * Internal note.
 * local values、Relationship id、part path、element ordering、typed PPTX-domain
 * units、stable source handles、diagnostics、raw preservation hooks .
 * Internal note.
 * Internal note.
 * kept for structural preservation。
 *
 * PptxSourceModel  renderer-specific fallback、environment-specific font
 * Internal note.
 * must not be included。slide/layout/master/theme cascade、relationship resolution、theme
 * color resolution、placeholder / text style resolution and similar effective values are
 * Internal note.
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
  /** Internal note. */
  readonly packageGraph: PackageGraph;
  readonly presentation: SourcePresentation;
  readonly slides: readonly SourceSlide[];
  readonly slideLayouts: readonly SourceSlideLayout[];
  readonly slideMasters: readonly SourceSlideMaster[];
  readonly themes: readonly SourceTheme[];
  /** Diagnostics about document correctness. */
  readonly diagnostics: readonly Diagnostic[];
  /** Internal note. */
  readonly edits?: readonly PptxSourceModelEdit[];
}

export type PptxSourceModelEdit = PptxSourceModelTextRunEdit;

export interface PptxSourceModelTextRunEdit {
  readonly kind: "replaceTextRunPlainText";
  readonly handle: SourceHandle;
  readonly text: string;
}
