/**
 * Top-level types for the PptxSourceModel source model.
 *
 * This is the canonical PPTX document representation owned by
 * `@pptx-glimpse/document`. Rather than exposing package parts directly as the public
 * API, it groups presentation, slides, layouts, masters, themes, relationships, media,
 * and content types as OOXML source semantics. Upper layers such as core, editor-core,
 * and pom may consume this package, but this package must not depend on them. Renderer
 * output is produced by the core adapter, and PptxSourceModel does not know about it.
 *
 * This model is the source of truth for writer, editor, and round-trip workflows. It
 * keeps source-local values, relationship ids, part paths, element ordering, typed
 * PPTX-domain units, stable source handles, diagnostics, and raw preservation hooks.
 * Unsupported OOXML, vendor extensions, mc:AlternateContent, and unsupported DrawingML
 * are not mixed into the typed operation API. They are preserved as raw sidecars or raw
 * package parts for structural round-tripping.
 *
 * PptxSourceModel must not include renderer-specific fallbacks, environment-specific
 * font substitution, SVG/PNG output, pixel-output values, or pom authoring primitives.
 * Slide/layout/master/theme cascades, relationship resolution, theme color resolution,
 * placeholder and text style resolution, and similar effective values are derived from
 * the source as a non-mutating computed view.
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
  /** Structure of package part / relationship / content type / media. */
  readonly packageGraph: PackageGraph;
  readonly presentation: SourcePresentation;
  readonly slides: readonly SourceSlide[];
  readonly slideLayouts: readonly SourceSlideLayout[];
  readonly slideMasters: readonly SourceSlideMaster[];
  readonly themes: readonly SourceTheme[];
  /** Diagnostics about document correctness. */
  readonly diagnostics: readonly Diagnostic[];
  /** typed PptxSourceModel operation and dirty scope. The writer determines the minimum update range from this. */
  readonly edits?: readonly PptxSourceModelEdit[];
}

export type PptxSourceModelEdit = PptxSourceModelTextRunEdit;

export interface PptxSourceModelTextRunEdit {
  readonly kind: "replaceTextRunPlainText";
  readonly handle: SourceHandle;
  readonly text: string;
}
