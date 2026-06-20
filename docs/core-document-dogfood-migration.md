# Core `@pptx-glimpse/document` dogfood migration

- Status: RFC decision for [#452](https://github.com/hirokisakabe/pptx-glimpse/issues/452)
- Date: 2026-06-20

This note records how the public `convertPptxToSvg` / `convertPptxToPng` path
should eventually dogfood `@pptx-glimpse/document`. It feeds into
[#445](https://github.com/hirokisakabe/pptx-glimpse/issues/445) and builds on
the package boundary decision in
[document-boundaries.md](./document-boundaries.md), the CleanDoc source/computed
view decision in
[cleandoc-source-computed-view.md](./cleandoc-source-computed-view.md), and the
raw OOXML round-trip policy in
[raw-ooxml-round-trip.md](./raw-ooxml-round-trip.md).

The target keeps the standalone `pptx-glimpse` conversion API intact while
moving reusable OOXML reading semantics into `@pptx-glimpse/document`.

## Current Core Flow

The current public conversion API is orchestrated by
`packages/pptx-glimpse/src/converter.ts`:

```text
PPTX Buffer | Uint8Array
  -> parsePptxData(input)
  -> slide path selection
  -> parseSlideWithLayout(slideNumber, path, data)
  -> merge master/layout/slide render elements
  -> renderSlideToSvg(slide, slideSize)
  -> svgToPng(svg, options) for PNG output
```

`parsePptxData` and `parseSlideWithLayout` are currently the effective core
reader. They unzip the package, parse XML, collect relationships/theme data, and
produce the renderer model consumed by `pptx-glimpse-renderer`.

The current path already contains computed-view behavior, but it is implicit and
spread across parser and converter code:

- `parseSlideWithLayout` resolves slide, layout, and master package chains.
- Theme colors are resolved through theme color schemes and effective color
  maps.
- Slide background fallback is applied in slide -> layout -> master order.
- Placeholder geometry and text styles are merged from layout/master sources.
- `converter.ts` merges master, layout, and slide elements for rendering.
- Template placeholders are filtered before SVG output.
- `convertPptxToPng` reuses `convertPptxToSvg` and then calls the PNG renderer,
  forcing path text output for resvg compatibility.

This means the current parser is already a parser plus computed render-view
builder. It is useful for rendering, but it is not the right canonical document
source model for writer/editor/round-trip work.

## Target Flow Through `document`

The final conversion flow should be:

```text
PPTX Buffer | Uint8Array
  -> @pptx-glimpse/document readPptx(input)
  -> CleanDoc source model
  -> createComputedView(source, options)
  -> createRenderView(computed, renderOptions)
  -> existing pptx-glimpse-renderer model
  -> renderSlideToSvg(slide, slideSize)
  -> svgToPng(svg, options) for PNG output
```

The ownership boundary should be:

- `@pptx-glimpse/document` owns OOXML package reading, CleanDoc source types,
  relationship/package bookkeeping, source handles, raw preservation sidecars,
  and document-level computed view generation.
- `@pptx-glimpse/core` owns the public conversion API, slide selection,
  compatibility options, warning behavior, font setup, and the
  CleanDoc-computed-view-to-renderer-model adapter.
- `pptx-glimpse-renderer` keeps owning SVG/PNG rendering, font measurement,
  text-to-path behavior, visual fallbacks, and renderer-specific warnings.

`@pptx-glimpse/document` must not import `core`, `editor-core`, the renderer, or
pom. Core depends on `document`; the renderer consumes the adapter output and
does not need to know CleanDoc directly.

## Initial Parallel Reader Period

Adopt a parallel period instead of replacing the existing parser in one step.

The first `document` reader should be introduced behind internal or experimental
paths and compared against the current parser. The public default
`convertPptxToSvg` / `convertPptxToPng` path should continue to use the current
parser until the CleanDoc reader, computed view, and adapter reach rendering
parity for the supported fixture set.

Recommended sequence during the parallel period:

1. Keep the current parser as the production render path.
2. Add `@pptx-glimpse/document` source types and `readPptx(input)` for a small
   but real subset of slides.
3. Add computed view generation for slide size, slide order, relationships,
   theme resolution, background fallback, placeholder cascade, and text style
   inheritance.
4. Add a CleanDoc computed view to current renderer model adapter in core.
5. Run dual-reader comparison tests where both paths parse the same fixtures and
   the adapter output is compared against the current renderer model.
6. Switch selected fixtures to render through the document path in tests.
7. Only after parity is stable, make the document path the default public
   conversion path.

Parallel reading is temporary dogfood scaffolding, not a permanent dual-source
architecture. Once the document path becomes the default, parser semantics that
are not renderer-specific should move into `@pptx-glimpse/document`; renderer
fallbacks and visual diagnostics should stay outside it.

## Adapter Policy

The adapter should translate from CleanDoc computed view to the existing
renderer model, not merge the renderer model into CleanDoc.

Adapter responsibilities:

- Convert computed slide order, slide size, backgrounds, and effective element
  trees into `pptx-glimpse-renderer` model objects.
- Convert computed theme colors, fills, outlines, effects, table/chart/image
  references, and text properties into the render-ready shapes expected by the
  renderer.
- Preserve source provenance in diagnostics where practical, so renderer or core
  warnings can point back to source handles without making raw OOXML a renderer
  concern.
- Apply render-specific pixel conversion at the renderer boundary using the
  existing 96 DPI policy.
- Keep font fallback, text measurement, text wrapping, and SVG/PNG output
  choices in renderer/core, not in `document`.

The adapter should be located with core orchestration, because it is a
compatibility layer between the lower-level document package and the current
renderer package. It can be split into a package later if multiple consumers need
the same render view contract.

The adapter should start as a direct mapping into the current renderer model.
After the document path is proven, renderer model fields that merely duplicate
computed-view fields can be simplified in separate, rendering-focused PRs.

## Dogfood Verification

Dogfood should be verified at multiple levels instead of relying on only VRT:

| Level                      | Purpose                                                                     | Suggested checks                                                                                           |
| -------------------------- | --------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------- |
| Document reader unit tests | Ensure OOXML parts become CleanDoc source nodes with stable references      | minimal PPTX fixtures, relationship graph assertions, raw sidecar preservation assertions                  |
| Computed view unit tests   | Ensure document semantics are resolved outside the renderer                 | theme/color map resolution, background fallback, placeholder matching, text style cascade, unit conversion |
| Adapter tests              | Ensure CleanDoc computed view maps to the current renderer contract         | compare adapter output with selected current parser model snapshots or focused structural assertions       |
| Core dual-reader tests     | Ensure public conversion can dogfood `document` without changing behavior   | run both paths on shared fixtures and compare selected render model fields                                 |
| Snapshot VRT               | Catch visual regressions once fixtures are opted into the document path     | existing `vrt/snapshot` cases and shared real PPTX fixtures                                                |
| LibreOffice VRT            | Compare generated output against external rendering references where useful | existing `vrt/libreoffice` cases after the document path affects rendering                                 |
| Package verification       | Ensure the published API shape remains compatible                           | existing build, typecheck, package verification, and API import smoke tests                                |

For early PRs, unit and adapter tests are more important than snapshot churn.
VRT snapshots should only be updated when the public render output intentionally
changes. The first dogfood milestone should aim for no visual output changes.

## PR Slicing

Split the migration into small PRs with explicit compatibility checkpoints:

1. Add `@pptx-glimpse/document` package skeleton, CleanDoc source types, and
   package boundary tests.
2. Implement minimal `readPptx(input)` that can read presentation metadata,
   slide list, relationships, and slide size while preserving raw package parts.
3. Add CleanDoc source coverage for shapes, text, images, theme references,
   layouts, and masters needed by the existing fixtures.
4. Add `createComputedView(source, options)` for slide size, slide order,
   relationship resolution, theme/color map resolution, background fallback,
   placeholder cascade, and text style inheritance.
5. Add the core-owned adapter from computed view to the current renderer model.
6. Add dual-reader comparison tests for shared fixtures without changing the
   public default path.
7. Add an internal or experimental option that lets tests render selected
   fixtures through the document path.
8. Move selected VRT/shared fixtures to the document path and fix parity gaps.
9. Make the document path the default for `convertPptxToSvg` and
   `convertPptxToPng` once parity is stable.
10. Remove or shrink obsolete parser code after the document path owns the
    reusable OOXML semantics.

Each PR should keep the published conversion API compatible unless it explicitly
declares a breaking change. The migration should not require VRT snapshot
updates until a PR intentionally changes visible rendering behavior.

## Non-goals

This plan does not implement the migration itself.

It also does not require:

- Replacing the renderer model immediately.
- Moving font discovery, font fallback, text measurement, or SVG/PNG output into
  `@pptx-glimpse/document`.
- Changing `convertPptxToSvg` / `convertPptxToPng` signatures.
- Updating VRT snapshots as part of this design decision.
- Making `document` depend on core, renderer, editor-core, or pom.

## Conclusion for #445

The conclusion to reflect in #445 is:

`pptx-glimpse` core should dogfood `@pptx-glimpse/document` by routing the public
SVG/PNG conversion path through CleanDoc source reading, computed document view
generation, and a core-owned adapter into the existing renderer model.

The migration should use a temporary parallel reader period. The current parser
remains the production render path while `document` reader coverage, computed
view behavior, and adapter parity are proven by unit tests, dual-reader tests,
package verification, and VRT. After parity is stable, the document path becomes
the default and reusable OOXML semantics move out of the current parser into
`@pptx-glimpse/document`.

The existing renderer model remains a render-specific adapter target during the
migration. CleanDoc is the canonical source model, computed view resolves
document semantics, and renderer/core continue to own rendering-specific
fallbacks and output behavior.
