# Legacy parser semantics audit after the PptxSourceModel default switch

- Status: implementation note for [#485](https://github.com/hirokisakabe/pptx-glimpse/issues/485)
- Date: 2026-06-27

This note records the first shrink pass after public SVG/PNG conversion moved to
the PptxSourceModel document path. It complements
[core-document-dogfood-migration.md](./core-document-dogfood-migration.md) and
[document-boundaries.md](./document-boundaries.md). For #485, this note
supersedes the earlier "current state" and parallel-reader descriptions in the
dogfood migration note where those descriptions still mention public conversion
flowing through the old parser.

It also updates the inventory requested by
[#527](https://github.com/hirokisakabe/pptx-glimpse/issues/527). Older issue text
in [#516](https://github.com/hirokisakabe/pptx-glimpse/issues/516) used the
transitionary `CleanDoc` name and `packages/pptx-glimpse/src/*` paths. In the
current repository those map to PptxSourceModel and the workspace package paths
under `packages/core`, `packages/document`, and `packages/renderer`; for example
the old `cleandoc-renderer-adapter.ts` reference is now
`packages/core/src/pptx-computed-view-renderer-adapter.ts`.

This is the current-state companion to the historical migration plan. The
default switch tracked by [#481](https://github.com/hirokisakabe/pptx-glimpse/issues/481)
is complete, and the package-boundary cleanup follow-ups
[#510](https://github.com/hirokisakabe/pptx-glimpse/issues/510) and
[#511](https://github.com/hirokisakabe/pptx-glimpse/issues/511) are closed.

## Current owner split

The public conversion path is now:

```text
convertPptxToSvg / convertPptxToPng
  -> convertPptxToSvgViaDocumentPath / convertPptxToPngViaDocumentPath
  -> @pptx-glimpse/document readPptx
  -> createComputedView
  -> core-owned PptxSourceModel renderer adapter
  -> renderer
```

The old parser path is no longer part of public conversion orchestration. Its
remaining render entry points live in
`packages/core/src/parser-path-oracle.ts` and are intentionally scoped as
an internal oracle for parity checks.

## Semantics moved to `document`

These reusable OOXML semantics are now owned by `@pptx-glimpse/document` source
or computed view code:

| Semantics                                                 | Document owner                                                      | Previous parser overlap removed or reduced                                          |
| --------------------------------------------------------- | ------------------------------------------------------------------- | ----------------------------------------------------------------------------------- |
| Package graph, slide order, relationships, media payloads | `reader/read-pptx.ts`, source package graph, computed relationships | Public conversion no longer calls `parsePptxData` for package traversal             |
| Slide/layout/master chain resolution                      | `createComputedView`                                                | Public conversion no longer calls `parseSlideWithLayout` for cascade assembly       |
| Background fallback                                       | `createComputedView`                                                | Public conversion uses computed `background` instead of parser-side fallback        |
| Theme color and color map resolution                      | source theme data + computed color resolution                       | Public conversion uses computed colors through the adapter                          |
| Placeholder filtering and effective element ordering      | `createComputedView`                                                | Core converter no longer owns parser-path `mergeElements`                           |
| Text style cascade for rendering/inspection               | `createComputedView`                                                | `collectUsedFonts` now reads PptxSourceModel/computed text instead of parser slides |
| Theme font source data for rendering/inspection           | source theme data + core runtime helpers                            | `collectUsedFonts` and rendering setup no longer depend on parser slide assembly    |
| Table/chart/image relationship resolution for rendering   | `createComputedView` + adapter                                      | Parser path remains only as comparison oracle                                       |

## Responsibilities left outside `document`

These remain in core or renderer by design:

| Responsibility                                                               | Owner                                                          | Reason                                                                                                    |
| ---------------------------------------------------------------------------- | -------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------- |
| Public API options                                                           | `packages/core/src/converter.ts`, `packages/core/src/index.ts` | Stable API compatibility remains a core package concern                                                   |
| Warning setup, font setup, SVG text-output mode, PNG sizing                  | `packages/core/src/experimental-document-renderer.ts`          | Renderer environment setup and runtime conversion options are core concerns                               |
| PptxSourceModel computed view to renderer model mapping                      | `packages/core/src/pptx-computed-view-renderer-adapter.ts`     | The renderer model is a display-oriented compatibility target, not the PptxSourceModel source model       |
| Chart XML to renderer chart model mapping                                    | Adapter calling parser `parseChart`                            | Chart rendering model is still renderer-specific; move only after a chart source/computed contract exists |
| SmartArt drawing XML fallback to renderer shape tree                         | Adapter calling parser `parseShapeTree`                        | This is a rendering fallback for resolved diagram drawing XML, not canonical document semantics yet       |
| Font discovery, font mapping, text measurement, text-to-path, SVG/PNG output | `@pptx-glimpse/renderer`                                       | Renderer-specific behavior per `document-boundaries.md`                                                   |

## Remaining legacy parser call-site inventory

The old parser code still has three explicit roles. This inventory lists the
current cross-boundary call sites using current file paths. Parser subsystem
internals under `packages/core/src/parser/` and their colocated unit tests remain
because they implement and protect these roles; they are not additional public
conversion orchestration.

| Current call site                                                          | Legacy parser dependency                                                                                                                             | Classification               | Current role                                                                                                                                       | Follow-up state                                                                    |
| -------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| `packages/core/src/parser-path-oracle.ts`                                  | Imports `parsePptxData` / `parseSlideWithLayout` from `packages/core/src/pptx-data-parser.ts` and XML cache helpers from `packages/core/src/parser/` | `parser oracle`              | Renders through the old parser only for document-path parity checks and VRT/test oracles.                                                          | Keep until the explicit parser oracle is retired.                                  |
| `packages/core/src/parser-path-oracle.ts`                                  | Exports `buildEffectiveSlideElements` around old parser slide/layout/master outputs                                                                  | `parser oracle`              | Preserves old parser effective element ordering and placeholder filtering for oracle comparisons.                                                  | Keep with the parser oracle; remove with `pptx-data-parser.ts` retirement.         |
| `packages/core/src/dual-reader-structural-comparison.test.ts`              | Imports `parsePptxData`, `parseSlideWithLayout`, and `buildEffectiveSlideElements`                                                                   | `structural comparison`      | Compares a focused supported subset of the old parser render model against the PptxSourceModel document path and core adapter output.              | Retire after VRT and public regressions fully cover the parser oracle's value.     |
| `packages/core/src/pptx-computed-view-renderer-adapter.ts` `adaptChart`    | Imports and calls `parseChart` from `packages/core/src/parser/chart-parser.ts`                                                                       | `renderer-specific fallback` | Maps resolved chart XML from PptxSourceModel into the current renderer chart model until a document-owned chart source/computed contract exists.   | Replace after chart source/computed contracts are designed.                        |
| `packages/core/src/pptx-computed-view-renderer-adapter.ts` `adaptSmartArt` | Imports `navigateOrdered` / `parseShapeTree` from `packages/core/src/parser/slide-parser.ts` and XML helpers from `packages/core/src/parser/`        | `renderer-specific fallback` | Turns resolved SmartArt diagram drawing XML into the current renderer group/shape model as a visual fallback, not as canonical document semantics. | Replace after a document-owned diagram drawing source model exists.                |
| `packages/core/src/parse-render.integration.test.ts`                       | Imports `parseShapeTree` and parser XML/archive relationship types                                                                                   | `renderer-specific fallback` | Keeps direct coverage for old parser shape-tree output that the renderer and SmartArt fallback still consume.                                      | Remove or move when adapter fallback use of `parseShapeTree` is replaced.          |
| `packages/core/src/parser-path-oracle.test.ts`                             | Imports `buildEffectiveSlideElements` and `ParsedSlide` from `packages/core/src/pptx-data-parser.ts`                                                 | `parser oracle`              | Protects oracle-only effective element merging semantics after that logic moved out of `converter.ts`.                                             | Remove with `parser-path-oracle.ts`.                                               |
| None found outside the roles above                                         | n/a                                                                                                                                                  | `obsolete`                   | No obsolete old-parser cross-boundary call site was found in the current tree during the #527 audit.                                               | If a future obsolete use appears, remove it or split a follow-up issue before use. |

No public API currently imports old parser render orchestration.
`packages/core/src/converter.ts` calls only
`convertPptxToSvgViaDocumentPath` and `convertPptxToPngViaDocumentPath` from
`packages/core/src/experimental-document-renderer.ts`. That document path reads
with `@pptx-glimpse/document` `readPptx`, builds `createComputedView`, adapts via
`packages/core/src/pptx-computed-view-renderer-adapter.ts`, and then invokes the
renderer.

## Shrink applied in #485

- Moved `collectUsedFonts` from `parsePptxData` / `parseSlideWithLayout` to
  `readPptx` / `createComputedView`.
- Moved parser-path SVG/PNG rendering and effective element merging out of
  `converter.ts` into `parser-path-oracle.ts`.
- Updated VRT zero-diff and document-path regression tests to import the parser
  oracle explicitly.

## Remaining duplication and follow-up candidates

Some duplication intentionally remains:

- `pptx-data-parser.ts` still assembles slide/layout/master render models. Keep
  it until the parser oracle is no longer needed for zero-diff gates.
- Parser unit tests under `packages/core/src/parser/` still protect the
  old oracle and adapter fallback helpers. Delete or narrow them only after the
  corresponding oracle role is removed.
- `pptx-computed-view-renderer-adapter.ts` still calls parser helpers for chart parsing
  and SmartArt fallback rendering. Split follow-up issues should define
  document-owned chart/diagram source contracts before moving this logic.
- Comments in `packages/document/src/computed/create-computed-view.ts`
  that mention current-parser compatibility are retained as parity signposts.
  They should be removed only when the document path intentionally owns a
  different behavior and tests are updated with that decision.

Suggested follow-up slices:

1. Replace adapter SmartArt fallback use of `parseShapeTree` with a
   document-owned diagram drawing source model.
2. Introduce a chart source/computed contract and remove adapter dependence on
   parser `parseChart`.
3. Retire `dual-reader-structural-comparison.test.ts` after VRT and public
   regression tests fully cover the parser oracle's remaining value.
4. Remove `pptx-data-parser.ts` and parser render-model tests after the explicit
   parser-path oracle is no longer used by VRT/CI.
