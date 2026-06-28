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

The SmartArt fallback contract requested by
[#528](https://github.com/hirokisakabe/pptx-glimpse/issues/528) is recorded in
[smartart-fallback-contract.md](./smartart-fallback-contract.md). That note
records the current computed diagram drawing contract, the renderer
`GroupElement` output contract, and the removal status for the old
`parseShapeTree` SmartArt fallback dependency.

This is the current-state companion to the historical migration plan. The
default switch tracked by [#481](https://github.com/hirokisakabe/pptx-glimpse/issues/481)
is complete, and the package-boundary cleanup follow-ups
[#510](https://github.com/hirokisakabe/pptx-glimpse/issues/510) and
[#511](https://github.com/hirokisakabe/pptx-glimpse/issues/511) are closed.

The retirement conditions requested by
[#531](https://github.com/hirokisakabe/pptx-glimpse/issues/531) are recorded
below. This issue defines when the parser oracle and dual-reader structural
comparison can be removed; it does not remove those files.

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

| Responsibility                                                               | Owner                                                          | Reason                                                                                                                                                                    |
| ---------------------------------------------------------------------------- | -------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Public API options                                                           | `packages/core/src/converter.ts`, `packages/core/src/index.ts` | Stable API compatibility remains a core package concern                                                                                                                   |
| Warning setup, font setup, SVG text-output mode, PNG sizing                  | `packages/core/src/experimental-document-renderer.ts`          | Renderer environment setup and runtime conversion options are core concerns                                                                                               |
| PptxSourceModel computed view to renderer model mapping                      | `packages/core/src/pptx-computed-view-renderer-adapter.ts`     | The renderer model is a display-oriented compatibility target, not the PptxSourceModel source model                                                                       |
| Chart XML to renderer chart model mapping                                    | `packages/core/src/renderer-chart-data-converter.ts`           | Chart rendering model is still renderer-specific; move only after a chart source/computed contract exists                                                                 |
| SmartArt computed diagram drawing fallback                                   | Document computed view + core adapter                          | The diagram drawing package contract is document-owned, while renderer `GroupElement` mapping remains a core adapter concern; see [smartart-fallback-contract.md](./smartart-fallback-contract.md) |
| Font discovery, font mapping, text measurement, text-to-path, SVG/PNG output | `@pptx-glimpse/renderer`                                       | Renderer-specific behavior per `document-boundaries.md`                                                                                                                   |

## Remaining legacy parser call-site inventory

The old parser code still has three explicit roles. This inventory lists the
current cross-boundary call sites using current file paths. Parser subsystem
internals under `packages/core/src/parser/` and their colocated unit tests remain
because they implement and protect these roles; they are not additional public
conversion orchestration.

| Current call site                                                          | Legacy parser dependency                                                                                                                                                                                                                    | Classification               | Current role                                                                                                                                       | Follow-up state                                                                    |
| -------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| `packages/core/src/parser-path-oracle.ts`                                  | Imports `parsePptxData` / `parseSlideWithLayout` from `packages/core/src/pptx-data-parser.ts` and XML cache helpers from `packages/core/src/parser/`                                                                                        | `parser oracle`              | Renders through the old parser only for document-path parity checks and VRT/test oracles.                                                          | Keep until the explicit parser oracle is retired.                                  |
| `packages/core/src/parser-path-oracle.ts`                                  | Exports `buildEffectiveSlideElements` around old parser slide/layout/master outputs                                                                                                                                                         | `parser oracle`              | Preserves old parser effective element ordering and placeholder filtering for oracle comparisons.                                                  | Keep with the parser oracle; remove with `pptx-data-parser.ts` retirement.         |
| `vrt/snapshot/document-path-regression.test.ts`                            | Imports and calls `convertPptxToPngViaParserPath` from `packages/core/src/parser-path-oracle.ts`                                                                                                                                            | `parser oracle`              | Compares document-path PNG output against the explicit parser oracle for the document-path VRT set.                                                | Remove or retarget when the parser oracle no longer backs document-path VRT.       |
| `vrt/snapshot/document-path-zero-diff-gate.test.ts`                        | Imports and calls `convertPptxToPngViaParserPath` from `packages/core/src/parser-path-oracle.ts`                                                                                                                                            | `parser oracle`              | Keeps the default-switch zero-diff gate against the explicit parser oracle.                                                                        | Remove or retarget when the parser oracle no longer backs the zero-diff gate.      |
| `packages/core/src/dual-reader-structural-comparison.test.ts`              | Imports `parsePptxData`, `parseSlideWithLayout`, and `buildEffectiveSlideElements`                                                                                                                                                          | `structural comparison`      | Compares a focused supported subset of the old parser render model against the PptxSourceModel document path and core adapter output.              | Retire after VRT and public regressions fully cover the parser oracle's value.     |
| `bench/conversion.bench.ts`                                                | Imports and calls `parsePptxData` / `parseSlideWithLayout` from `packages/core/src/pptx-data-parser.ts`                                                                                                                                     | `parser oracle`              | Benchmarks the old parser pipeline as an explicit baseline beside the public document-path conversion APIs.                                        | Remove or retarget with `pptx-data-parser.ts` retirement.                          |
| `packages/core/src/pptx-computed-view-renderer-adapter.ts` `adaptChart`    | Calls `convertChartXmlToRendererChartData` from `packages/core/src/renderer-chart-data-converter.ts`                                                                                                                                        | `renderer-specific fallback` | Maps resolved chart XML from PptxSourceModel into the current renderer chart model until a document-owned chart source/computed contract exists.   | Replace after chart source/computed contracts are designed.                        |
| `packages/core/src/pptx-computed-view-renderer-adapter.ts` `adaptSmartArt` | No old parser import; consumes `ComputedSmartArtElement.diagramDrawing` from `@pptx-glimpse/document`                                                                                                                                       | `renderer-specific fallback` | Turns document computed diagram drawing children into the current renderer `GroupElement` model.                                                   | Keep as core adapter mapping until renderer/model ownership changes.               |
| `packages/core/src/parse-render.integration.test.ts`                       | Imports `parseShapeTree` and parser XML/archive relationship types                                                                                                                                                                          | `parser oracle`              | Keeps direct coverage for old parser shape-tree output while parser oracle consumers remain.                                                       | Remove or move with the parser oracle, not with SmartArt fallback.                 |
| `packages/core/src/parser-path-oracle.test.ts`                             | Imports `buildEffectiveSlideElements` and `ParsedSlide` from `packages/core/src/pptx-data-parser.ts`                                                                                                                                        | `parser oracle`              | Protects oracle-only effective element merging semantics after that logic moved out of `converter.ts`.                                             | Remove with `parser-path-oracle.ts`.                                               |
| `packages/core/src/text-style-resolver.ts`                                 | Imports `resolveThemeFont` from `packages/core/src/parser/text-style-parser.ts`                                                                                                                                                             | `parser oracle`              | Supports old parser text-style inheritance used by `packages/core/src/pptx-data-parser.ts`; public font collection no longer depends on it.        | Remove or narrow with `pptx-data-parser.ts` retirement.                            |
| `packages/core/src/text-style-resolver.test.ts`                            | Imports `applyTextStyleInheritance` and `TextStyleContext` from `packages/core/src/text-style-resolver.ts`                                                                                                                                  | `parser oracle`              | Protects the old parser text-style inheritance helper while `pptx-data-parser.ts` still consumes it.                                               | Remove or narrow with `text-style-resolver.ts` retirement.                         |
| None found outside the roles above                                         | n/a                                                                                                                                                                                                                                         | `obsolete`                   | No obsolete old-parser cross-boundary call site was found in the current tree during the #527 audit.                                               | If a future obsolete use appears, remove it or split a follow-up issue before use. |

No public API currently imports old parser render orchestration.
`packages/core/src/converter.ts` calls only
`convertPptxToSvgViaDocumentPath` and `convertPptxToPngViaDocumentPath` from
`packages/core/src/experimental-document-renderer.ts`. That document path reads
with `@pptx-glimpse/document` `readPptx`, builds `createComputedView`, adapts via
`packages/core/src/pptx-computed-view-renderer-adapter.ts`, and then invokes the
renderer.

## Parser oracle retirement plan

The explicit parser oracle currently exists for these consumers and should not
be removed until each role has a non-parser replacement or is intentionally
deleted:

| Consumer                                                      | Oracle use                                                                                                             | Reason to keep now                                                                                                     | Replacement or deletion signal                                                                                                    |
| ------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------- |
| `vrt/snapshot/document-path-regression.test.ts`               | Calls `convertPptxToPngViaParserPath` and compares document-path PNG output against it in memory.                      | Provides broad visual parity coverage across all shared and generated VRT fixtures without committed parser snapshots. | Replace with committed document-path snapshots, external references, or another stable non-parser baseline for the same case set. |
| `vrt/snapshot/document-path-zero-diff-gate.test.ts`           | Calls `convertPptxToPngViaParserPath` with zero mismatch tolerance.                                                    | Keeps the post-default-switch document path pixel-identical to the old parser path for the VRT gate.                   | Delete after the default-switch gate is no longer needed, or retarget it to a non-parser golden baseline.                         |
| `packages/core/src/dual-reader-structural-comparison.test.ts` | Builds old-parser slides with `parsePptxData`, `parseSlideWithLayout`, and `buildEffectiveSlideElements`.              | Gives focused structural assertions for supported fields before PNG rendering hides the source of a regression.        | Delete after equivalent non-parser unit/regression tests cover the same fields and fixture intent.                                |
| `packages/core/src/parser-path-oracle.test.ts`                | Unit-tests `buildEffectiveSlideElements`.                                                                              | Protects oracle-only master/layout/slide ordering and placeholder filtering while VRT still depends on the oracle.     | Delete with `parser-path-oracle.ts` once no remaining test/VRT imports that helper.                                               |
| `bench/conversion.bench.ts`                                   | Benchmarks `parsePptxData` / `parseSlideWithLayout` as the old-parser baseline.                                        | Keeps performance comparisons with the historical parser path explicit.                                                | Remove the parser benchmark lane or retarget it to a document-path-only baseline when parser performance is no longer tracked.    |
| `packages/core/src/text-style-resolver.ts` and its unit test  | Supports text inheritance consumed by `pptx-data-parser.ts`.                                                           | Remains needed only because the parser oracle still assembles old-parser render models.                                | Remove or narrow after `pptx-data-parser.ts` no longer exists or no longer consumes this helper.                                  |
| Parser subsystem unit tests under `packages/core/src/parser/` | Protect parser helpers still reached by `pptx-data-parser.ts` and parser-oracle tests.                                   | Parser helpers still serve the explicit oracle, but not the public SmartArt fallback.                                  | Delete or narrow with the parser oracle once no remaining oracle consumer imports them.                                            |

`parser-path-oracle.ts` itself can be removed only after no VRT, unit test,
benchmark, or adapter fallback imports `convertPptxToPngViaParserPath`,
`convertPptxToSvgViaParserPath`, or `buildEffectiveSlideElements`. Until then,
it remains the intentional quarantine boundary that prevents old parser
rendering from leaking back into public conversion orchestration.

## Dual-reader structural comparison value

`packages/core/src/dual-reader-structural-comparison.test.ts` is narrower than
VRT by design. It compares renderer-model structure from the old parser against
the PptxSourceModel document path and adapter for selected real fixtures. Its
current verification value is:

| Value category             | What it catches before VRT                                                                                             |
| -------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| Slide-level contract       | Slide size, slide numbers, and background fill differences before they become opaque PNG diffs.                        |
| Effective element ordering | Master/layout/slide ordering, `showMasterSp` visibility, template placeholder filtering, and empty slide placeholders. |
| Basic shape semantics      | Transform, preset geometry, placeholder metadata, solid fills, outlines, and theme-resolved colors.                    |
| Supported text subset      | Plain text runs, body margins/anchor, and basic run styling for the subset currently exposed by the document path.     |
| Raster image semantics     | Image transform, MIME type, payload presence, and crop rectangle mapping.                                              |
| Adapter diagnostics        | Ensures raw skipped elements are expected warnings instead of silent structural drift.                                 |

The test intentionally does not validate alt text/source names, adjustment
handles, complex effects/fills, raw graphic frames, groups, bullets, numbering,
tabs, hyperlinks, text fields not yet exposed by the document path, or
renderer-only fallback fields. Those categories must be covered by dedicated
document/computed/adapter tests before this structural comparison can be
retired as a safety net for them.

## Replacement coverage map

The parser oracle's value can be replaced only by a combination of VRT,
document-path regression tests, and focused unit tests. No single layer replaces
it fully today.

| Verification layer                                           | Already replaces or partially replaces                                                                                                         | Still not enough for retirement by itself                                                                                                        |
| ------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| Snapshot VRT and document-path VRT                           | Broad visual coverage across shared and generated fixtures, including final SVG/PNG behavior and renderer integration.                         | PNG diffs do not identify whether a regression came from source reading, computed cascade, adapter mapping, or renderer drawing.                 |
| `document-path-zero-diff-gate.test.ts`                       | Enforces exact parser-vs-document visual parity for the default-switch fixture set.                                                            | It still depends on the parser oracle, so it is a gate for keeping the oracle honest rather than a replacement for it.                           |
| Public conversion regression tests in core                   | Ensure `convertPptxToSvg` / `convertPptxToPng` keep using the document path and preserve public API behavior.                                  | They do not compare detailed old-parser render-model semantics or cover all visual fixture cases.                                                |
| `@pptx-glimpse/document` reader and computed-view unit tests | Cover package graph, relationships, theme/color resolution, background fallback, placeholder cascade, source provenance, and raw preservation. | They do not prove the core adapter still maps computed values into the renderer model expected by existing SVG/PNG rendering.                    |
| Adapter unit tests                                           | Cover focused computed-view-to-renderer mappings and renderer-specific defaults/fallbacks without invoking the old parser.                     | Existing adapter tests do not yet cover the full dual-reader comparison scope for both shared fixtures and all current edge cases.               |
| Renderer unit tests                                          | Protect SVG/PNG rendering behavior once a renderer model is already constructed.                                                               | They cannot validate that PptxSourceModel reading, computed cascade, or adapter mapping produced the same model the old parser previously did.   |
| LibreOffice VRT                                              | Provides an external rendering reference for generated cases where LibreOffice is an acceptable comparator.                                    | LibreOffice is not PowerPoint and does not define old-parser compatibility; tolerances make it unsuitable as the exact parser-oracle substitute. |

The remaining un-replaced gap is structural parity for the supported shared
fixture subset before render-time rasterization. Close that gap by expanding
document/computed/adapter tests until they cover the value categories listed in
the previous section without importing `pptx-data-parser.ts` or
`parser-path-oracle.ts`.

## Retirement conditions

`packages/core/src/dual-reader-structural-comparison.test.ts` can be deleted
when all of these are true:

1. Adapter and document/computed-view tests cover every category in the
   dual-reader comparison value table for `real-product-page.pptx` and
   `real-basic-theme.pptx`, or the fixture-specific expectation has been
   intentionally moved to VRT/public regression coverage with a documented
   reason.
2. Inherited placeholder/theme text styling that is currently guarded by
   `includeRunProperties: false` for `real-basic-theme.pptx` is either exposed
   and covered by non-parser tests, or explicitly documented as out of scope for
   parser parity.
3. Raw skipped element diagnostics covered by the dual-reader test are asserted
   in adapter or document-path regression tests without constructing old-parser
   slides.
4. Document-path VRT continues to cover the full shared fixture set and generated
   VRT set, or an intentional replacement baseline has been committed.
5. A final `rg` check shows the only remaining imports of
   `parsePptxData`, `parseSlideWithLayout`, and `buildEffectiveSlideElements`
   are unrelated to this structural comparison.

`packages/core/src/pptx-data-parser.ts` and old parser render-model tests can be
deleted when all of these are true:

1. `parser-path-oracle.ts` has no remaining consumers, including VRT,
   `parser-path-oracle.test.ts`, `dual-reader-structural-comparison.test.ts`,
   and `bench/conversion.bench.ts`.
2. Document-path VRT no longer uses the parser path as its in-memory PNG
   baseline. It either compares against committed document-path snapshots,
   external references, or another non-parser baseline with equivalent fixture
   coverage and documented tolerance policy.
3. Old-parser-only helpers such as effective slide merging, placeholder
   filtering, text-style inheritance, and theme font resolution have either
   moved into `@pptx-glimpse/document` / adapter tests or are explicitly removed
   as obsolete behavior.
4. Parser subsystem tests have been split by remaining consumer: oracle-only
   render-model tests are deleted when the parser oracle is deleted, while
   SmartArt fallback coverage lives in document/computed and adapter tests.
5. Public conversion tests, package verification, and VRT prove that
   `convertPptxToSvg` / `convertPptxToPng` remain document-path-only after the
   deletion.

No deletion is performed for #531. The current conclusion is to keep
`parser-path-oracle.ts`, `dual-reader-structural-comparison.test.ts`, and
`pptx-data-parser.ts` until the above replacement coverage exists.

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
- Parser unit tests under `packages/core/src/parser/` still protect the old
  oracle helpers. Delete or narrow them only after the corresponding oracle role
  is removed.
- `pptx-computed-view-renderer-adapter.ts` still calls renderer-specific fallback
  helpers for chart parsing. SmartArt now consumes the document computed diagram
  drawing contract, while the final mapping to renderer `GroupElement` remains
  in core. Chart XML conversion is isolated in `renderer-chart-data-converter.ts`;
  keep the renderer `ChartData` adapter outside `document` unless it is replaced
  by a renderer-owned contract. Split follow-up issues should define
  document-owned chart source contracts only after chart data, style/color parts,
  embedded workbook references, and compatibility expectations are covered.
- Comments in `packages/document/src/computed/create-computed-view.ts`
  that mention current-parser compatibility are retained as parity signposts.
  They should be removed only when the document path intentionally owns a
  different behavior and tests are updated with that decision.

Suggested follow-up slices:

1. Introduce a chart source/computed contract in `document`, then replace or
   narrow `renderer-chart-data-converter.ts` from the core/renderer side without
   moving renderer `ChartData` ownership into `document`.
2. Retire `dual-reader-structural-comparison.test.ts` after VRT and public
   regression tests fully cover the parser oracle's remaining value.
4. Remove `pptx-data-parser.ts` and parser render-model tests after the explicit
   parser-path oracle is no longer used by VRT/CI.
