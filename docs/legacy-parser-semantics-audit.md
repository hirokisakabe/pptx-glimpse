# Legacy parser semantics audit after the CleanDoc default switch

- Status: implementation note for [#485](https://github.com/hirokisakabe/pptx-glimpse/issues/485)
- Date: 2026-06-27

This note records the first shrink pass after public SVG/PNG conversion moved to
the CleanDoc document path. It complements
[core-document-dogfood-migration.md](./core-document-dogfood-migration.md) and
[document-boundaries.md](./document-boundaries.md).

## Current owner split

The public conversion path is now:

```text
convertPptxToSvg / convertPptxToPng
  -> convertPptxToSvgViaDocumentPath / convertPptxToPngViaDocumentPath
  -> @pptx-glimpse/document readPptx
  -> createComputedView
  -> core-owned CleanDoc renderer adapter
  -> renderer
```

The old parser path is no longer part of public conversion orchestration. Its
remaining render entry points live in
`packages/pptx-glimpse/src/parser-path-oracle.ts` and are intentionally scoped as
an internal oracle for parity checks.

## Semantics moved to `document`

These reusable OOXML semantics are now owned by `@pptx-glimpse/document` source
or computed view code:

| Semantics | Document owner | Previous parser overlap removed or reduced |
| --- | --- | --- |
| Package graph, slide order, relationships, media payloads | `reader/read-pptx.ts`, source package graph, computed relationships | Public conversion no longer calls `parsePptxData` for package traversal |
| Slide/layout/master chain resolution | `createComputedView` | Public conversion no longer calls `parseSlideWithLayout` for cascade assembly |
| Background fallback | `createComputedView` | Public conversion uses computed `background` instead of parser-side fallback |
| Theme color and color map resolution | source theme data + computed color resolution | Public conversion uses computed colors through the adapter |
| Placeholder filtering and effective element ordering | `createComputedView` | Core converter no longer owns parser-path `mergeElements` |
| Text style and theme font resolution for rendering/inspection | `createComputedView` | `collectUsedFonts` now reads CleanDoc/computed text instead of parser slides |
| Table/chart/image relationship resolution for rendering | `createComputedView` + adapter | Parser path remains only as comparison oracle |

## Responsibilities left outside `document`

These remain in core or renderer by design:

| Responsibility | Owner | Reason |
| --- | --- | --- |
| Public API options, warning setup, font setup, SVG text-output mode, PNG sizing | `packages/pptx-glimpse/src/experimental-document-renderer.ts` | API compatibility and renderer environment setup are core concerns |
| CleanDoc computed view to renderer model mapping | `packages/pptx-glimpse/src/cleandoc-renderer-adapter.ts` | The renderer model is a display-oriented compatibility target, not the CleanDoc source model |
| Chart XML to renderer chart model mapping | Adapter calling parser `parseChart` | Chart rendering model is still renderer-specific; move only after a chart source/computed contract exists |
| SmartArt drawing XML fallback to renderer shape tree | Adapter calling parser `parseShapeTree` | This is a rendering fallback for resolved diagram drawing XML, not canonical document semantics yet |
| Font discovery, font mapping, text measurement, text-to-path, SVG/PNG output | `pptx-glimpse-renderer` | Renderer-specific behavior per `document-boundaries.md` |

## Remaining legacy parser uses

The old parser code still has three explicit roles:

1. `parser-path-oracle.ts` renders through the old parser for document-path
   parity checks.
2. `dual-reader-structural-comparison.test.ts` compares a focused structural
   subset against the old parser while the migration still needs a readable
   model-level oracle.
3. `cleandoc-renderer-adapter.ts` reuses `parseChart` and `parseShapeTree` for
   renderer-specific chart and SmartArt fallbacks.

No public API currently imports old parser render orchestration. `converter.ts`
only calls the document path.

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
- Parser unit tests under `packages/pptx-glimpse/src/parser/` still protect the
  old oracle and adapter fallback helpers. Delete or narrow them only after the
  corresponding oracle role is removed.
- `cleandoc-renderer-adapter.ts` still calls parser helpers for chart parsing
  and SmartArt fallback rendering. Split follow-up issues should define
  document-owned chart/diagram source contracts before moving this logic.
- Comments in `packages/pptx-glimpse-document/src/computed/create-computed-view.ts`
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
