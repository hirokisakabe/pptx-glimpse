# PptxSourceModel and computed view

- Status: RFC decision for [#450](https://github.com/hirokisakabe/pptx-glimpse/issues/450)
- Date: 2026-06-20

This note records the PptxSourceModel layering decision that should feed into
[#445](https://github.com/hirokisakabe/pptx-glimpse/issues/445). It builds on
the package boundary decision in
[document-boundaries.md](./document-boundaries.md): `@pptx-glimpse/document` is
the lower-level OOXML / PptxSourceModel foundation, and renderer / core / editor-core /
pom are consumers of it.

> Naming note, 2026-06-27: this concept was formerly called CleanDoc. The
> current name is `PptxSourceModel` because the model is PPTX-domain source
> material for reader / writer / computed-view workflows, including raw OOXML
> preservation sidecars. `PptxSource` was rejected because it can sound like the
> whole `@pptx-glimpse/document` package or an input stream rather than the
> canonical in-memory source model. `PptxModel` was rejected because it is too
> broad and blurs the source model with `PptxComputedView`, renderer adapter
> models, and future editor projections.

The related raw preservation and round-trip decision is recorded in
[raw-ooxml-round-trip.md](./raw-ooxml-round-trip.md). In short, PptxSourceModel targets
structural preservation rather than byte equality, and raw OOXML is kept as
source-model sidecars or untouched package parts instead of becoming the primary
editing surface.

The related core dogfood migration decision is recorded in
[core-document-dogfood-migration.md](./core-document-dogfood-migration.md). In
short, the public SVG/PNG conversion path now routes through a computed view and
a core-owned adapter after the [#481](https://github.com/hirokisakabe/pptx-glimpse/issues/481)
default switch. The old parser remains only as an explicit internal parity
oracle and targeted fallback source.

The current PptxComputedView-to-renderer-model duplication boundary is recorded
in
[pptx-computed-renderer-model-boundaries.md](./pptx-computed-renderer-model-boundaries.md).
In short, PptxComputedView keeps source provenance and effective document
semantics, while the renderer model is a render-ready adapter target and should
not be merged back into PptxSourceModel.

## Decision

Adopt a two-layer PptxSourceModel design:

```text
OOXML package parts
  -> PptxSourceModel source model
  -> computed document view
  -> renderer adapter / editor projections / AI-readable output
```

The source model is the canonical document representation owned by
`@pptx-glimpse/document`. It keeps enough structure and source references for
writer, editor, and round-trip work.

The computed view is generated from the source model for consumers that need
effective values. It resolves inheritance, relationships, theme references, and
unit conversions into a deterministic, easy-to-consume view. It is derived data,
not the editable source of truth.

Do not merge the existing renderer model directly into PptxSourceModel. Treat it as the
first rendering-oriented computed view shape, then migrate toward an explicit
PptxSourceModel-to-renderer adapter.

## Source Model Responsibilities

The source model should preserve document semantics close to the PPTX package,
without exposing OOXML's scattered ZIP/XML layout as the primary API.

It owns:

- Presentation hierarchy: presentation, slides, layouts, masters, themes, media,
  charts, tables, notes, relationships, and content types.
- Stable source IDs and relationship references needed for incremental edits.
- Element identity and ordering as authored in each source part.
- Source-local values, including unresolved theme color references, style
  references, placeholder declarations, relationship IDs, and package part paths.
- Raw or partially parsed OOXML for constructs PptxSourceModel does not model directly,
  such as vendor extensions, `mc:AlternateContent`, uncommon DrawingML nodes, and
  future features.
- Lossless-preservation metadata where it matters for round-trip output:
  original part path, original relationship target, namespace-sensitive raw
  nodes, and source ranges or node IDs when available.
- Normalized source-level value wrappers when normalization removes OOXML noise
  but does not erase intent, for example typed EMU values instead of bare
  numbers.

The source model should not contain renderer decisions:

- No font fallback chosen from the local OS.
- No text-to-path, SVG, resvg, browser, or pixel-output decisions.
- No hidden-shape filtering that is only for a preview target.
- No destructive merge of master/layout/slide elements into a single rendered
  element list.

## Computed View Responsibilities

The computed view is a deterministic projection from source model plus an
explicit computation context. It can have multiple variants as long as they all
derive from the same source model.

It owns:

- Effective slide order, slide size, and per-slide master/layout chain.
- Theme color resolution from `schemeClr` through effective `clrMap` to concrete
  colors, including color transforms and alpha.
- Theme font token resolution to typefaces, but not environment-specific font
  substitution.
- Background fallback: slide -> layout -> master.
- Placeholder matching and cascade resolution for transform, geometry, and text
  styles.
- Master/layout/slide visibility decisions that are part of PowerPoint document
  semantics, such as `showMasterSp` and layout `showMasterSp`.
- Effective text properties across presentation defaults, master text styles,
  layout placeholder styles, and slide-local overrides.
- Relationship resolution from `rId` references to canonical package parts or
  embedded resource handles.
- Unit conversion into consumer-friendly values when the consumer requires it,
  while retaining source references where round-trip diagnostics need them.

The computed view is allowed to be lossy relative to source when the loss is
explicit and documented. For example, a render view can collapse a theme color to
`#RRGGBB` because the renderer needs a paint color, but the source model must
still preserve the original theme reference for writer/editor use.

## Source to Computed Boundary

The boundary should be an explicit projection API, not incidental mutation during
parse.

Recommended shape:

```text
readPptx(input) -> PptxSourceModel

createComputedView(source, options) -> PptxComputedView

createRenderView(computed, options) -> renderer model
```

`createComputedView` should be pure from the caller's perspective: it must not
rewrite the source model in place. Cache internal lookup tables if needed, but
expose the computed result as derived data.

The projection context should include:

- Target slide selection.
- Whether to include hidden slides or hidden elements.
- The desired unit view for consumers: source units, EMU, points, or pixels.
- Optional diagnostics mode that records where each effective value came from.

Renderer-only context stays outside `@pptx-glimpse/document`:

- System font discovery and fallback mapping.
- Text measurement and wrapping behavior that depends on available fonts.
- SVG/PNG output size decisions.
- Renderer warnings about unsupported visual features.

## Cascade Resolution

Resolve the PPTX cascade in the computed layer, not in the source model.

The source model should keep each part separate:

```text
presentation
theme
slide master
slide layout
slide
```

The computed view should expose effective values for a slide:

```text
presentation defaults
  -> theme + master color map
  -> master
  -> layout
  -> slide
  -> run / element local overrides
```

Specific rules:

- Theme color references remain references in source. Computed view resolves
  `schemeClr` through the effective `clrMap` and `colorScheme`.
- `clrMapOvr` belongs to the cascade resolver. Apply layout override before
  slide override.
- Background fallback is computed as slide first, then layout, then master.
- Placeholder transform and geometry are computed by matching layout
  placeholders first and falling back to master placeholders.
- Text styles are computed using layout placeholder style, master placeholder
  style, master `txStyles`, and presentation `defaultTextStyle`.
- Visibility and ordering for rendering should be computed without mutating the
  source element arrays. For rendering, this currently means master decorative
  elements, layout decorative elements, then slide elements, while template
  placeholders are excluded.

## Unit Normalization

Use typed units at source boundaries and make unit conversion explicit.

Recommended policy:

- Source model stores PPTX-native units as typed values where they carry domain
  meaning: EMU for coordinates, points for font sizes, hundredth-points for
  `spcPts`, OOXML percentages for values such as `spcPct`, and OOXML angle units
  where applicable.
- Do not store raw unbranded numbers for source-level geometry or typography.
- Computed document view should keep EMU / point values by default so writer and
  editor workflows do not accumulate pixel rounding.
- Render view may convert to pixels at the renderer boundary using a declared
  DPI policy. The current renderer uses 96 DPI, where a 16:9 slide
  `9144000 x 5143500` EMU becomes `960 x 540` px.
- When a computed value is converted, retain enough provenance in diagnostics to
  explain the source value and conversion policy.

This means unit normalization is not a single global conversion to pixels. It is
a typed-source policy plus explicit consumer-specific conversion.

## Relationship to the Current Renderer Model

The existing model in `packages/renderer/src/model/` already acts
like a render-oriented computed view:

| Current renderer model                               | Computed view role                                                                | Source model counterpart                                             |
| ---------------------------------------------------- | --------------------------------------------------------------------------------- | -------------------------------------------------------------------- |
| `Slide`                                              | Effective rendered slide with background and elements                             | Slide part plus links to layout/master                               |
| `ShapeElement` / `ConnectorElement` / `GroupElement` | Effective drawing element with transform, geometry, fill, line, effects           | Source shape tree with unresolved refs and raw OOXML                 |
| `TextBody`, `Paragraph`, `TextRun`                   | Effective text body with inherited run/paragraph properties resolved where needed | Source text body with local `a:pPr`, `a:rPr`, `lstStyle`, defaults   |
| `Fill`, `Outline`, `ResolvedColor`                   | Paint-ready values for renderer                                                   | Source fill/line nodes with scheme colors, style refs, and raw nodes |
| `ImageElement`                                       | Renderable media payload with resolved image data and crop/tile info              | Source blip relationship plus media part reference                   |
| `ChartElement`                                       | Simplified chart data for current chart renderer                                  | Source chart part, embedded workbook refs, style refs, raw chart XML |
| `TableElement`                                       | Effective grid/cell data for rendering                                            | Source table XML with merge/style references and raw table nodes     |
| `Theme`                                              | Parsed theme values used to compute colors/fonts                                  | Theme part as source, with scheme and format/style data preserved    |

Migration direction:

1. Keep the current renderer model as the SVG/PNG render view contract.
2. Introduce PptxSourceModel source types separately in `@pptx-glimpse/document`.
3. Move source-preserving parse responsibility into `document`.
4. Add a computed view generator that resolves cascade and references.
5. Add an adapter from computed view to the current renderer model.
6. Only after that, simplify or replace renderer model fields that duplicate
   computed view fields.

## Historical Implementation Mapping Before #481

Before the PptxSourceModel default switch, parser/core code already contained several
computed-view behaviors:

- `parseSlideWithLayout` resolves slide -> layout -> master chains.
- `ColorResolver` resolves `schemeClr` through `clrMap` and `colorScheme`.
- `resolveSlideColorResolver` applies layout and slide `clrMapOvr`.
- `parseSlideWithLayout` applies slide -> layout -> master background fallback.
- `mergePlaceholderGeometry` resolves layout/master placeholder geometry.
- `applyTextStyleInheritance` resolves text style inheritance.
- `converter.ts` merges master, layout, and slide elements for rendering and
  filters template placeholders.

Under the two-layer design, these behaviors moved behind explicit computed-view
generation APIs for public conversion. Any remaining old-parser overlap is
scoped to the parser oracle or renderer-specific adapter fallbacks rather than
the public default path.

## Schema Normalization Scope

Normalize enough to make PptxSourceModel usable, but do not normalize away source
intent.

Normalize:

- Package paths and relationships into stable document references.
- Namespaced OOXML names into PptxSourceModel field names.
- Common geometry, transform, fill, line, text, table, chart, and media
  structures.
- Units into typed domain values.
- Element identity and ordering.

Preserve:

- Original relationship IDs and part paths when useful for round-trip output.
- Raw OOXML for unsupported or partially supported nodes.
- Theme references, style references, placeholder references, and source-local
  overrides.
- Ordering and extension nodes that may affect writer output.

Do not normalize source into:

- A single flattened slide element list.
- Renderer-specific fill/text/image structures only.
- Pixels as the canonical geometry unit.
- pom authoring primitives such as flex layout concepts.

## Conclusion for #445

The conclusion to reflect in #445 is:

PptxSourceModel should adopt a two-layer architecture. The canonical
`@pptx-glimpse/document` model is a source model that preserves document
semantics, source references, typed OOXML-domain units, and raw escape hatches
for writer/editor/round-trip work. A computed view is generated from that source
model to resolve theme, master, layout, slide, placeholder, relationship, and
text-style cascades for renderer, AI reading, and editor projections.

The existing renderer model should be treated as a render-specific computed view
or adapter target, not as the PptxSourceModel schema itself. Unit normalization should
use typed PPTX-domain units in source and computed document views, with pixel
conversion limited to renderer boundaries. If a future design rejects the
two-layer approach, it must explain how one model can simultaneously preserve
round-trip source details and expose effective values without becoming both
lossy for writer/editor use and too heavy for renderer/AI consumers.
