# PptxComputedView and renderer model boundaries

- Status: implementation note for [#530](https://github.com/hirokisakabe/pptx-glimpse/issues/530)
- Date: 2026-06-27

This note records the current intentional duplication between
`@pptx-glimpse/document` source types, `PptxComputedView`, and the renderer
model in `packages/renderer/src/model/`. It builds on the package boundary
decision in [document-boundaries.md](./document-boundaries.md), the
source/computed layering decision in
[pptx-source-model-computed-view.md](./pptx-source-model-computed-view.md), and
the post-switch parser audit in
[legacy-parser-semantics-audit.md](./legacy-parser-semantics-audit.md).

The public conversion path currently flows through:

```text
PPTX package
  -> PptxSourceModel source model
  -> PptxComputedView
  -> core-owned renderer adapter
  -> renderer Slide / SlideElement model
```

The renderer model remains a display-oriented adapter target. It should not be
merged directly into PptxSourceModel, and `@pptx-glimpse/document` should not
import renderer types.

## Model correspondence

| Source model owner                                                                                                                   | Computed view owner                                                                                                                                                                                 | Renderer model target                                                                              | Boundary status                                                                                                                                                                                                   |
| ------------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PptxSourceModel`, `SourcePresentation`, `SourceSlide`, `SourceSlideLayout`, `SourceSlideMaster`, `SourceTheme`                      | `PptxComputedView`, `ComputedSlideSize`, `ComputedSlide`                                                                                                                                            | `SlideSize`, `Slide`, `Background`                                                                 | Intentional boundary. Source keeps package hierarchy and source handles, computed resolves effective slide values, renderer receives a render contract.                                                           |
| `SourceShape`, `SourceConnector`, `SourceGroup`, `SourceImage`, `SourceTable`, `SourceChart`, `SourceSmartArt`, `SourceRawShapeNode` | `ComputedShapeElement`, `ComputedConnectorElement`, `ComputedGroupElement`, `ComputedImageElement`, `ComputedTableElement`, `ComputedChartElement`, `ComputedSmartArtElement`, `ComputedRawElement` | `ShapeElement`, `ConnectorElement`, `GroupElement`, `ImageElement`, `TableElement`, `ChartElement` | Intentional boundary with some future shrink opportunities in the adapter. Computed elements keep `sourceLayer`, `sourcePartPath`, and `sourceNode`; renderer elements are pure renderable objects.               |
| `SourceTransform`                                                                                                                    | `SourceTransform` reused on computed elements                                                                                                                                                       | `Transform`                                                                                        | Intentional boundary. Source/computed transform can be absent and uses source naming; renderer transform is required and includes renderer defaults.                                                              |
| `SourceFill`, `SourceColor`, `SourceOutline`                                                                                         | `ComputedFill`, `ComputedColor`, `ComputedOutline`                                                                                                                                                  | `Fill`, `ResolvedColor`, `Outline`                                                                 | Intentional boundary. Computed resolves colors and relationships while retaining source provenance; renderer receives paint-ready values or `null`.                                                               |
| `SourceTextBody`, `SourceParagraph`, `SourceTextRun`, source text properties                                                         | `ComputedTextBody`, `ComputedParagraph`, `ComputedTextRun`, computed text properties                                                                                                                | `TextBody`, `Paragraph`, `TextRun`, renderer text properties                                       | Mixed. Text cascade resolution belongs in computed view; renderer default materialization and hyperlink/text-outline fields remain renderer-specific. Some naming and default-shape duplication can shrink later. |
| `SourceTable`, `SourceTableCell`, table source styles                                                                                | `ComputedTableData`, `ComputedTableCell`, computed cell borders/fills/text                                                                                                                          | `TableElement`, `TableData`, `TableCell`, `CellBorders`                                            | Mixed. Table structure and merge state duplicate by design today; future table style/source contracts can reduce adapter-only shape conversions without moving renderer table rendering into `document`.          |
| `SourceEffectList`, source blip effects                                                                                              | `ComputedEffectList`, `ComputedBlipEffects`                                                                                                                                                         | `EffectList`, `BlipEffects`                                                                        | Intentional boundary. Computed resolves colors in effects; renderer uses `null` defaults and render-ready effect objects.                                                                                         |
| `SourceChart` plus chart relationship/package data                                                                                   | `ComputedChartElement` with resolved chart XML                                                                                                                                                      | `ChartElement`, `ChartData`                                                                        | Intentional renderer-specific fallback. Chart XML to renderer `ChartData` conversion stays outside `document` until a document-owned chart source/computed contract exists.                                       |
| `SourceSmartArt` plus diagram data/drawing package data                                                                              | `ComputedSmartArtElement` with resolved raw drawing XML and media                                                                                                                                   | Renderer `GroupElement` produced through adapter fallback                                          | Temporary bridge. The current parser-backed fallback is intentionally scoped and tracked by [#535](https://github.com/hirokisakabe/pptx-glimpse/issues/535).                                                      |
| `SourceRawShapeNode`, raw sidecars, raw backgrounds/fills                                                                            | `ComputedRawElement`, raw computed background/fill variants                                                                                                                                         | No direct renderer equivalent; adapter warning and skip/ignore behavior                            | Intentional boundary. Raw material is for preservation and diagnostics, not direct SVG/PNG output.                                                                                                                |

## `ComputedSlide` vs renderer `Slide`

`ComputedSlide` is a document-derived effective slide projection. It still
carries source context:

- `partPath`, `layoutPartPath`, `masterPartPath`, and `themePartPath` identify
  the source package chain.
- `relationships` expose resolved relationship targets, target part paths, and
  media handles.
- `colorMap` and `colorScheme` keep the effective theme context available to
  adapter fallbacks such as chart and SmartArt conversion.
- `background` records the resolved background plus its source layer and source
  background object.
- `showMasterShapes` and `layoutShowMasterShapes` record document visibility
  decisions before renderer mapping.
- `elements` carry `sourceLayer`, `sourcePartPath`, and `sourceNode` for
  provenance and diagnostics.

Renderer `Slide` is a render contract:

- It has only `slideNumber`, `background`, `elements`, and `showMasterSp`.
- It does not expose source part paths, source nodes, relationships, raw OOXML,
  or theme provenance.
- Its `background` and element fields use `null` to mean "not renderable or not
  present" because renderer code is optimized for drawing decisions.
- It expects each element to already be in renderable order with required
  renderer defaults materialized by the adapter.

The duplication is therefore intentional. `ComputedSlide` answers "what did the
PPTX document mean after cascade and relationship resolution?" Renderer `Slide`
answers "what should the SVG/PNG renderer draw?"

## Major duplication classification

### Fill and color

`SourceFill` and `SourceColor` preserve authoring intent: theme references,
system colors, OOXML color transforms, image relationship IDs, raw fallback, and
raw sidecars. `ComputedFill` resolves colors to `ComputedColor`, resolves image
relationships to package media when possible, and keeps `source` for
provenance. Renderer `Fill` contains render-ready paint values:
`ResolvedColor`, base64 `imageData`, MIME type, and tile data.

This boundary is intentional. Future adapter shrink can align naming and helper
functions, but renderer `Fill` should not become the source or computed fill
schema because it has already lost theme/style provenance and raw preservation
hooks.

### Outline

`SourceOutline` can omit width and fill because OOXML can rely on inherited
style references or defaults. `ComputedOutline` keeps the source outline and a
computed fill when available. Renderer `Outline` requires a concrete width,
dash style, arrow endpoint defaults, and a renderable solid/gradient fill or
`null`.

This is an intentional boundary. The adapter owns default materialization for
the current renderer contract. If renderer line defaults are redesigned, that
should be a renderer-focused change and not a reason to move renderer `Outline`
into `document`.

### Transform

`SourceTransform` stores PPTX-domain values from `a:xfrm` using EMU and OOXML
angle units. It can be absent because some source nodes rely on placeholder or
group context. Computed elements keep that source-shaped transform after cascade
resolution. Renderer `Transform` is required for renderable elements and uses
renderer field names (`extentWidth`, `extentHeight`, `flipH`, `flipV`) plus
numeric defaults.

This is an intentional boundary with a narrow future shrink opportunity:
adapter-local conversion helpers can be simplified, but source/computed
transforms should remain optional/provenance-friendly while renderer transforms
remain required for drawing.

### Text

Source text types preserve source-local paragraph/run properties, list styles,
theme font tokens, raw sidecars, and edit handles. Computed text resolves the
text style cascade and color values while keeping document-domain units and
source-shaped properties. Renderer text materializes the defaults needed by text
measurement and SVG generation: body margins, anchor/wrap/autofit defaults,
paragraph defaults, `RunProperties`, hyperlink slots, and text outline slots.

This is mixed duplication. The source/computed/renderer split is intentional,
but renderer and computed text property names can converge when there is a
separate text-model cleanup. That cleanup must keep environment-specific font
fallback, text measurement, wrapping, and text-to-path behavior outside
`document`.

### Table

Source tables preserve table XML structure, row/column sizes, merge flags,
cell text/fill/borders, table style IDs, raw sidecars, and source handles.
Computed tables resolve cell text, fill, and border values into computed paint
and text types. Renderer tables receive `TableElement`, `TableData`, and
`TableCell` objects with renderer `Fill`, `Outline`, and `TextBody` values or
`null`.

This is mixed duplication. Current duplication is acceptable because table
rendering is still renderer-oriented while document table source/style coverage
is incomplete. A future table source/computed contract may reduce adapter logic,
but renderer table drawing details should remain renderer-owned.

### Effect

Source effects preserve OOXML-domain values and source colors. Computed effects
resolve effect colors and blip color effects into computed colors. Renderer
effects use render-ready color values and `null` defaults for absent effects.

This is an intentional boundary. The renderer may continue to have a compact
effect list optimized for SVG generation, while `document` keeps source
provenance and raw preservation separate.

### Chart

`SourceChart` currently stores the source relationship identity for the chart
graphic frame. `ComputedChartElement` resolves the relationship enough to expose
chart XML to core. The core adapter then calls
`convertChartXmlToRendererChartData`, producing renderer `ChartData`.

This is an intentional renderer-specific fallback, not a PptxSourceModel chart
schema. A future document-owned chart source/computed contract should model
chart parts, style/color parts, embedded workbook references, relationship
graphs, raw preservation, and diagnostics before the adapter is narrowed.
Renderer `ChartData` should not be moved into `document`.

### SmartArt and raw elements

SmartArt is currently a raw-resolved fallback. `ComputedSmartArtElement`
provides the resolved diagram drawing XML, drawing relationships, and media.
The adapter combines that element data with the surrounding `ComputedSlide`
color context and maps it into a renderer `GroupElement` using the temporary
`parseShapeTree` bridge documented in
[smartart-fallback-contract.md](./smartart-fallback-contract.md).

Raw elements, raw backgrounds, and raw fills have no renderer model equivalent.
The adapter warns and skips or ignores them because raw OOXML preservation is a
source/writer concern, not a rendering contract.

## Why the renderer model should not be merged into PptxSourceModel

Merging the renderer model into PptxSourceModel would conflict with the
decisions in `document-boundaries.md` and
`pptx-source-model-computed-view.md`:

- Renderer fields contain visual fallback decisions and `null` defaults that are
  useful for SVG/PNG output but lossy for writer/editor workflows.
- Renderer media fields use base64 payloads and MIME types, while source and
  computed views need package part paths, relationship IDs, and media handles.
- Renderer text fields are shaped for measurement, wrapping, and text-to-path
  conversion; source text must retain theme tokens, source handles, raw
  sidecars, and edit provenance.
- Renderer elements do not retain source layer, source part path, raw OOXML, or
  relationship provenance needed for diagnostics and round-trip preservation.
- Renderer chart and SmartArt fallbacks are compatibility targets for current
  rendering behavior, not canonical document semantics.
- Making `document` depend on renderer types would reverse the intended package
  dependency direction.

The correct boundary remains:

```text
PptxSourceModel
  -> PptxComputedView
  -> core adapter
  -> renderer model
```

## Follow-up separation

This issue is documentation-only and does not require immediate model
redesign. Required implementation follow-ups stay separate:

- SmartArt fallback replacement is tracked by
  [#535](https://github.com/hirokisakabe/pptx-glimpse/issues/535).
- Chart source/computed design should be split from renderer `ChartData` changes
  when chart editing or deeper chart semantics become in scope; the current
  renderer-specific chart conversion was isolated by
  [#529](https://github.com/hirokisakabe/pptx-glimpse/issues/529).
- Parser oracle retirement, table/text model cleanup, and renderer default
  simplification should remain separate rendering or parser-retirement PRs so
  they can be tested without changing the PptxSourceModel boundary.

No new follow-up issue is required for this documentation pass because the only
immediate adapter replacement called out by the current audit is already tracked
by #535. New issues should be created when a specific chart, table, text, or
renderer-default redesign is selected for implementation.
