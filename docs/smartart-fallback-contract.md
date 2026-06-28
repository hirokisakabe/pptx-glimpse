# SmartArt fallback diagram drawing contract

- Status: implementation note for [#528](https://github.com/hirokisakabe/pptx-glimpse/issues/528)
  and [#535](https://github.com/hirokisakabe/pptx-glimpse/issues/535)
- Date: 2026-06-27
- Updated: 2026-06-28

This note records the SmartArt fallback boundary after public SVG/PNG conversion
moved to the PptxSourceModel document path and after the adapter stopped calling
the old parser `parseShapeTree` helper for resolved diagram drawing XML. It
complements the package/source/computed boundaries in
[document-boundaries.md](./document-boundaries.md) and
[pptx-source-model-computed-view.md](./pptx-source-model-computed-view.md).

## Current Call Chain

The public conversion path reaches SmartArt fallback through this sequence:

```text
convertPptxToSvg / convertPptxToPng
  -> readPptx
  -> createComputedView
  -> ComputedSmartArtElement.diagramDrawing
  -> adaptComputedViewToRendererModel
  -> adaptSmartArt
  -> renderer GroupElement
```

The fallback remains limited to the already-resolved diagram drawing part. It is
not the old slide parser path and it does not rebuild the slide, layout, or
master cascade.

## Parser Dependency Status

`packages/core/src/pptx-computed-view-renderer-adapter.ts` does not import the
retired core render parser for SmartArt fallback. The old adapter dependency on
parser XML helpers, parser relationship maps, `navigateOrdered`, and
`parseShapeTree` was replaced by a document-owned computed diagram drawing
contract.

## Input Contract

`@pptx-glimpse/document` provides the source/computed input boundary:

| Field                                                  | Owner                  | Contract                                                                                                                        |
| ------------------------------------------------------ | ---------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| `SourceSmartArt`                                       | document source model  | Represents the SmartArt `graphicFrame` source node, including the data model relationship ID and source transform when present. |
| `ComputedSmartArtElement.dataRelationship`             | document computed view | Resolves the slide relationship to the diagram data part.                                                                       |
| `ComputedSmartArtElement.drawingRelationship`          | document computed view | Resolves the diagram data part relationship to the diagram drawing part.                                                        |
| `ComputedSmartArtElement.drawingPartPath`              | document computed view | Gives the canonical package part path for the resolved drawing XML.                                                             |
| `ComputedSmartArtElement.drawingXml`                   | document computed view | Exposes the raw drawing XML text for compatibility and diagnostics.                                                             |
| `ComputedSmartArtElement.drawingRelationships`         | document computed view | Resolves relationships owned by the drawing part, including media references.                                                   |
| `ComputedSmartArtElement.media`                        | document computed view | Exposes package media bytes so drawing-part image relationships can be rendered.                                                |
| `ComputedSmartArtElement.diagramDrawing`               | document computed view | Exposes the parsed computed diagram drawing view used by the adapter.                                                           |
| `ComputedSlide.colorScheme` / `ComputedSlide.colorMap` | document computed view | Provide document-level theme color context for computed child fills/outlines/text.                                              |

`ComputedDiagramDrawing` contains:

- `sourcePartPath`: the resolved diagram drawing part path.
- `rawXml`, `rawPart`, and `rawHandle`: raw preservation/provenance handles for
  the source part.
- `relationships`: relationships owned by the drawing part.
- `media`: package media references available to computed image children.
- `childTransform`: the root `spTree.grpSpPr.xfrm` child coordinate system.
- `children`: ordered computed child elements parsed from the drawing shape tree.
- `diagnostics`: document-level diagnostics with source part provenance, such as
  missing `spTree`.

The child elements reuse the existing document shape tree reader and computed
element pipeline. Drawing-part relationships are used while computing diagram
children, so `p:pic` and image fills inside the drawing resolve against the
diagram drawing part instead of the outer slide part.

## Output Contract

`packages/core/src/pptx-computed-view-renderer-adapter.ts` owns the renderer
output contract:

1. If `diagramDrawing` is missing, emit
   `pptx-computed-view-adapter.unresolved-smartart-skipped` and skip the element.
2. If `diagramDrawing.diagnostics` contains a warning, emit
   `pptx-computed-view-adapter.unresolved-smartart-skipped` with the diagram
   drawing source part path and skip the element.
3. Convert `diagramDrawing.children` through the normal computed-element adapter
   path.
4. If no renderer-supported children are produced, emit
   `pptx-computed-view-adapter.unresolved-smartart-skipped` and skip the element.
5. Return a renderer `GroupElement`:
   - `transform` comes from the outer SmartArt `graphicFrame`.
   - `childTransform` comes from `diagramDrawing.childTransform`, falling back to
     the outer group extent.
   - `children` are the adapted renderer `SlideElement[]`.
   - `effects` is `null`.
   - `altText` comes from the `SourceSmartArt` name when present.

The renderer receives an ordinary `GroupElement`. It has no SmartArt-specific
knowledge and does not parse OOXML, resolve package relationships, or know
PptxSourceModel types.

## Historical Parser Bridge

Before #535, the adapter parsed `drawingXml` locally and called
`packages/core/src/parser/slide-parser.ts` `parseShapeTree`. That bridge gave
SmartArt fallback old-parser behavior for z-order, shapes, pictures, connectors,
groups, text, fills, outlines, effects, image relationships, compatibility
branches, and group fill inheritance.

The replacement keeps those responsibilities split by owner:

```text
diagram drawing part in @pptx-glimpse/document computed view
  -> computed diagram drawing view with ordered children, relationships, media,
     raw handles, group fill inheritance, and diagnostics provenance
  -> core adapter mapping into renderer GroupElement/children
  -> renderer
```

Do not move the fallback into renderer. Renderer owns SVG/PNG output and a
display-oriented model; document/core own PPTX package interpretation and
adapter mapping.

## Removal Status

The old `parseShapeTree` SmartArt adapter dependency is removed when all of
these remain true:

1. `@pptx-glimpse/document` exposes diagram drawing computed data for the
   resolved drawing part, including ordered child nodes, source part paths,
   relationships, media references, raw preservation handles, and diagnostics
   provenance.
2. The computed view exposes effective diagram drawing data needed by the
   adapter without depending on renderer types.
3. The core adapter maps the computed diagram drawing view to renderer
   `GroupElement` children without importing `packages/core/src/parser/*`.
4. Focused tests cover SmartArt fallback ordering, group child transforms,
   shape fills/outlines/text, media relationships, group fill inheritance, and
   skip diagnostics.
5. Visual output for existing SmartArt fixtures remains intentionally unchanged,
   or any intentional change is covered by VRT updates in a separate rendering
   PR.

The old parser shape-tree tests were removed with the parser oracle. SmartArt
fallback coverage now lives in focused document/core adapter tests plus VRT.
