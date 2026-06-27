# SmartArt fallback contract before `parseShapeTree` replacement

- Status: implementation note for [#528](https://github.com/hirokisakabe/pptx-glimpse/issues/528)
- Date: 2026-06-27
- Follow-up replacement issue:
  [#535](https://github.com/hirokisakabe/pptx-glimpse/issues/535)

This note records the current SmartArt fallback boundary after public SVG/PNG
conversion moved to the PptxSourceModel document path. It complements
[legacy-parser-semantics-audit.md](./legacy-parser-semantics-audit.md) and the
package/source/computed boundaries in
[document-boundaries.md](./document-boundaries.md) and
[pptx-source-model-computed-view.md](./pptx-source-model-computed-view.md).

This note does not replace the fallback implementation. It defines the contract
needed to decide where that replacement should live.

## Current call chain

The public conversion path reaches SmartArt fallback through this sequence:

```text
convertPptxToSvg / convertPptxToPng
  -> readPptx
  -> createComputedView
  -> adaptComputedViewToRendererModel
  -> adaptSmartArt
  -> parseShapeTree
  -> renderer GroupElement
```

The fallback is limited to the already-resolved diagram drawing part. It is not
the old slide parser path and it does not rebuild the slide, layout, or master
cascade.

## Current old parser dependencies

`packages/core/src/pptx-computed-view-renderer-adapter.ts` currently depends on
these old parser helpers for SmartArt fallback:

| Dependency                                                            | Current use                                                                               |
| --------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| `parseXml` from `packages/core/src/parser/xml-parser.ts`              | Parse the resolved diagram drawing XML and find `drawing.spTree`.                         |
| `parseXmlOrdered` from `packages/core/src/parser/xml-parser.ts`       | Parse the same drawing XML with preserve-order semantics.                                 |
| `navigateOrdered` from `packages/core/src/parser/slide-parser.ts`     | Locate ordered `drawing.spTree` children so z-order follows XML order.                    |
| `parseShapeTree` from `packages/core/src/parser/slide-parser.ts`      | Convert the diagram drawing `spTree` into renderer `SlideElement[]`.                      |
| `Relationship` from `packages/core/src/parser/relationship-parser.ts` | Adapt computed relationships into the map shape expected by `parseShapeTree`.             |
| `ColorResolver` from `packages/core/src/color/color-resolver.ts`      | Recreate a parser-compatible resolver from the computed slide color scheme and color map. |

Calling `parseShapeTree` also pulls in its parser-side implementation fan-out:

| Parser behavior reached through `parseShapeTree`                                               | Why SmartArt fallback currently gets it                                                                                                                                                                |
| ---------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Ordered shape tree traversal and `mc:AlternateContent` first-choice traversal                  | Diagram drawings must preserve child z-order and compatibility-branch content where the old parser already knows how.                                                                                  |
| `p:sp` parsing                                                                                 | SmartArt drawing shapes need transform, preset/custom geometry, fill, outline, effects, text body, hyperlinks, and alt text mapped to renderer shapes.                                                 |
| `p:pic` parsing                                                                                | Diagram drawings can contain generated raster/vector picture elements with drawing-part relationships to media.                                                                                        |
| `p:cxnSp` parsing                                                                              | Diagram drawings can contain connectors between generated nodes.                                                                                                                                       |
| `p:grpSp` recursion                                                                            | Diagram drawings can contain nested groups with child transforms and inherited group fills.                                                                                                            |
| `p:graphicFrame` parsing                                                                       | The parser may encounter table/chart/diagram graphic frames while walking the shape tree, though the adapter's synthetic archive only makes already-resolved drawing XML and media reliably available. |
| Fill, outline, shape style, effects, blip effects, text style, table, chart, and image helpers | These are not imported directly by the adapter, but they are part of the renderer model conversion provided by `parseShapeTree`.                                                                       |

The adapter passes `undefined` for parser `fontScheme`, `fmtScheme`, and
`placeholderStyles`. Therefore SmartArt fallback currently relies on the
diagram drawing's local formatting plus the computed slide color scheme/map; it
does not inherit placeholder geometry or theme font/style scheme data through
this old parser call.

## Input contract

`@pptx-glimpse/document` currently provides the source/computed input boundary:

| Field                                                  | Owner                  | Contract                                                                                                                        |
| ------------------------------------------------------ | ---------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| `SourceSmartArt`                                       | document source model  | Represents the SmartArt `graphicFrame` source node, including the data model relationship ID and source transform when present. |
| `ComputedSmartArtElement.dataRelationship`             | document computed view | Resolves the slide relationship to the diagram data part.                                                                       |
| `ComputedSmartArtElement.drawingRelationship`          | document computed view | Resolves the diagram data part relationship to the diagram drawing part.                                                        |
| `ComputedSmartArtElement.drawingPartPath`              | document computed view | Gives the canonical package part path for the resolved drawing XML.                                                             |
| `ComputedSmartArtElement.drawingXml`                   | document computed view | Exposes the raw drawing XML text for the resolved drawing part.                                                                 |
| `ComputedSmartArtElement.drawingRelationships`         | document computed view | Resolves relationships owned by the drawing part, including media references.                                                   |
| `ComputedSmartArtElement.media`                        | document computed view | Exposes package media bytes so drawing-part image relationships can be rendered.                                                |
| `ComputedSlide.colorScheme` / `ComputedSlide.colorMap` | document computed view | Provide document-level theme color context for a parser-compatible color resolver.                                              |

This is intentionally a raw-resolved fallback contract. `document` resolves the
package graph and provides the source material, but it does not yet expose a
typed diagram drawing source model or a computed diagram render view.

## Output contract

`packages/core/src/pptx-computed-view-renderer-adapter.ts` owns the current
output contract:

1. If `drawingXml` or `drawingPartPath` is missing, emit
   `pptx-computed-view-adapter.unresolved-smartart-skipped` and skip the element.
2. Parse `drawingXml` and require `drawing.spTree`; if absent, emit the same
   warning code and skip the element.
3. Convert `drawingRelationships` into the legacy parser relationship map.
4. Build a minimal archive facade:
   - `files.get(path)` returns `drawingXml` only for `drawingPartPath`.
   - `files.has(path)` is true only for `drawingPartPath`.
   - `media.get(path)` resolves bytes from `ComputedSmartArtElement.media`.
5. Build a `ColorResolver` from the computed slide color scheme and color map.
6. Call `parseShapeTree` with the diagram `spTree`, drawing relationships,
   drawing part path, archive facade, color resolver, fill context, and ordered
   children.
7. If no children are produced, skip the element.
8. Return a renderer `GroupElement`:
   - `transform` is the SmartArt source transform adapted at the outer
     `graphicFrame` boundary.
   - `childTransform` is read from `drawing.spTree.grpSpPr.xfrm` and falls back
     to the outer group extent.
   - `children` are the renderer `SlideElement[]` produced by `parseShapeTree`.
   - `effects` is `null`.
   - `altText` comes from the `SourceSmartArt` name when present.

The renderer receives an ordinary `GroupElement`. It has no SmartArt-specific
knowledge and should not need to parse OOXML, resolve package relationships, or
know PptxSourceModel types.

## Owner options

### `document` source/computed contract

This is the right long-term owner for package-level diagram drawing source
material. Diagram drawing XML, its relationships, media references, raw
preservation, and stable source handles are document-package semantics. A future
document-owned contract should expose a typed diagram drawing source tree, plus a
computed view with effective values that are document semantics.

It should not emit renderer `GroupElement` directly. Renderer model shape,
warnings, pixel output choices, font fallback, and SVG/PNG behavior remain
outside `document`.

### core adapter-local parser

This is the acceptable interim owner if the old parser must be retired before a
typed document diagram drawing model exists. The adapter could own a narrow
diagram drawing shape-tree parser that maps the raw resolved drawing XML into
the current renderer model.

The tradeoff is duplication: such a parser would need to copy or reimplement
shape, image, connector, group, fill, text, effect, and ordering behavior that
the old parser already provides. It would remove the cross-boundary import from
`parser/slide-parser.ts`, but it would not create reusable document semantics.

### renderer fallback

This is not the right owner. The renderer should consume render-oriented model
objects, not OOXML package parts. Moving SmartArt fallback into renderer would
make renderer depend on XML parsing, PPTX relationship resolution, package part
paths, media lookup, and theme color maps. That conflicts with the boundary that
renderer owns SVG/PNG output and display-oriented rendering, while core/document
own PPTX package interpretation.

## Decision

Keep the current SmartArt fallback in the core adapter until there is a typed
diagram drawing source/computed contract in `@pptx-glimpse/document`.

The replacement target should be:

```text
diagram drawing part in @pptx-glimpse/document source model
  -> computed diagram drawing view with resolved relationships/media/colors
  -> core adapter mapping into renderer GroupElement/children
  -> renderer
```

Do not move the fallback into renderer. Use a core adapter-local parser only as
a temporary bridge if parser retirement is blocked before the document-owned
diagram contract is ready.

## Why `parseShapeTree` remains for now

The old helper remains temporarily because it is currently the only local code
that converts a DrawingML shape tree into the existing renderer model with
z-order, groups, shapes, pictures, connectors, text, fills, outlines, effects,
image relationships, and compatibility branches.

Keeping it avoids an output-changing rewrite in this design issue and keeps
SmartArt fallback behavior tied to the same tested shape-tree semantics as the
parser oracle. The dependency is acceptable only while it is explicitly
classified as a renderer-specific fallback and not as reusable document
semantics.

## Conditions for removal

Stop using old `parseShapeTree` for SmartArt fallback when all of these are
true:

1. `@pptx-glimpse/document` exposes diagram drawing source nodes for the
   resolved drawing part, including ordered child nodes, source part paths,
   relationships, media references, raw preservation handles, and stable
   diagnostics provenance.
2. The computed view exposes the effective diagram drawing data needed by the
   adapter without depending on renderer types.
3. The core adapter maps the computed diagram drawing view to renderer
   `GroupElement` children without importing `packages/core/src/parser/*`.
4. Focused tests cover SmartArt fallback ordering, group child transforms,
   shape fills/outlines/text, media relationships, and skip diagnostics.
5. Visual output for existing SmartArt fixtures is intentionally unchanged, or
   any intentional change is covered by VRT updates in a separate rendering PR.
6. `parse-render.integration.test.ts` and parser shape-tree tests are no longer
   needed for SmartArt fallback coverage, or their role is narrowed to the
   parser oracle only.

Replacement implementation is intentionally out of scope for #528 and is split
into [#535](https://github.com/hirokisakabe/pptx-glimpse/issues/535).
