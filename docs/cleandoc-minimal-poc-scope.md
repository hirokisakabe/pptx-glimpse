# CleanDoc minimal PoC scope

- Status: RFC decision for [#453](https://github.com/hirokisakabe/pptx-glimpse/issues/453)
- Date: 2026-06-20

This note records the first implementation slice for `@pptx-glimpse/document`.
It builds on the package boundary decision in
[document-boundaries.md](./document-boundaries.md), the source/computed layering
decision in
[cleandoc-source-computed-view.md](./cleandoc-source-computed-view.md), the raw
round-trip policy in [raw-ooxml-round-trip.md](./raw-ooxml-round-trip.md), and
the core dogfood migration plan in
[core-document-dogfood-migration.md](./core-document-dogfood-migration.md).
Issue #453 is the PoC-scope child decision for the broader CleanDoc RFC in
[#445](https://github.com/hirokisakabe/pptx-glimpse/issues/445), so this note
ends with a #445-ready conclusion.

The PoC is not a full PPTX implementation. It is the smallest slice that can
prove the intended architecture end to end:

```text
PPTX package
  -> CleanDoc source model
  -> computed document view
  -> current renderer model adapter
  -> SVG/PNG render verification

PPTX package
  -> CleanDoc source model
  -> writer
  -> structurally preserved PPTX package
```

## Decision

The first PoC should target a small, real subset of the existing conversion
domain: simple themed slides with shapes, text, images, slide/layout/master
inheritance, and raw preservation for everything else.

The PoC must prove two things at the same time:

- `@pptx-glimpse/document` can read enough source structure to feed a computed
  render view without depending on core, renderer, editor-core, or pom.
- The writer can preserve unedited package material and safely apply one narrow
  text edit without trying to become a byte-equivalent OOXML patcher.

## Included PPTX Subset

### Package and presentation structure

The PoC should read and preserve:

- `[Content_Types].xml`, package relationships, presentation relationships, and
  slide relationships.
- Presentation slide size and slide order.
- Slide parts, slide layout parts, slide master parts, theme parts, media parts,
  and their relationships.
- Existing relationship IDs and part paths for unchanged package material.
- Raw package bytes or raw XML trees for untouched or unsupported parts.

The source model should expose stable source handles for parts and parsed nodes
so diagnostics, computed views, and the writer can point back to original source
material.

### Slides, layouts, masters, and themes

The PoC should model:

- Visible slides in presentation order.
- The slide -> layout -> master -> theme chain for each slide.
- Slide size.
- Slide, layout, and master backgrounds, with computed fallback in slide ->
  layout -> master order.
- Layout and master `clrMap` / `clrMapOvr` data required for theme color
  resolution.
- Layout and master placeholders needed by supported slide shapes.
- `showMasterSp` / layout master-shape visibility needed to reproduce the
  current render path.

The source model keeps slide, layout, and master parts separate. The computed
view resolves effective rendering order and inheritance.

### Shapes

The PoC should type the following shape data:

- Non-grouped `p:sp` autoshapes with `p:nvSpPr`, `p:spPr`, and optional
  `p:txBody`.
- Stable shape identity from non-visual properties.
- Transform: offset, extent, rotation when present, and flip flags when present.
- Preset geometry for the common shapes needed by fixtures, initially `rect`,
  `roundRect`, and `ellipse`.
- Solid fills from `srgbClr`, `schemeClr`, and basic alpha/tint/shade/luminance
  transforms already needed by the current renderer.
- Outline color and width for simple solid lines.
- Placeholder type/index metadata for supported placeholder matching.

Unsupported shape children and attributes must be preserved as raw node
sidecars.

### Text

The PoC should type enough text to verify source editing and render parity:

- Text bodies, paragraphs, and runs.
- Plain run text.
- Run properties for font size, typeface tokens or explicit typefaces, bold,
  italic, underline when present, and solid run color.
- Paragraph alignment and basic indentation when present.
- Text box body properties that affect current rendering for the selected
  fixtures, including margins and vertical anchor when present.
- Text style inheritance from slide, layout, master, theme, and presentation
  defaults in the computed view where the current parser already depends on it.

The first text-edit API should edit one existing run's plain text while
preserving its run and paragraph properties.

### Images

The PoC should type:

- Embedded raster `p:pic` elements that reference package media by relationship
  ID.
- Image transform and stable source identity.
- Media part references and media bytes.
- Basic crop data if present in selected fixtures.

Unsupported image effects, recolor operations, and extension nodes are preserved
as raw sidecars.

### Render adapter target

The PoC computed render view should contain enough resolved data for core to map
it into the current renderer model for the selected subset:

- Effective slide size and background.
- Effective element ordering across master, layout, and slide content.
- Effective shape transform, geometry, fill, outline, and text body.
- Effective image transform and media payload.
- Resolved theme colors and typeface tokens, while leaving system font fallback
  to core/renderer.

The adapter belongs outside `@pptx-glimpse/document`, with core orchestration.

## Intentionally Excluded Elements

The PoC should preserve but not type or edit:

- Tables, charts, SmartArt, diagrams, embedded workbooks, OLE objects, and
  embedded documents.
- Audio, video, transitions, animations, timings, comments, and notes.
- Group shapes and nested shape trees beyond preserving their raw source.
- Freeform/custom geometry, connectors, shape adjustment handles, and complex
  geometry guides.
- Gradients, pattern fills, picture fills on shapes, theme effects, 3D effects,
  shadows, glows, reflections, soft edges, and complex line styles.
- Complex text features such as rich bullets, numbered lists, tab stops,
  fields, hyperlinks, bidirectional text, vertical writing modes, complex script
  shaping, and text autofit beyond what existing fixture parity requires.
- `mc:AlternateContent` branch selection as a typed editing feature. The source
  model preserves the full compatibility container; a computed view may select a
  branch for rendering diagnostics only.
- New slide creation, slide deletion, media replacement, layout editing, theme
  editing, and relationship graph rewrites beyond what the one-text-edit writer
  needs.
- Public raw OOXML editing APIs beyond narrow internal handles and diagnostics.

Excluded material must be retained through raw package preservation or raw node
sidecars unless an explicit supported edit invalidates the owning node.

## Read to Computed View to Render Verification

The PoC should verify the read -> CleanDoc -> computed view -> render path with
focused structural tests before relying on VRT.

Recommended checks:

1. Read a fixture PPTX into `CleanDocSource`.
2. Assert package graph basics: slide count, slide order, slide size,
   slide/layout/master/theme references, relationship IDs, and media references.
3. Assert typed source nodes for supported shapes, text runs, and images while
   confirming unsupported material is still attached as raw source material.
4. Generate a computed view for selected slides.
5. Assert computed values for background fallback, theme color resolution,
   placeholder matching, text style inheritance, `showMasterSp` visibility, and
   effective element ordering.
6. Adapt the computed view into the current renderer model in core.
7. Compare the adapter output with the current parser output for the supported
   subset using structural assertions.
8. Render the adapted output and verify SVG generation succeeds with no intended
   public output change.

The initial fixture set should prefer existing real/shared fixtures before new
ones, especially `shared-fixtures/real-basic-theme.pptx` and
`shared-fixtures/real-product-page.pptx`, then add minimal synthetic fixtures
only when a specific inheritance or preservation case is not covered.

VRT snapshot updates should not be part of the first PoC unless the public render
output intentionally changes.

## Read to CleanDoc to Write Verification

The PoC writer should prove structural preservation, not byte equality.

Recommended no-edit round-trip checks:

1. Read a fixture PPTX into `CleanDocSource`.
2. Write it without applying edits.
3. Assert the result is a valid PPTX ZIP package with required content types and
   relationship parts.
4. Assert slide count, slide size, slide order, slide/layout/master/theme
   references, relationship IDs, and media part bytes are preserved for
   unchanged material.
5. Re-read the written PPTX and assert the supported CleanDoc source subset is
   equivalent to the original read.
6. Render original and round-tripped PPTX through the current public conversion
   path and assert no meaningful output difference for the supported subset.

Recommended one-text-edit checks:

1. Read a fixture PPTX into `CleanDocSource`.
2. Locate one existing text run by stable source handle.
3. Replace only that run's plain text.
4. Write the PPTX.
5. Re-read the written PPTX and assert the edited run contains the new text while
   its paragraph/run formatting and stable surrounding structure are preserved.
6. Assert unrelated package parts, media bytes, relationship IDs, and unsupported
   raw material outside the dirty text scope are preserved.
7. Render the edited PPTX and assert the changed text is visible while unrelated
   rendered elements remain equivalent.

The writer may regenerate the dirty slide XML part in the first implementation.
It must preserve untouched parts and emit diagnostics if a supported edit forces
raw material inside the dirty scope to be dropped.

## Expected Round-trip Results

### No-edit round-trip

Expected:

- The output opens as a PPTX and can be read again by `@pptx-glimpse/document`.
- Supported source semantics are structurally equivalent after re-read.
- Untouched package parts, media bytes, content type entries, relationship IDs,
  and raw unsupported material are preserved where practical.
- Rendering through the current public path remains visually equivalent for the
  supported subset.

Not expected:

- Byte-equivalent ZIP output.
- Identical XML attribute order, namespace prefix placement, insignificant
  whitespace, compression settings, or ZIP metadata.
- Full semantic verification of unsupported features beyond preservation.

### One text edit

Expected:

- Exactly one selected run's plain text changes.
- Existing run/paragraph formatting for that run is retained.
- The edited PPTX reopens and re-reads with the new text.
- Unrelated slides, layouts, masters, themes, media, and relationships remain
  unchanged or structurally preserved.
- Unsupported content outside the dirty scope remains preserved.

Not expected:

- Arbitrary text reflow equivalence across PowerPoint, LibreOffice, and SVG
  output.
- Safe edits to partially supported rich text constructs beyond plain run text.
- Byte-level preservation of the edited XML part.

## Follow-up Implementation Issues

After this PoC scope is accepted, split implementation into issues similar to:

1. Add `@pptx-glimpse/document` workspace package skeleton, public experimental
   entry points, package-boundary tests, and build wiring.
2. Define CleanDoc source model types for package graph, presentation, slides,
   layouts, masters, themes, media, simple shapes, text, images, source handles,
   raw sidecars, and diagnostics.
3. Implement `readPptx(input)` for package graph, presentation metadata, slide
   order, slide size, relationships, content types, raw package preservation,
   and media preservation.
4. Add source reader coverage for simple autoshapes, text bodies/runs, embedded
   raster images, theme references, layout/master links, placeholders, and raw
   sidecars for unsupported nodes.
5. Implement `createComputedView(source, options)` for slide size/order,
   relationship resolution, theme color resolution, background fallback,
   placeholder matching, text style inheritance, and `showMasterSp` visibility.
6. Add a core-owned CleanDoc computed-view-to-current-renderer-model adapter for
   the supported subset.
7. Add dual-reader structural comparison tests against the current parser for
   selected shared fixtures.
8. Implement no-edit writer output that preserves untouched package material and
   can be re-read.
9. Implement one-run text edit and dirty-slide writer behavior with diagnostics
   for invalidated raw sidecars.
10. Add end-to-end no-edit and one-text-edit round-trip tests.
11. Add an internal or experimental core render path that can render selected
    fixtures through `@pptx-glimpse/document` without changing the public default.
12. Expand fixture coverage and opt selected VRT/shared fixtures into the
    document path only after structural parity is stable.

## Conclusion for #445

The conclusion to reflect in #445 is:

The first `@pptx-glimpse/document` PoC should be a narrow end-to-end slice, not a
general OOXML rewrite. It should read simple themed PPTX slides into a CleanDoc
source model, generate a computed view that resolves slide/layout/master/theme
semantics, adapt that view into the current renderer model for comparison, and
write structurally preserved PPTX output for no-edit and one-plain-text-run edit
cases.

The included subset is presentation package structure, slide order and size,
slide/layout/master/theme chains, background fallback, placeholder and theme
resolution, simple autoshapes, text runs, embedded raster images, and raw
preservation hooks. Tables, charts, SmartArt, media playback, animations,
transitions, notes, complex geometry/effects/text, group editing, and broad raw
OOXML editing APIs remain outside the first PoC.

Success is measured by structural source assertions, computed-view assertions,
adapter parity against the current parser for the supported subset, successful
SVG rendering with no intended public output change, no-edit structural
round-trip, and one text-run edit round-trip. Byte equality is explicitly not a
goal.
