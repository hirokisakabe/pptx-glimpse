# `@pptx-glimpse/document` responsibility boundaries

- Status: RFC decision for [#449](https://github.com/hirokisakabe/pptx-glimpse/issues/449)
- Date: 2026-06-20

This note records the package boundary decision that should feed into
[#445](https://github.com/hirokisakabe/pptx-glimpse/issues/445). It assumes the
decision from
[hirokisakabe/pom#895](https://github.com/hirokisakabe/pom/issues/895): **pom
will eventually depend on `@pptx-glimpse/document`, and pptx-glimpse will not
depend on pom.**

The related source/computed layering decision is recorded in
[cleandoc-source-computed-view.md](./cleandoc-source-computed-view.md). In short,
CleanDoc is the source model owned by `@pptx-glimpse/document`, and computed
views are generated projections for renderer, AI reading, and editor consumers.

The related raw preservation and round-trip decision is recorded in
[raw-ooxml-round-trip.md](./raw-ooxml-round-trip.md). In short,
`@pptx-glimpse/document` should preserve raw OOXML for untouched or unsupported
source material, while keeping typed CleanDoc operations as the normal editing
API.

The related core dogfood migration decision is recorded in
[core-document-dogfood-migration.md](./core-document-dogfood-migration.md). In
short, public SVG/PNG conversion should eventually route through CleanDoc source
reading, computed view generation, and a core-owned adapter into the existing
renderer model.

The related minimal implementation PoC scope is recorded in
[cleandoc-minimal-poc-scope.md](./cleandoc-minimal-poc-scope.md). In short, the
first slice should prove simple themed slide read, computed render view, current
renderer adapter parity, no-edit structural round-trip, and one plain text-run
edit round-trip.

Package names with the `@pptx-glimpse/*` scope describe the intended long-term
package boundaries. Older repository revisions used transitionary names such as
`packages/pptx-glimpse` and `pptx-glimpse-renderer`; the current directory
layout uses `packages/core`, `packages/document`, and `packages/renderer` to make
those roles explicit. The public npm package name remains `pptx-glimpse`, while
the core role corresponds to the future `@pptx-glimpse/core` boundary.

## Dependency direction

`@pptx-glimpse/document` is the lower-level OOXML / CleanDoc foundation. Higher
packages may consume it, but it must not know product-level APIs, preview APIs,
editing workflows, or pom's authoring DSL.

CleanDoc means the clean intermediate document model discussed in
[#445](https://github.com/hirokisakabe/pptx-glimpse/issues/445): it keeps the
semantic structure and round-trip preservation hooks needed for PPTX documents
without exposing OOXML's package scattering as the primary authoring surface.

```text
@hirokisakabe/pom
  -> @pptx-glimpse/document

@pptx-glimpse/editor-core
  -> @pptx-glimpse/document

@pptx-glimpse/core
  -> @pptx-glimpse/document
  -> @pptx-glimpse/renderer
  (owns orchestration and the CleanDoc-to-render-view adapter)

@pptx-glimpse/document
  -> OOXML package parts

@pptx-glimpse/renderer
  -> computed render view
```

`@pptx-glimpse/document` must not depend on:

- `@pptx-glimpse/core`
- `@pptx-glimpse/editor-core`
- `@pptx-glimpse/renderer`
- `@hirokisakabe/pom`
- browser UI packages or demo application code

The dependency contract is one-way:

```text
authoring / preview / editor layers
  -> document
  -> OOXML file structure
```

There should be no reverse dependency from `document` into authoring, preview, or
editor-specific concepts.

## What `document` owns

`@pptx-glimpse/document` owns the PPTX document domain model and the mechanics
needed to read, write, validate, and preserve that domain model:

- CleanDoc schema and TypeScript types.
- Presentation package structure: slides, layouts, masters, themes, media,
  relationships, and content types.
- Reader-side conversion from OOXML package parts into CleanDoc.
- Writer-side conversion from CleanDoc back into OOXML package parts.
- Stable IDs and relationship references that allow round-trip editing.
- Normalized units and value types where CleanDoc intentionally hides OOXML
  encoding details.
- Source-level values plus the metadata required to recover or preserve OOXML
  constructs that CleanDoc does not model directly.
- Raw OOXML escape hatch for vendor extensions, uncommon DrawingML constructs,
  `mc:AlternateContent`, and other cases needed to keep round-trips lossless.
- Validation and diagnostics that are about document correctness, not rendering
  fidelity or UI behavior.

The package may expose utilities that compute effective document values when
those values are part of the document semantics, for example resolving a slide's
master/layout inheritance into a document-level view. Those utilities should not
encode renderer fallbacks, font substitution policy, browser behavior, or pom
authoring conveniences.

## What `document` does not own

`@pptx-glimpse/document` does not own higher-level workflows:

- SVG/PNG rendering APIs.
- Renderer-specific intermediate model shapes.
- Font discovery, font fallback, font subsetting, or text-to-path conversion.
- Browser editor commands, selection state, undo history, collaboration state, or
  UI interaction models.
- pom's Flexbox-like authoring DSL, layout primitives, or AI-first document
  authoring shortcuts.
- `convertPptxToPom` lossy authoring conversion policy.
- CLI, demo, VRT harnesses, or publish-time package bundling behavior.

Those concerns belong in `core`, `editor-core`, renderer packages, pom, or
application-level packages that depend on `document`.

## Package roles

| Package                     | Role                                               | Depends on `document`?                        | `document` may depend on it? |
| --------------------------- | -------------------------------------------------- | --------------------------------------------- | ---------------------------- |
| `@pptx-glimpse/document`    | CleanDoc, OOXML reader/writer, document validation | n/a                                           | n/a                          |
| `@pptx-glimpse/core`        | Public preview/conversion API and orchestration    | Yes                                           | No                           |
| `@pptx-glimpse/editor-core` | Headless editing commands and state machine        | Yes                                           | No                           |
| `@pptx-glimpse/renderer`    | Render-oriented model and SVG/PNG generation       | No direct dependency; consumes adapter output | No                           |
| `@hirokisakabe/pom`         | Authoring DSL and generation workflow              | Yes, eventually                               | No                           |

## Renderer model and CleanDoc

The existing renderer model should **not** be merged directly into CleanDoc.

The renderer model is a computed, display-oriented view optimized for SVG/PNG
generation. It can contain resolved values, fallback decisions, pixel-oriented
measurements, rendering warnings, and simplifications that are useful for
preview output but too lossy or too presentation-specific to be the canonical
document model.

CleanDoc should instead be the source/semantic model. Rendering should use an
adapter:

```text
OOXML package
  -> CleanDoc
  -> computed render view / renderer adapter
  -> SVG / PNG
```

This lets `document` preserve enough structure for writer/editor/round-trip
work, while `renderer` keeps a purpose-built shape for visual output.

Short term, the current parser-to-renderer path can continue to exist. The
migration path is:

1. Introduce CleanDoc types in `@pptx-glimpse/document`.
2. Add a reader path that builds CleanDoc from OOXML.
3. Add an adapter from CleanDoc to the current renderer model.
4. Gradually move parser semantics that are not renderer-specific into
   `document`.
5. Keep renderer-only fallbacks and visual diagnostics outside `document`.

## Conclusion

Adopt `@pptx-glimpse/document` as a lower-level foundation. `core`,
`editor-core`, and pom may depend on it; `document` must not depend on them.

The package boundary should be:

- `document`: canonical PPTX/CleanDoc data model, OOXML read/write mechanics,
  preservation hooks, validation.
- `core`: public conversion orchestration and compatibility API.
- `editor-core`: editing commands and headless editor state built on CleanDoc.
- renderer: computed display model and SVG/PNG output.
- pom: authoring DSL that may emit or consume CleanDoc but is never known by
  pptx-glimpse packages.

For #445, this means CleanDoc should be designed as the canonical
`@pptx-glimpse/document` model, with a generated computed render view rather than
an in-place replacement of the existing renderer model.
