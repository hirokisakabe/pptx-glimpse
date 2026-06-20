# Raw OOXML escape hatch and round-trip policy

- Status: RFC decision for [#451](https://github.com/hirokisakabe/pptx-glimpse/issues/451)
- Date: 2026-06-20

This note records the raw OOXML preservation and round-trip policy that should
feed into [#445](https://github.com/hirokisakabe/pptx-glimpse/issues/445). It
builds on the package boundary decision in
[document-boundaries.md](./document-boundaries.md) and the source/computed
layering decision in
[cleandoc-source-computed-view.md](./cleandoc-source-computed-view.md).

The related core dogfood migration decision is recorded in
[core-document-dogfood-migration.md](./core-document-dogfood-migration.md). In
short, public SVG/PNG conversion should eventually route through CleanDoc source
reading, computed view generation, and a core-owned adapter into the existing
renderer model.

The goal is to make `@pptx-glimpse/document` reliable for existing PPTX files
without turning CleanDoc into a byte-level OOXML editor. CleanDoc should expose a
clean document model for supported semantics, while retaining enough raw source
material to write back unsupported or unedited content.

## Round-trip Target

Adopt **structural round-trip preservation** as the writer target.

Structural preservation means:

- The package parts, relationships, content types, media, theme, master, layout,
  slide, chart, table, and notes parts that are not edited are kept intact where
  practical.
- Supported edits update only the affected semantic nodes and their required
  package bookkeeping.
- Unknown elements, vendor extensions, and compatibility branches remain in the
  written package unless an explicit edit invalidates their owning node.
- Output is expected to reopen in PowerPoint, LibreOffice, and other PPTX
  consumers with the same visible meaning for untouched content.

Do not target byte equality.

Byte equality is intentionally out of scope because XML serialization can change
attribute order, namespace prefix placement, insignificant whitespace, ZIP entry
metadata, compression settings, relationship ID allocation, and defaulted OOXML
values. Preserving those details would force CleanDoc to become a package patcher
first and a document model second.

Do not claim full semantic equality for edited content.

Edited nodes are regenerated according to the capabilities of
`@pptx-glimpse/document`. If a user edits a partially supported construct, the
writer may need to replace that construct with the supported representation and
emit diagnostics about discarded raw material. Semantic equality is a best-effort
goal for edited content, not a guarantee.

## Raw Preservation Granularity

Adopt a hybrid preservation model:

```text
PPTX package part
  -> parsed CleanDoc source nodes
  -> raw node sidecars for unsupported or partially supported XML
  -> raw package part fallback for untouched parts
```

### Adopted Granularities

Preserve raw data at the following levels:

| Level              | Use case                                              | Policy                                                                                 |
| ------------------ | ----------------------------------------------------- | -------------------------------------------------------------------------------------- |
| Package part       | Untouched slide/layout/master/theme/chart/media parts | Keep the original bytes or original XML tree and write them back unchanged.            |
| XML node sidecar   | Unknown child elements, vendor extensions, `extLst`   | Attach raw XML to the nearest CleanDoc source node with ordering metadata.             |
| Alternate branch   | `mc:AlternateContent`                                 | Preserve all choices/fallback branches unless the owning node is regenerated.          |
| Relationship entry | `_rels/*.rels` references                             | Preserve existing relationship IDs and targets where the referenced part is unchanged. |
| Content type entry | `[Content_Types].xml` overrides/defaults              | Preserve entries for existing parts; add/remove entries only for package edits.        |
| Binary asset       | Images, media, embedded workbooks                     | Preserve bytes and part paths unless the asset is replaced or removed.                 |

CleanDoc source nodes should carry stable source handles, for example original
part path, relationship ID, element ID, child ordering slots, and raw sidecar
references. These handles let the writer decide whether it can splice a generated
node into an existing part or must regenerate a larger scope.

### Rejected Granularities

Reject byte-range patching as the primary model.

Byte-range patching would preserve formatting and attribute ordering, but it is
too fragile for XML namespace changes, relationship rewrites, and edits that move
nodes across parts. It can be considered later as an optimization for specific
safe edits, not as the foundation.

Reject element-property raw storage as the default.

Storing raw XML beside every property would make the public model noisy and would
still not solve child ordering, namespace declarations, `mc:AlternateContent`,
or relationship updates. Properties should become typed CleanDoc fields when they
are supported; unsupported material should stay as raw node sidecars.

Reject a single opaque raw XML string per slide as the only escape hatch.

That approach preserves untouched slides, but it cannot support precise edits or
diagnostics. It also prevents editor-core from knowing which unsupported content
is attached to which shape, table, chart, or text node.

Reject full normalization into CleanDoc without raw sidecars.

That would produce the cleanest model, but it would drop unsupported OOXML and
make existing-file editing unsafe.

## Write Policy

Use an edit-tracking writer that prefers preserving unedited source material and
regenerating only dirty scopes.

Recommended dirty-scope levels:

```text
field edit
  -> owning XML element dirty
  -> owning XML part dirty
  -> package manifest / relationship dirty
```

Policy:

- If a package part is untouched, write the original part back as-is.
- If a part is dirty but contains mostly supported edits, regenerate only the
  dirty XML elements and splice preserved raw sidecars back into their original
  order.
- If node-level splicing is unsafe, regenerate the whole XML part from CleanDoc
  source plus preserved raw sidecars.
- If an edit invalidates an unsupported node that cannot be preserved, drop that
  raw node only with an explicit diagnostic.
- If package topology changes, update the affected `.rels` files and
  `[Content_Types].xml` entries while preserving unrelated entries.

Do not make whole-package regeneration the default. Whole-package regeneration is
acceptable for new documents and for operations that intentionally rewrite the
document topology, but it is too destructive for existing-file round-trips.

Do not require node-level splicing for the first writer implementation. A
part-level dirty writer with raw sidecars is an acceptable first slice as long as
the API and data model leave room for more precise node-level preservation later.

## Package Bookkeeping

`@pptx-glimpse/document` should treat package bookkeeping as source-model data,
not renderer data.

Preserve:

- `_rels` files, relationship IDs, relationship types, targets, and target modes.
- `[Content_Types].xml` defaults and overrides for existing parts.
- Media assets and embedded object bytes.
- Theme, slide master, slide layout, notes master, and notes slide parts.
- Chart parts, embedded workbook parts, chart style/color parts, and their
  relationship graphs.
- Unknown package parts and their relationships when they remain reachable.

When adding or removing parts, update only the affected package graph:

- Allocate new part names deterministically and avoid collisions.
- Keep existing `rId` values where possible.
- Add required content type overrides for new part types.
- Remove orphaned parts only when the edit explicitly deletes the owning object,
  or when a cleanup option requests pruning.

## Unsupported Elements and Extensions

Preserve unsupported XML by default.

This includes:

- `mc:AlternateContent`, including all `mc:Choice` and `mc:Fallback` branches.
- Vendor extension elements such as `p:extLst`, `a:extLst`, and vendor-specific
  extension namespaces.
- Unknown DrawingML, PresentationML, chart, table, media, timing, transition, and
  animation nodes.
- Unknown attributes on otherwise supported elements.

`mc:AlternateContent` should be represented in the source model as a compatibility
container, not eagerly flattened. The computed view may choose the best branch
for rendering or AI reading, but the source model must keep the complete
alternate content block for writing.

When an edit targets a node that owns unsupported material:

- Preserve unsupported children and attributes when they do not conflict with the
  edit.
- Mark unsupported material as invalidated when the edit changes the semantics of
  the owning node in a way the writer cannot reconcile.
- Emit diagnostics that identify the source part, element, and reason raw content
  was dropped or ignored.

## Raw Escape Hatch API

Keep the raw OOXML escape hatch primarily as an internal source-model mechanism
for the first implementation.

Public APIs should expose raw material cautiously through explicit expert-level
handles, not as the default editing surface. The default CleanDoc API should stay
semantic and typed.

Recommended public shape:

```text
readPptx(input, { preserveRaw: true }) -> CleanDocSource

source.getRaw(handle) -> RawOoxmlNode | RawPackagePart | undefined
source.replaceRaw(handle, raw, options) -> CleanDocSource
source.listDiagnostics() -> Diagnostic[]
```

API principles:

- Raw handles are stable source references, not direct mutable object pointers.
- Raw replacement is opt-in and should require namespace-aware XML validation.
- Raw access should be available for advanced tools, migrations, and emergency
  preservation workflows, but ordinary editor commands should use typed CleanDoc
  operations.
- Computed views may expose provenance back to raw handles for diagnostics, but
  they should not expose raw XML as their primary data.

Defer a broad public raw editing API until writer behavior and diagnostics are
proven. Exposing raw XML too early would make it difficult to evolve the source
schema and could encourage consumers to depend on OOXML internals that
`@pptx-glimpse/document` is meant to hide.

## Conclusion for #445

For #445, CleanDoc should define lossless behavior as structural preservation,
not byte equality. `@pptx-glimpse/document` should preserve raw OOXML at package
part and XML-node sidecar granularity, keep relationships/content types/assets as
source-model data, and use edit tracking so unedited parts can be written back
without regeneration.

The initial writer may regenerate dirty parts, but it must preserve untouched
parts and attach raw sidecars in a way that allows future node-level splicing.
Unsupported elements, vendor extensions, and `mc:AlternateContent` are preserved
by default and are only dropped when an explicit edit invalidates them, with a
diagnostic.

The raw escape hatch should remain mostly internal at first, with a narrow
expert-level public API based on stable raw handles. CleanDoc remains the normal
semantic editing API; raw OOXML is a preservation and emergency escape mechanism,
not the primary programming model.
