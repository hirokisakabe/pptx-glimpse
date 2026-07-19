# Computing effective document values

Use `createComputedView(source, options?)` to derive effective PPTX document semantics without
mutating the `PptxSourceModel`:

```ts
import { readFile } from "node:fs/promises";
import { createComputedView, readPptx } from "@pptx-glimpse/document";

const source = readPptx(await readFile("input.pptx"));
const computed = createComputedView(source, { slides: [1] });

for (const slide of computed.slides) {
  console.log(slide.slideNumber, slide.background, slide.elements);
}
```

`slides` contains 1-based slide numbers. Omit it to compute every slide in presentation order.
`applyMasterVisibility` controls `showMasterSp` handling and defaults to `true`.

## Responsibility of the computed view

`PptxComputedView` resolves document-level effective values such as:

- presentation order and slide size;
- the slide/layout/master/theme chain;
- internal relationship targets and embedded media;
- background fallback, color maps, and theme color schemes;
- placeholder matches and supported text-style cascades;
- master-shape visibility and computed element ordering.

Computed elements retain source layer, part path, and source node provenance. This lets consumers
relate effective values and diagnostics back to the source.

The computed view is a read-only projection, not an editable source of truth. Apply edits to the
original `PptxSourceModel`, then create a new computed view from the returned source model.

## What it does not compute

The document package does not add rendering-contract defaults or pixel-output decisions. Font
discovery and substitution, text measurement and wrapping, text-to-path conversion, pixel layout,
SVG/PNG output, and rendering warning policy belong to upper layers. Raw elements, fills, and
backgrounds remain available for preservation and diagnostics rather than being assigned a
rendering policy here.

Import `createComputedView`, `PptxComputedView`, and computed types from the package root. Computed
implementation modules are internal.
