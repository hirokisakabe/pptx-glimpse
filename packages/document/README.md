# @pptx-glimpse/document

OOXML document foundation for reading, inspecting, authoring, editing, and writing PowerPoint
`.pptx` files.

The package owns PPTX source semantics and a non-mutating computed view. It does not render SVG or
PNG. Higher-level packages may consume `@pptx-glimpse/document`, but this lower-level package does
not depend on a renderer, editor UI, or other consumer.

## Install

```bash
npm install @pptx-glimpse/document
```

Node.js 22 or later is required.

## Quick start

```ts
import { readFile, writeFile } from "node:fs/promises";
import { createComputedView, readPptx, writePptx } from "@pptx-glimpse/document";

const source = readPptx(await readFile("input.pptx"));
const computed = createComputedView(source);

console.log(`slides: ${computed.slides.length}`);
await writeFile("round-trip.pptx", writePptx(source));
```

`PptxSourceModel` is the editable source of truth. It retains authored values, typed nodes, package
relationships, source handles, and raw preservation material. `PptxComputedView` is a derived,
read-only projection that resolves effective values across the slide/layout/master/theme cascade
without mutating the source.

## Choose a workflow

- [Read an existing PPTX](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/reading.md)
  into the typed `PptxSourceModel` while retaining unsupported material for round trips.
- [Derive effective values](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/computed-view.md)
  with `createComputedView` when you need resolved document semantics rather than authored values.
- [Author a PPTX from scratch](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/authoring.md)
  with `createPptx` and the typed drawing helpers.
- [Edit an existing PPTX](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/editing.md)
  through source handles and supported immutable editing operations.
- [Write and preserve a PPTX](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/writing.md)
  with structural round-trip preservation and dirty-part serialization.
- Check the [feature support matrix](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/feature-support.md)
  before relying on a particular reader, computed-view, authoring, editing, or preservation feature.

These documents are also included in the published package under `docs/`, so they remain available
in an installed copy.

## Public API and dependency boundary

Import supported APIs only from the package root:

```ts
import {
  addEmptySlideFromLayout,
  createComputedView,
  createPptx,
  moveSlide,
  readPptx,
  replaceTextRunPlainText,
  writePptx,
} from "@pptx-glimpse/document";
```

The root entry point exports the source and computed-view types, branded OOXML units, reader,
factory, authoring/editing operations, and writer that make up the supported public surface.
Parser helpers, raw replacement mechanisms, dirty-scope implementation details, and deep source or
`dist` paths are internal and may change without notice.

The dependency direction is one way: `@pptx-glimpse/document` provides lower-level document
semantics to upper layers. Rendering defaults, font discovery and fallback, text measurement,
pixel layout, and SVG/PNG output belong to those upper layers and are not document APIs.

## Stability

`@pptx-glimpse/document` is a `0.x` package. Root exports are the intended public surface, but minor
releases may refine exported types or behavior while the package remains below `1.0.0`.
