# @pptx-glimpse/document

[![npm version](https://img.shields.io/npm/v/%40pptx-glimpse%2Fdocument)](https://www.npmjs.com/package/@pptx-glimpse/document)
[![CI](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml/badge.svg)](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/hirokisakabe/pptx-glimpse/blob/main/LICENSE)

The public OOXML document foundation for reading, inspecting, authoring, editing, and writing
PowerPoint (`.pptx`) files. It owns PPTX source semantics and a non-mutating computed view; it does
not render SVG/PNG or provide editor history and UI state.

This package is published on npm and can be installed directly:

```bash
npm install @pptx-glimpse/document
```

## Quick start

```ts
import { readFile, writeFile } from "node:fs/promises";
import { createComputedView, readPptx, writePptx } from "@pptx-glimpse/document";

const source = readPptx(await readFile("input.pptx"));
const computed = createComputedView(source);

console.log(`slides: ${computed.slides.length}`);
await writeFile("round-trip.pptx", writePptx(source));
```

`PptxSourceModel` is the editable source of truth. It retains authored values, typed nodes,
relationships, source handles, and raw preservation material. `PptxComputedView` is a derived,
read-only projection that resolves effective values across the slide/layout/master/theme cascade
without mutating the source.

## Choose a document workflow

- [Read an existing PPTX](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/reading.md)
  into a typed `PptxSourceModel` while retaining unsupported material for round trips.
- [Derive effective values](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/computed-view.md)
  with `createComputedView`.
- [Author a PPTX from scratch](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/authoring.md)
  with `createPptx` and typed drawing helpers.
- [Edit an existing PPTX](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/editing.md)
  through source handles and supported immutable operations.
- [Write and preserve a PPTX](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/writing.md)
  with `writePptx`, structural round-trip preservation, and dirty-part serialization.
- Check the [feature support matrix](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/docs/feature-support.md)
  for the evidence and constraints of each workflow.

These guides are included in the published package under `docs/`.

## Public boundary and related packages

Import supported APIs only from the package root:

```ts
import {
  asOoxmlPercent,
  createComputedView,
  createPptx,
  readPptx,
  replaceTextRunPlainText,
  setPictureCrop,
  writePptx,
} from "@pptx-glimpse/document";
```

The root entry point exports the reader, source and computed-view types, branded OOXML units,
authoring/editing operations, and writer. Parser helpers, raw replacement mechanisms, dirty-scope
details, and deep `src` or `dist` paths are internal.

Use [`pptx-glimpse`](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/core/README.md)
for high-level SVG/PNG rendering or an edit/rerender/save session. Add
[`@pptx-glimpse/editor`](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/editor/README.md)
when you need validated commands, selection, and undo/redo without a UI. If your application
imports this package directly, keep it as a direct dependency even when `pptx-glimpse` is already
installed.

The dependency direction is one way: document semantics flow to upper layers. Rendering defaults,
font discovery, text measurement, pixel layout, SVG/PNG output, UI state, and history do not belong
to this package.

## Runtime support, stability, and limitations

- Node.js 22 or later is supported.
- Browser bundlers can use the byte-oriented APIs with `Uint8Array`; this package has no DOM or
  filesystem requirement. File loading and saving remain the application's responsibility.
- `@pptx-glimpse/document` is a `0.x` package. Root exports are the intended public surface, but
  minor releases may refine exported types or behavior before `1.0.0`.
- OOXML support is incremental. Unsupported material is preserved where documented, but not every
  element is available through typed read, computed-view, authoring, or editing APIs.
- Round trips preserve document structure rather than byte-for-byte ZIP or XML formatting.

See the [project README](https://github.com/hirokisakabe/pptx-glimpse#readme) for package selection
and the demo.

## License

MIT
