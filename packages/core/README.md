# pptx-glimpse

[![npm version](https://img.shields.io/npm/v/pptx-glimpse)](https://www.npmjs.com/package/pptx-glimpse)
[![CI](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml/badge.svg)](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/hirokisakabe/pptx-glimpse/blob/main/LICENSE)

The public, high-level PPTX render and edit toolkit for Node.js and the browser. It converts
PowerPoint (`.pptx`) bytes to SVG or PNG and provides a browser-oriented editor session that
combines document reading, headless commands, rerendering, history, selection, and writing.

This package is published on npm and can be installed directly:

```bash
npm install pptx-glimpse
```

## Render quick start

```ts
import { readFile, writeFile } from "node:fs/promises";
import { convertPptxToPng, convertPptxToSvg } from "pptx-glimpse";

const input = await readFile("presentation.pptx");
const svgReport = await convertPptxToSvg(input);
const pngReport = await convertPptxToPng(input);
const firstPng = pngReport.slides[0];

console.log(svgReport.slides[0]?.svg);
if (firstPng !== undefined) {
  await writeFile("slide-1.png", firstPng.png);
}
```

Both conversion functions return a report with slide results, diagnostics, and support coverage.
SVG conversion works in Node.js and browser bundles. Browser PNG conversion additionally requires
explicit `@resvg/resvg-wasm` initialization through `initResvgWasm`.

## High-level editing

`createBrowserPptxEditorSession` is the released high-level editing API. Despite its name, it is
also available from the Node.js entry point.

```ts
import { readFile, writeFile } from "node:fs/promises";
import { createBrowserPptxEditorSession } from "pptx-glimpse";

const input = await readFile("presentation.pptx");
const editor = await createBrowserPptxEditorSession(input, { skipSystemFonts: true });
const run = editor
  .shapes(1)
  .flatMap((shape) => shape.textBody?.paragraphs ?? [])
  .flatMap((paragraph) => paragraph.runs)
  .find((candidate) => candidate.handle !== undefined);

if (run?.handle === undefined) {
  throw new Error("No editable text run found");
}

await editor.apply({
  kind: "replaceTextRunPlainText",
  handle: run.handle,
  text: "Edited with pptx-glimpse",
});

await writeFile("edited.pptx", editor.save().pptx);
```

The session exposes rendered slides, editable shape information, command application, selection,
undo/redo history, and PPTX serialization. It is headless: the package does not provide a complete
editor UI.

## Choose a package

| Use case                                                                     | Package                                                                                                        |
| ---------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------- |
| SVG/PNG conversion or a high-level edit/rerender/save session                | `pptx-glimpse`                                                                                                 |
| Direct OOXML source, computed-view, authoring, editing, or writing workflows | [`@pptx-glimpse/document`](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/README.md) |
| UI-independent commands, selection, validation, and undo/redo                | [`@pptx-glimpse/editor`](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/editor/README.md)     |

Compatible document and editor packages are runtime dependencies of `pptx-glimpse`. If your code
imports either lower-level package, install it as a direct dependency instead of relying on
transitive dependency hoisting.

## Runtime support and stability

- Node.js 22 or later is supported.
- Browser bundles support SVG conversion and high-level editing from `Uint8Array` input.
- Browser font bytes should be supplied through conversion/editor render options. Browser PNG
  conversion requires `initResvgWasm`.
- `pptx-glimpse` follows stable major-version releases, but PPTX feature coverage is intentionally
  incomplete. Consult the root [feature support](https://github.com/hirokisakabe/pptx-glimpse#feature-support)
  section before relying on a specific PowerPoint feature.
- Rendering prioritizes common static slide content; it is not a pixel-identical PowerPoint
  replacement. Macros, animations, embedded audio/video, and several advanced chart/effect types
  are not supported.

See the [project README](https://github.com/hirokisakabe/pptx-glimpse#readme) for the demo, detailed
rendering options, font setup, and the full feature matrix.

## License

MIT
