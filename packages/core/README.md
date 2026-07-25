# pptx-glimpse

[![npm version](https://img.shields.io/npm/v/pptx-glimpse)](https://www.npmjs.com/package/pptx-glimpse)
[![CI](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml/badge.svg)](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/hirokisakabe/pptx-glimpse/blob/main/LICENSE)

The public, high-level PPTX render and edit toolkit for Node.js and the browser. It converts
PowerPoint (`.pptx`) bytes to SVG or PNG and provides an environment-independent editor session that
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

`createPptxEditorSession` is the shared high-level editing API exported by both the Node.js main
entry and the browser conditional entry. It accepts and returns `Uint8Array` data and does not
perform file or DOM operations.

### Node.js

```ts
import { readFile, writeFile } from "node:fs/promises";
import { createPptxEditorSession } from "pptx-glimpse";

const input = new Uint8Array(await readFile("presentation.pptx"));
const editor = await createPptxEditorSession(input, { skipSystemFonts: false });
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

### Browser

```ts
import { createPptxEditorSession } from "pptx-glimpse";

const [pptx, inter] = await Promise.all([
  fetch("/slides/presentation.pptx").then((response) => response.arrayBuffer()),
  fetch("/fonts/Inter-Regular.ttf").then((response) => response.arrayBuffer()),
]);
const editor = await createPptxEditorSession(new Uint8Array(pptx), {
  fonts: [{ name: "Inter", data: inter }],
  textOutput: "text",
});

await editor.apply(command);
document.querySelector("#preview")!.innerHTML = editor.slides[0]?.svg ?? "";

const downloadBytes = editor.save().pptx;
```

The session exposes rendered slides, editable shape information, command application, selection,
undo/redo history, and PPTX serialization. It is headless: the package does not provide a complete
editor UI.

### Errors and warnings

The high-level session throws `PptxEditorError` for expected editor rejection and read, render, or
write failures. Use `isPptxEditorError()` to narrow an unknown caught value and branch on its
machine-readable code:

```ts
import { createPptxEditorSession, isPptxEditorError } from "pptx-glimpse";

try {
  const editor = await createPptxEditorSession(input);
  await editor.undo();
} catch (error) {
  if (isPptxEditorError(error)) {
    console.error(error.code, error.message, error.cause);
    // invalid-command | invalid-selection | empty-undo-stack | empty-redo-stack
    // | read-failed | render-failed | write-failed
  } else {
    throw error;
  }
}
```

Command warnings do not throw. They remain on successful responses:

```ts
const response = await editor.apply(replaceImageCommand);
for (const warning of response.warnings ?? []) {
  if (warning.code === "shared-media-part") {
    console.warn(warning.message);
  }
}
```

Use `pptx-glimpse` when one object should coordinate PPTX reading, editor commands, SVG rerendering,
history, and saving. Use `@pptx-glimpse/editor` directly when your application already owns the
`PptxSourceModel` and rendering/writing lifecycle and only needs UI-independent commands,
selection, validation, and undo/redo.

### Runtime render options

The session accepts the SVG `ConvertOptions` except `slides`; it rerenders the complete current
document after each successful edit.

- `fonts` supplies font bytes directly and is the portable option for browsers, workers, Edge
  Runtime, and Node.js. Browser applications should normally provide it.
- `fontDirs` and OS system-font discovery use Node.js filesystem support. The session defaults
  `skipSystemFonts` to `true` in every runtime for deterministic, browser-safe behavior; pass
  `skipSystemFonts: false` in Node.js to enable system-font discovery, or combine
  `skipSystemFonts: true` with `fontDirs` to scan only explicit directories.
- `textOutput: "text"` produces selectable browser SVG text with embedded subset fonts.
  `"path"` produces font outlines.
- The editor preview is SVG-only, so `PptxEditorSession` does not initialize or use resvg WASM.
  WASM initialization with `initResvgWasm` is required only when the separate browser
  `convertPptxToPng` API is used.

For the breaking v4 rename from the former `Browser*` names, see the
[v4 migration guide](https://github.com/hirokisakabe/pptx-glimpse/blob/main/docs/migration-v4.md).

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
  conversion requires `initResvgWasm`; the SVG-only editor session does not.
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
