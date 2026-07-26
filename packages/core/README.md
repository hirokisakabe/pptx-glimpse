# pptx-glimpse

[![npm version](https://img.shields.io/npm/v/pptx-glimpse)](https://www.npmjs.com/package/pptx-glimpse)
[![CI](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml/badge.svg)](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/hirokisakabe/pptx-glimpse/blob/main/LICENSE)

The high-level PPTX rendering and editing toolkit for Node.js and the browser. It converts
PowerPoint (`.pptx`) bytes to SVG or PNG and provides an integrated editor session for applying
commands, rerendering slides, maintaining history, and saving updated PPTX bytes.

```bash
npm install pptx-glimpse
```

## Render a presentation

```ts
import { readFile, writeFile } from "node:fs/promises";
import { convertPptxToPng } from "pptx-glimpse";

const input = await readFile("presentation.pptx");
const report = await convertPptxToPng(input);
const firstSlide = report.slides[0];

if (firstSlide !== undefined) {
  await writeFile("slide-1.png", firstSlide.png);
}

console.log(report.diagnostics);
```

Use
[`convertPptxToSvg`](https://glimpse.pptx.app/docs/api/node/functions/convertPptxToSvg) for
embeddable SVG output. Both conversion functions return slide results, diagnostics, and support
coverage. See the
[`convertPptxToPng`](https://glimpse.pptx.app/docs/api/node/functions/convertPptxToPng) and
[`ConvertOptions`](https://glimpse.pptx.app/docs/api/node/interfaces/ConvertOptions) reference,
and
[Rendering presentations](https://glimpse.pptx.app/docs/rendering) for slide selection, output
options, and report handling.

## Edit a presentation

`createPptxEditorSession` accepts and returns bytes and does not perform file or DOM operations.
The same high-level API is available from the Node.js entry point and browser conditional entry.

```ts
import { readFile, writeFile } from "node:fs/promises";
import { createPptxEditorSession } from "pptx-glimpse";

const input = new Uint8Array(await readFile("presentation.pptx"));
const editor = await createPptxEditorSession(input, {
  skipSystemFonts: false,
});

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

The session exposes rendered SVG slides, editable shape information, command application,
selection, undo/redo history, and PPTX serialization. It is headless and does not provide a
complete editor UI.

Browser applications pass PPTX and font bytes as `Uint8Array` data:

```ts
const editor = await createPptxEditorSession(new Uint8Array(pptx), {
  fonts: [{ name: "Inter", data: inter }],
  textOutput: "text",
});
```

See [Editing presentations](https://glimpse.pptx.app/docs/editing) for commands and workflow
details. Complete signatures are available for
[`createPptxEditorSession`](https://glimpse.pptx.app/docs/api/node/functions/createPptxEditorSession),
[`PptxEditorSession`](https://glimpse.pptx.app/docs/api/node/classes/PptxEditorSession), and every
[`EditorCommand`](https://glimpse.pptx.app/docs/api/node/type-aliases/EditorCommand) payload.

## Runtime notes

- Node.js 22 or later is supported.
- Browser bundles support SVG conversion and high-level editing from `Uint8Array` input.
- Pass `fonts` when the application owns font bytes. Node.js applications can instead use
  `fontDirs` or opt into system-font discovery.
- Browser PNG conversion requires explicit `initResvgWasm` initialization. SVG conversion and the
  editor session do not.
- Expected editor failures throw `PptxEditorError`; successful commands can still return warnings.

Read the detailed guides for
[font loading and mapping](https://glimpse.pptx.app/docs/fonts),
[browser usage](https://glimpse.pptx.app/docs/browser),
[Node.js usage](https://glimpse.pptx.app/docs/nodejs), and the
[API Reference](https://glimpse.pptx.app/docs/api).

## Choose a package

| Use case                                                                     | Package                                                                                                        |
| ---------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------- |
| SVG/PNG conversion or a high-level edit/rerender/save session                | `pptx-glimpse`                                                                                                 |
| Direct OOXML source, computed-view, authoring, editing, or writing workflows | [`@pptx-glimpse/document`](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/README.md) |
| UI-independent commands, selection, validation, and undo/redo                | [`@pptx-glimpse/editor`](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/editor/README.md)     |

Compatible document and editor packages are runtime dependencies of `pptx-glimpse`. Install a
lower-level package directly when your application imports it. The
[package guide](https://glimpse.pptx.app/docs/packages) explains their boundaries and dependency
direction.

## Feature coverage and stability

`pptx-glimpse` follows stable major-version releases, but PowerPoint feature coverage is
incremental. Rendering prioritizes common static slide content and is not a pixel-identical
PowerPoint replacement. Review [Feature support](https://glimpse.pptx.app/docs/feature-support)
and the diagnostics returned for the presentations your application handles.

For the breaking v4 rename from the former `Browser*` APIs, see the
[v4 migration guide](https://github.com/hirokisakabe/pptx-glimpse/blob/main/docs/migration-v4.md).

## Project

[Documentation](https://glimpse.pptx.app/docs) ·
[Demo](https://glimpse.pptx.app/) ·
[GitHub](https://github.com/hirokisakabe/pptx-glimpse)

## License

MIT
