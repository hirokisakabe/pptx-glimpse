# @pptx-glimpse/editor

[![npm version](https://img.shields.io/npm/v/%40pptx-glimpse%2Feditor)](https://www.npmjs.com/package/@pptx-glimpse/editor)
[![CI](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml/badge.svg)](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/hirokisakabe/pptx-glimpse/blob/main/LICENSE)

The public, headless editing layer for `PptxSourceModel`. It applies validated commands and manages
selection and undo/redo history without depending on a DOM, UI framework, or renderer.

This package is published on npm. Install it together with its document-layer dependency so both
packages are direct dependencies of your application:

```bash
npm install @pptx-glimpse/document @pptx-glimpse/editor
```

## Quick start

```ts
import { readFile, writeFile } from "node:fs/promises";
import { readPptx, writePptx } from "@pptx-glimpse/document";
import { createEditorSession } from "@pptx-glimpse/editor";

const source = readPptx(await readFile("input.pptx"));
const run = source.slides[0]?.shapes.find((shape) => shape.kind === "shape")?.textBody
  ?.paragraphs[0]?.runs[0];

if (run?.handle === undefined) {
  throw new Error("No editable text run found");
}

const session = createEditorSession(source);
const result = session.apply({
  kind: "replaceTextRunPlainText",
  handle: run.handle,
  text: "Edited",
});

if (!result.ok) {
  throw new Error(`${result.code}: ${result.message}`, { cause: result.cause });
}

for (const warning of result.warnings ?? []) {
  console.warn(warning.code, warning.message);
}

await writeFile("edited.pptx", writePptx(session.document));
```

## Commands, history, and selection

`session.apply(command)` validates and applies one command. `session.applyAll(commands)` applies a
batch as one undo-history entry. Failed validation returns an `invalid-command` result without
changing the document.

The released command set includes:

- Text: `replaceTextRunPlainText`, `replaceParagraphPlainText`, `setTextRunProperties`,
  `clearTextRunProperties`, `setParagraphProperties`, and `clearParagraphProperties`.
- Shapes: `moveShape`, `resizeShape`, `setShapeTransform`, `setShapeFill`, `setShapeOutline`,
  `addTextBox`, `addConnector`, and `deleteShape`.
- Media: `replaceImage`.
- Slides: `addEmptySlideFromLayout`, `duplicateSlide`, `moveSlide`, and `deleteSlide`.

Use `undo()`, `redo()`, `canUndo`, `canRedo`, `undoDepth`, and `redoDepth` to integrate history.
`selectShape(handle)` and `deselectShape()` manage a single shape selection. Selection is reconciled
after commands and history changes, and is cleared if its shape no longer exists.

## Operation failures and warnings

Expected failures do not throw. `apply()` / `applyAll()`, `selectShape()`, `undo()`, and `redo()`
return a discriminated result with the shared `EditorOperationFailure` shape:

```ts
const result = session.undo();
if (!result.ok) {
  switch (result.code) {
    case "empty-undo-stack":
      // Disable or ignore the undo action.
      break;
    default:
      console.error(result.message, result.cause);
  }
}
```

The operation codes are `invalid-command`, `invalid-selection`, `empty-undo-stack`, and
`empty-redo-stack`. Unexpected programmer errors and invariant violations still throw.

Warnings describe successful operations and remain on the success result:

```ts
const result = session.apply(replaceImageCommand);
if (result.ok) {
  for (const warning of result.warnings ?? []) {
    if (warning.code === "shared-media-part") {
      console.warn(warning.message);
    }
  }
}
```

## Package boundary

Use [`pptx-glimpse`](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/core/README.md)
when you want high-level rendering plus an edit/rerender/save session. Use
[`@pptx-glimpse/document`](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/README.md)
directly for source/computed document semantics, authoring, immutable editing operations, and
writing without session state.

`@pptx-glimpse/editor` adds command validation, selection, warnings, and history to the document
model. It does not include React, ProseMirror, a DOM UI, SVG/PNG rendering, or file I/O. Import
supported APIs from the package root; deep source and `dist` paths are internal.

## Runtime support, stability, and limitations

- Node.js 22 or later is supported.
- Browser bundlers can use the same byte- and model-oriented APIs; the package has no DOM or
  Node.js filesystem dependency.
- `@pptx-glimpse/editor` is a `0.x` package. Commands and result types may change before `1.0.0`.
- Elements without source handles, and edits the document layer cannot safely preserve, are
  rejected as `invalid-command`.
- New table, chart, and SmartArt editing commands are not available.
- Replacing a shared media part can change multiple image references and returns a
  `shared-media-part` warning.
- Expected operation rejection uses the exported `EditorOperationErrorCode` and
  `EditorOperationFailure` contract.

See the [project README](https://github.com/hirokisakabe/pptx-glimpse#readme) for package selection
and the demo.

## License

MIT
