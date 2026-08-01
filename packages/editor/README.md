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

if (run === undefined) {
  throw new Error("No editable text run found");
}

const session = createEditorSession(source);
const result = session.replaceTextRunPlainText(run, "Edited");

if (!result.ok) {
  throw new Error(`${result.code}: ${result.message}`, { cause: result.cause });
}

for (const warning of result.warnings ?? []) {
  console.warn(warning.code, warning.message);
}

await writeFile("edited.pptx", writePptx(session.document));
```

## Source-node methods, commands, history, and selection

For ordinary edits, pass source nodes directly to the corresponding `EditorSession` convenience
method. Text-run and paragraph methods replace plain text or set/clear properties. Shape methods
move, resize, transform, style, or delete a `SourceShapeNode`; `replaceImage()`,
`setPictureCrop()`, and `clearPictureCrop()` accept a `SourceImage`; `updateChartData()` accepts a
`SourceChart`; slide topology methods accept a
`SourceSlide`. `addTextBox()` and `addConnector()` also accept the target `SourceSlide`, so
application code does not need to extract source handles.
Nodes captured before earlier edits remain usable because the session resolves their stable handle
against its current document.

Text methods and commands accept paragraphs and runs from ordinary shape text or existing Table
cells with the same validation and history behavior. They edit only cell text; Table rows,
columns, merges, fills, borders, margins, and styles are outside the editor command surface.

Every convenience method creates and applies the matching command through the same validation,
warning, selection-reconciliation, and undo/redo-history path. A successful convenience-method
call that changes the document creates one undo-history entry.

Use `session.apply(command)` when commands need to be stored, logged, transported by UI code, or
constructed independently of a source-node object. Use `session.applyAll(commands)` to apply
multiple operations atomically as one undo-history entry; convenience methods are single-operation
calls and cannot be grouped into one history entry by calling them sequentially. Failed validation
returns an `invalid-command` result without changing the document.

The released command set includes:

- Text: `replaceTextRunPlainText`, `replaceParagraphPlainText`, `setTextRunProperties`,
  `clearTextRunProperties`, `setParagraphProperties`, and `clearParagraphProperties`.
- Shapes: `moveShape`, `resizeShape`, `setShapeTransform`, `setShapeFill`, `setShapeOutline`,
  `addTextBox`, `addConnector`, `deleteShape`, `groupShapes`, and `ungroupShape`. Grouping accepts
  two or more consecutive siblings with the same direct parent; the typed source-node convenience
  method covers shape, picture, connector, Table/Chart graphic frames, and native group nodes.
  Ungrouping is limited to non-empty groups whose identity child mapping, appearance, unknown XML,
  and connector references can be expanded losslessly; see the document package's
  [group constraints](../document/docs/editing.md#lossless-group-and-ungroup).
- Media: `replaceImage`, `setPictureCrop`, and `clearPictureCrop`. Crop uses integer OOXML
  percentages (`100000` = 100%), rejects invalid or fully consumed axes without clamping, and is
  limited to pictures with exactly one stretch fill and no tile fill.
- Charts: `updateChartData` for series addition/removal, names, shared category labels, and numeric
  values of a supported existing category Chart with an internal editable workbook.
- Slides: `addEmptySlideFromLayout`, `duplicateSlide`, `moveSlide`, and `deleteSlide`.

Use `undo()`, `redo()`, `canUndo`, `canRedo`, `undoDepth`, and `redoDepth` to integrate history.
`selectShape(handle)` and `deselectShape()` manage a single shape selection. Selection is reconciled
after commands and history changes, and is cleared if its shape no longer exists. Grouping selects
the new group. Ungrouping selects the first expanded child in document order. Undo/redo restore the
selection snapshot paired with each group topology transition.

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

Warnings describe successful operations and remain on the success result. Shared image replacement
does not warn because the document layer isolates the selected picture with copy-on-write.

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
- Replacing a shared media part uses copy-on-write, so only the selected picture changes. The
  exported `shared-media-part` warning variant remains for source compatibility but is not emitted
  by the current replacement operation.
- Expected operation rejection uses the exported `EditorOperationErrorCode` and
  `EditorOperationFailure` contract.

See the [project README](https://github.com/hirokisakabe/pptx-glimpse#readme) for package selection
and the demo.

## License

MIT
