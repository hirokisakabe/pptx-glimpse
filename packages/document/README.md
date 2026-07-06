# @pptx-glimpse/document

OOXML document foundation for pptx-glimpse.

This package reads PowerPoint `.pptx` files into a `PptxSourceModel`, derives a non-mutating computed view, and writes the source model back to PPTX bytes. It does not render SVG or PNG; use `pptx-glimpse` for rendering.

## Install

```bash
npm install @pptx-glimpse/document
```

Node.js 22 or later is required.

## Public API

Import supported APIs from the package root:

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

The stable entry point includes:

- `readPptx(input)` for reading PPTX bytes into a `PptxSourceModel`
- `createPptx(options?)` for creating a minimal from-scratch `PptxSourceModel`
- `createComputedView(source, options?)` for deriving slide/layout/master/theme effective values without mutating the source
- `writePptx(source)` for structural round-trip writing
- Text editing helpers such as `replaceTextRunPlainText(source, handle, text)` and related source handle lookup types exported from the root entry point
- Slide topology helpers such as `addEmptySlideFromLayout(source, { layoutPartPath })`, `duplicateSlide(source, slideHandle)`, `moveSlide(source, slideHandle, { toIndex })`, and `deleteSlide(source, slideHandle)`
- Source model, computed view, and unit types needed to consume those APIs

Parser helpers, raw replacement internals, writer dirty-scope implementation details, and OOXML implementation modules outside the package root are internal and may change without notice.

## Read and Write Round Trip

```ts
import { readFile, writeFile } from "node:fs/promises";
import { createComputedView, readPptx, writePptx } from "@pptx-glimpse/document";

const input = await readFile("input.pptx");
const source = readPptx(input);
const computed = createComputedView(source);

console.log(`slides: ${computed.slides.length}`);

const output = writePptx(source);
await writeFile("round-trip.pptx", output);
```

`writePptx` targets structural round-trip preservation rather than byte-for-byte equality. Unedited package material is preserved where possible, and supported edits mark dirty PPTX parts for serialization.

## From-Scratch PPTX

```ts
import { writeFile } from "node:fs/promises";
import { addTextBox, asEmu, createPptx, writePptx } from "@pptx-glimpse/document";

const source = createPptx();
const firstSlide = source.slides[0];

if (firstSlide?.handle === undefined) {
  throw new Error("No slide was created");
}

const edited = addTextBox(source, firstSlide.handle, {
  offsetX: asEmu(914400),
  offsetY: asEmu(914400),
  width: asEmu(3657600),
  height: asEmu(914400),
  text: "Hello from pptx-glimpse",
});

await writeFile("from-scratch.pptx", writePptx(edited));
```

## Text-Run Edit

```ts
import { readFile, writeFile } from "node:fs/promises";
import { readPptx, replaceTextRunPlainText, writePptx } from "@pptx-glimpse/document";

const source = readPptx(await readFile("input.pptx"));
const firstTextRun = source.slides
  .flatMap((slide) => slide.shapes)
  .flatMap((shape) =>
    shape.kind === "shape" && shape.textBody !== undefined
      ? shape.textBody.paragraphs.flatMap((paragraph) => paragraph.runs)
      : [],
  )
  .find((run) => run.handle !== undefined);

if (firstTextRun?.handle === undefined) {
  throw new Error("No editable text run found");
}

const edited = replaceTextRunPlainText(source, firstTextRun.handle, "Edited text");
await writeFile("edited.pptx", writePptx(edited));
```

## Stability

`@pptx-glimpse/document` starts as a `0.x` package. APIs exported from the package root are the intended public surface, but the document model is still evolving. Minor releases may refine exported types or behavior while the package remains below `1.0.0`.

Do not import from deep internal paths such as `@pptx-glimpse/document/dist/...` or source module paths. Those modules are implementation details.
