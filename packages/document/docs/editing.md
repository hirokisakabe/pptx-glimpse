# Editing an existing PPTX

Existing-file editing starts with `readPptx`. Locate a stable source handle, apply a supported
immutable operation, and write the returned `PptxSourceModel`:

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

The public root API also provides focused operations for supported run and paragraph properties,
paragraph text, shape transforms/fills/outlines, shape deletion, same-format image replacement,
slide backgrounds, and slide topology. The authoring helpers can add new supported content to a
slide loaded from an existing PPTX.

## Typed edits and raw preservation

The source model deliberately carries both typed nodes and preserved raw material. A supported edit
updates the typed representation immediately and appends an edit record that identifies the dirty
scope. `writePptx` later patches that scope into preserved raw XML while copying untouched package
material where possible.

This design avoids regenerating an entire input package from an incomplete typed model. It also
means some combinations are rejected instead of guessed. For example, an operation that cannot
safely merge pending edits into preserved raw material may throw at runtime. Treat such an error as
an unsupported workflow rather than mutating `source.edits` or raw sidecars yourself.

Create a fresh [computed view](./computed-view.md) after an edit when you need updated effective
values. A previously created computed view does not update itself.

## Editing is not from-scratch authoring

[From-scratch authoring](./authoring.md) creates a known package skeleton and supported typed
content. Existing editing must additionally preserve unknown input parts and source-local OOXML.
The available operations and preservation constraints therefore differ even when both workflows
eventually call the same writer.

Only root-exported operations are public. Raw replacement helpers, edit descriptors, XML locators,
and writer patching modules are internal. See the [feature support matrix](./feature-support.md) for
the currently supported edit subset and constraints.
