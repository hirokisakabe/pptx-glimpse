# Editing an existing PPTX

Existing-file editing starts with `readPptx`. Locate a stable source handle, apply a supported
immutable operation, and write the returned `PptxSourceModel`:

```ts
import { readFile, writeFile } from "node:fs/promises";
import {
  readPptx,
  replaceTextRunPlainText,
  updateChartData,
  writePptx,
} from "@pptx-glimpse/document";

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
existing Chart data, slide backgrounds, and slide topology. The authoring helpers can add new
supported content to a slide loaded from an existing PPTX.

## Lossless group and ungroup

`groupShapes(source, shapeHandles)` replaces two or more consecutive siblings with one native
group while retaining each child node id, source handle identity, transform, and relative z-order.
The handles may target root-level nodes or children of the same existing group:

```ts
import { groupShapes, ungroupShape } from "@pptx-glimpse/document";

const slide = source.slides[0];
const selected = slide?.shapes
  .slice(0, 2)
  .flatMap((shape) => (shape.handle === undefined ? [] : [shape.handle]));
if (selected === undefined || selected.length !== 2) {
  throw new Error("Two editable sibling shapes are required");
}

const grouped = groupShapes(source, selected);
const group = grouped.slides[0]?.shapes.find((shape) => shape.kind === "group");
if (group?.handle === undefined) throw new Error("Group was not created");

const restored = ungroupShape(grouped, group.handle);
```

The new group uses the union of the child transform bounds and an identity child-coordinate
mapping (`off == chOff`, `ext == chExt`) so child transforms do not change. Grouping rejects
different parents, non-consecutive selections, missing/duplicate node ids,
`mc:AlternateContent`, incomplete child transforms, and connector endpoints crossing the
selection boundary. Internal connector endpoint ids remain unchanged.

`ungroupShape` expands children into the group z-order slot only when the authored group mapping
is identity and the group has no fill, effects, or unknown group-level XML whose removal could
change appearance. It also rejects a group referenced by a connector. A removed group id stays
reserved for the edit session and is not reused by a later group operation. Every rejection is
atomic because validation completes before a new source model or edit record is created.

`updateChartData(source, chartHandle, input)` replaces the series topology, names, shared category
labels, and finite numeric values of an existing supported category Chart. It keeps the Chart
type, title, legend, axes, and unknown chart-level XML, and updates the Chart formulas and caches
together with its embedded workbook:

```ts
const chart = source.slides
  .flatMap((slide) => slide.shapes)
  .find((shape) => shape.kind === "chart" && shape.handle !== undefined);

if (chart?.handle === undefined) throw new Error("No editable chart found");

const edited = updateChartData(source, chart.handle, {
  series: [
    { name: "Revenue", categories: ["Apr", "May"], values: [40, 55] },
    { name: "Cost", categories: ["Apr", "May"], values: [25, 30] },
  ],
});
```

This operation currently supports bar, line, pie, area, doughnut, and radar Charts with one
internal embedded workbook, one worksheet, and the standard tabular layout (series names in row 1,
categories in column A, values in columns B onward). Retained series are matched by position and
keep their formatting and unknown series XML. New trailing series clone the last original series
as a formatting template before their `idx`, `order`, formulas, and caches are rewritten; this
means explicitly formatted source series intentionally provide the default appearance for added
series. Removing series truncates the trailing series, so formatting and unknown XML owned by a
removed series are removed with it while retained and chart-level XML stay intact. The worksheet
range and cells expand or shrink to the resulting series count.

The operation rejects linked/external data, missing or unresolved relationships, combo Charts,
workbooks shared by multiple Charts, workbook formulas in the data range, and other data layouts
before changing the model. It patches only the target worksheet data and preserves other embedded
workbook parts such as styles, themes, and document properties.

The text operations use the same `SourceParagraph` / `SourceTextRun` contract for ordinary shape
text and existing Table cell text. To edit a Table, locate a node through
`table.table.rows[rowIndex].cells[cellIndex].textBody?.paragraphs[paragraphIndex]`, then pass its
handle (or a child run handle) to `replaceParagraphPlainText`, `replaceTextRunPlainText`,
`setTextRunProperties`, `clearTextRunProperties`, `setParagraphProperties`, or
`clearParagraphProperties`. `replaceParagraphPlainText` replaces all runs with one run and carries
forward only the first run's properties. Table structure, merges, fills, borders, margins, and
table style are preserved but are not editable through these text operations.

## Typed edits and raw preservation

The source model deliberately carries both typed nodes and preserved raw material. When a supported
edit changes the source, it updates the typed representation immediately and appends an edit record
that identifies the dirty scope. A no-op may return the source without adding a record. `writePptx`
later patches dirty scopes into preserved raw XML while copying untouched package material where
possible.

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

Only root-exported operations are public. Edit record types are exported for observing the source
model, but constructing or mutating records directly is not a supported editing operation. Raw
replacement helpers, XML locators, and writer patching modules are internal. See the
[feature support matrix](./feature-support.md) for the currently supported edit subset and
constraints.
