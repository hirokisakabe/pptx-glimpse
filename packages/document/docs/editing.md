# Editing an existing PPTX

Existing-file editing starts with `readPptx`. Locate a stable source handle, apply a supported
immutable operation, and write the returned `PptxSourceModel`:

```ts
import { readFile, writeFile } from "node:fs/promises";
import {
  readPptx,
  replaceTextRunPlainText,
  updateChartData,
  updateScatterChartData,
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

## Existing theme color and font schemes

`updateThemeScheme` applies a non-empty field-level patch to exactly one existing theme handle.
Color values use the same six-digit RGB strings as `createPptx`: `dk1`, `lt1`, `dk2`, `lt2`,
`accent1` through `accent6`, `hlink`, and `folHlink`. Major and minor font sets independently
accept `latin`, `eastAsian`, and `complexScript`; Latin must be non-empty, while the other two may
be the empty OOXML typeface.

```ts
import { updateThemeScheme, writePptx } from "@pptx-glimpse/document";

const theme = source.themes[0];
if (theme?.handle === undefined) throw new Error("No editable theme found");

const branded = updateThemeScheme(source, theme.handle, {
  colorScheme: {
    accent1: "0067C5",
    accent2: "F59E0B",
  },
  fontScheme: {
    major: { latin: "Brand Display" },
    minor: { latin: "Brand Text", eastAsian: "Noto Sans CJK JP" },
  },
});

const output = writePptx(branded);
```

Every omitted color/font field remains unchanged. The dirty writer patches only requested
`a:clrScheme` color choices and `a:majorFont` / `a:minorFont` typeface attributes, retaining the
theme part path, master relationships, `a:fmtScheme`, script fallback entries, child order, and
unrelated or unknown XML. Theme creation/deletion/duplication, format-scheme edits, relationship
rewiring, and multi-theme batch operations are not supported.

The operation rejects a missing or ambiguous handle, invalid/empty patches, invalid colors or
typefaces, missing raw package material, and missing/duplicate/unsupported target structures
before returning a changed source model. In the editor packages, the matching
`updateThemeScheme` command creates one undo entry. Core rerenders slides under every master that
references the edited shared theme and does not invalidate masters that reference other themes.

### Image replacement copy-on-write contract

`replaceImageBytes` keeps the replacement media content type equal to the source media content
type. When exactly one picture/reference uses the media part, it replaces that part's bytes in
place. When the media part has two or more known package references, it allocates the next unused
`ppt/media/imageN.<extension>` path and owner-local `rIdN`, registers the matching content type,
and patches only the selected picture's `a:blip@r:embed`. Other pictures continue to reference the
original media bytes.

The operation does not run a package-wide orphan sweep: in-place replacement creates no orphan,
and copy-on-write retains the original because other references still use it. New media paths are
reserved by the active edit journal. Editor undo restores the pre-edit snapshot (and removes that
reservation); redo restores the exact post-edit snapshot, including the allocated path and
relationship. A new edit branch after undo may therefore reuse the released allocation.

## Picture crop

`setPictureCrop` changes the DrawingML `a:srcRect` of an existing picture and
`clearPictureCrop` removes it:

```ts
import { asOoxmlPercent, clearPictureCrop, setPictureCrop } from "@pptx-glimpse/document";

const picture = source.slides[0]?.shapes.find(
  (shape) => shape.kind === "image" && shape.blipFillMode === "stretch",
);
if (picture?.handle === undefined) throw new Error("No editable picture found");

const cropped = setPictureCrop(source, picture.handle, {
  left: asOoxmlPercent(25000),
  top: asOoxmlPercent(10000),
});
const uncropped = clearPictureCrop(cropped, picture.handle);
```

Insets use integer OOXML percentage units (`100000` = 100%). Omitted edges are zero. Each edge
must be in `0..100000`, and `left + right` and `top + bottom` must each be less than `100000` so
some source area remains visible. Invalid values are rejected; they are never clamped. An
all-zero set is equivalent to clearing the crop.

The existing-edit subset is a `p:pic` on a slide whose `p:blipFill` contains exactly one
`a:stretch` and no `a:tile`. Tile fills, missing or ambiguous fill modes, layout/master pictures,
and `mc:AlternateContent` targets are rejected. The writer patches only the targeted
`a:srcRect`; sibling blip-fill attributes, effects, extension XML, shape properties, media bytes,
and unrelated package parts are preserved. Setting the same effective insets or clearing an
already absent crop returns the original source object without adding an edit record.

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

## Same-parent partial drawing moves

`moveShapes` moves one or more consecutive direct-child drawings within the same slide, layout,
master, or native group. Input handles are normalized to source document order. Pass a
direct-child anchor to insert the block immediately before it, or omit the anchor to move the
block to the end of the container's drawing order:

```ts
import { moveShapes } from "@pptx-glimpse/document";

const slide = source.slides[0];
const [first, second, third] =
  slide?.shapes.flatMap((shape) => (shape.handle === undefined ? [] : [shape.handle])) ?? [];
if (
  slide?.handle === undefined ||
  first === undefined ||
  second === undefined ||
  third === undefined
) {
  throw new Error("Three editable sibling drawings are required");
}

const beforeFirst = moveShapes(source, [third, second], slide.handle, {
  beforeShapeHandle: first,
});
const movedToEnd = moveShapes(beforeFirst, [first], slide.handle);
```

The operation preserves node ids, handles, transforms, connector endpoints, relationship ids,
package parts, non-target sibling order, and authored slots for non-drawing or unknown XML. It
rejects empty, duplicate, idless, non-consecutive, cross-part, or non-direct-child inputs; an
anchor inside the moved block; and SmartArt, raw, or `mc:AlternateContent` targets. This operation
does not move drawings across parents and never converts coordinate systems.

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
as a formatting template before their formulas and caches are rewritten. Retained `idx` / `order`
values stay unchanged so chart-level references remain valid; additions receive values after the
largest existing `idx` / `order`. This means explicitly formatted source series intentionally
provide the default appearance for added series. A cloned Microsoft `c16:uniqueId` receives a new,
non-conflicting GUID rather than copying series identity. Removing series truncates the trailing
series, so formatting and unknown XML owned by a removed series are removed with it while retained
and chart-level XML stay intact. Because explicit legend entries have chart-type-dependent index
semantics, removal rejects Charts that contain them instead of leaving a stale reference. The
worksheet range and cells expand or shrink to the resulting series count.

The same operation supports one fixed-topology category combo subset: exactly one nonempty,
column-oriented `barChart` plot group (`barDir=col`) and one nonempty `lineChart` plot group sharing
the standard category worksheet layout. Grouping must be non-stacked (`clustered` or `standard` for
the bar group and `standard` for the line group); `stacked`, `percentStacked`, and unknown grouping
semantics are rejected. Each group must author exactly two populated `axId` elements that resolve
to exactly one value axis and the same single `catAx`; extra unresolved IDs and `axId` elements
without `val` are rejected. Discover each existing series identity from computed
`chartData.series[].source`, then include that descriptor in the update input:

```ts
const comboData = createComputedView(source).slides[0]?.elements.find(
  (element) => element.kind === "chart",
)?.chartData;

if (comboData?.chartType !== "combo") throw new Error("No supported combo chart found");

const editedCombo = updateChartData(source, comboChart.handle, {
  series: comboData.series.map((series, index) => ({
    source: series.source,
    name: index === 0 ? "Actual" : "Forecast",
    categories: ["Apr", "May"],
    values: index === 0 ? [40, 55] : [42, 60],
  })),
});
```

For combo Charts every source descriptor is required and consists of the plot `chartType` plus the
existing OOXML `c:idx`. Inputs may be supplied in any order; formulas and worksheet columns remain
bound to chart XML document order. Series count, identity set, plot-group membership, group order,
`idx` / `order`, axes, title, legend, formatting, and unknown XML are fixed. Duplicate, missing, or
changed identities are rejected before source/editor history changes.

The computed view also retains `plotGroups` in chart XML order, each group's authored `grouping`,
`axisReferenceCount`, ordered populated `axisIds`, uniquely resolved category/value-axis IDs, and
`valueAxisId`, plus the referenced `valueAxes`. The rendering adapter skips unsupported grouping or
inexact/incomplete/ambiguous axis topology with an explicit diagnostic. For supported topology, the renderer uses one zero-inclusive scale for
groups that share a value-axis ID and an independent right-side scale for a secondary value axis;
zero-only and all-negative series remain renderable. A horizontal bar plus line combination is not
geometrically compatible with this category-combo renderer, so it is likewise exposed as
unsupported by the rendering adapter and rejected by editing preflight.

The operation rejects linked/external data, missing or unresolved relationships, workbooks shared
by multiple Charts, workbook formulas in the data range, and other data layouts before changing the
model. Unsupported combo shapes include three or more plot groups, scatter/bubble mixtures,
horizontal bars, stacked grouping, empty groups, missing/ambiguous/different category axes,
missing/ambiguous value axes, extra or unpopulated axis references, and anything other than one bar
plus one line group. It patches only
the target worksheet data and preserves other embedded workbook parts such as styles, themes, and
document properties.

`updateScatterChartData(source, chartHandle, input)` is the separate typed operation for an
existing scatter Chart. Its input uses paired finite `xValues` / `yValues` rather than category
labels, so category and XY data models cannot be mixed accidentally:

```ts
const editedScatter = updateScatterChartData(source, scatterChart.handle, {
  series: [
    { name: "Observed", xValues: [1, 2, 3], yValues: [2.5, 4, 3.5] },
    { name: "Forecast", xValues: [10, 20], yValues: [5, 8] },
  ],
});
```

The supported embedded worksheet uses the standard XY layout: each series has a two-column table
with X values in column A, Y values in column B, its name in the column-B header cell, and one
empty row before the next series table. Series may have different non-empty point counts, but each
series must supply the same number of X and Y values. Chart `tx`, `xVal`, and `yVal` formulas,
number caches, point counts, and worksheet cells are updated together.

Retained, appended, and removed series use the same positional identity policy as
`updateChartData`: retained series keep formatting, unknown XML, `idx`, and `order`; appended
series clone the last original series and receive fresh `idx` / `order` and Microsoft
`c16:uniqueId` values; removal truncates trailing series and rejects explicit legend entries.
Scatter style, title, legend, both value axes, series formatting, chart-level unknown XML, and
unrelated workbook/package parts remain unchanged.

The operation rejects category, bubble, combo, and other chart types, non-standard XY formulas or
worksheet cells, formula cells, external/shared workbooks, and workbooks with multiple worksheets.
Validation and new chart/workbook bytes complete before the immutable source model and editor
history are changed.

`updateBubbleChartData(source, chartHandle, input)` extends the same topology and preservation
contract to an existing bubble Chart. Each series supplies equal, non-empty arrays of finite
`xValues`, `yValues`, and `bubbleSizes`:

```ts
const editedBubble = updateBubbleChartData(source, bubbleChart.handle, {
  series: [{ name: "Observed", xValues: [1, 2], yValues: [3, 5], bubbleSizes: [8, 13] }],
});
```

The supported standard bubble worksheet layout places X in column A, Y and the series name in
column B, and bubble size under a `Size` header in column C. Series tables are separated by one
empty row. The operation synchronizes `tx`, `xVal`, `yVal`, and `bubbleSize` formulas, caches,
point counts, `idx` / `order`, and worksheet cells. It rejects scatter/category/combo Charts,
non-standard formulas or cells, formula cells, external/shared workbooks, and multiple worksheets
before changing the source model or editor history.

The text operations use the same `SourceParagraph` / `SourceTextRun` contract for ordinary shape
text and existing Table cell text. To edit a Table, locate a node through
`table.table.rows[rowIndex].cells[cellIndex].textBody?.paragraphs[paragraphIndex]`, then pass its
handle (or a child run handle) to `replaceParagraphPlainText`, `replaceTextRunPlainText`,
`setTextRunProperties`, `clearTextRunProperties`, `setParagraphProperties`, or
`clearParagraphProperties`. `replaceParagraphPlainText` replaces all runs with one run and carries
forward only the first run's properties. Table structure, merges, fills, borders, margins, and
table style are preserved by these text operations.

## Table cell fill, border, and margin

`setTableCellProperties` and `clearTableCellProperties` address a physical cell with the native
Table's `SourceHandle` plus zero-based `rowIndex` and `cellIndex` values:

```ts
import { asEmu, clearTableCellProperties, setTableCellProperties } from "@pptx-glimpse/document";

const table = source.slides[0]?.shapes.find((shape) => shape.kind === "table");
if (table?.handle === undefined) throw new Error("No editable Table found");

const address = { tableHandle: table.handle, rowIndex: 0, cellIndex: 1 };
const formatted = setTableCellProperties(source, address, {
  fill: { kind: "solid", color: { kind: "srgb", hex: "D9EAF7" } },
  borders: {
    bottom: {
      width: asEmu(12700),
      fill: { kind: "solid", color: { kind: "srgb", hex: "4472C4" } },
    },
  },
  marginLeft: asEmu(91440),
});
const inheritedFill = clearTableCellProperties(formatted, address, ["fill"]);
```

The supported fill subset is explicit `none` or solid sRGB. Each of the four physical border
sides supports a positive EMU width and explicit `none` or solid sRGB fill. Margins accept
non-negative integer EMU values. Setting an omitted field preserves it. Clearing `fill`, one of
`borderTop` / `borderBottom` / `borderLeft` / `borderRight`, or one of the four margin properties
removes only that inline `a:tcPr` value so the Table style/default can apply. This differs from
setting `{ kind: "none" }`, which writes an explicit `a:noFill` override.

The address is physical OOXML row/cell position, including merge-continuation cells; the operation
does not resolve a visual grid coordinate or modify merge topology. A missing/handleless Table,
out-of-range address, unsupported fill style, or invalid width/margin is rejected before a new
model or edit record is created. A value-identical set or clearing an already absent property
returns the original source object. Dirty writing patches only the addressed `a:tcPr`, preserving
cell text, merges, sibling cells, unknown cell XML, and unrelated package parts. Row/column
topology, merge/unmerge, diagonal borders, dash/join/cap/arrow line styles, and Table style
definition editing remain unsupported.

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
