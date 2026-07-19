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
- From-scratch authoring helpers such as `addTextBox(source, targetHandle, input)` and `addPicture(source, targetHandle, input)`; slide, layout, and master handles are supported
- `createPptxAuthoringSession(source)` for applying consecutive authoring operations through slide/layout/master target scopes while retaining the latest source and returning new drawing/slide handles
- `reorderShapes(source, targetHandle, orderedShapeHandles)` (or `target.reorderShapes(...)` in a session) for setting the complete shape-tree drawing order after authoring
- Master/layout authoring through `createPptx({ slideMaster, slideLayout })` and `addSlideNumber(source, masterOrLayoutHandle, input)`
- Slide topology helpers such as `addEmptySlideFromLayout(source, { layoutPartPath })`, `duplicateSlide(source, slideHandle)`, `moveSlide(source, slideHandle, { toIndex })`, and `deleteSlide(source, slideHandle)`
- Source model, computed view, and unit types needed to consume those APIs

Parser helpers, raw replacement internals, writer dirty-scope implementation details, and OOXML implementation modules outside the package root are internal and may change without notice.

For a stage-by-stage view of reader, computed view, from-scratch writer, existing edit, and
round-trip preservation support, see the [Document Feature Support matrix](docs/feature-support.md).

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

For consecutive operations, a target-scoped session retains the latest immutable source and
returns each newly created handle without requiring collection searches:

```ts
import { asEmu, createPptx, createPptxAuthoringSession } from "@pptx-glimpse/document";

const session = createPptxAuthoringSession(createPptx());
const slide = session.source.slides[0];
if (slide?.handle === undefined) throw new Error("Missing slide");

const target = session.target(slide.handle);
const shapeHandle = target.addShape({
  geometry: { kind: "preset", preset: "rect" },
  offsetX: asEmu(0),
  offsetY: asEmu(0),
  width: asEmu(914400),
  height: asEmu(914400),
});
const connectorHandle = target.addConnector({
  preset: "straightConnector1",
  offsetX: asEmu(914400),
  offsetY: asEmu(457200),
  width: asEmu(914400),
  height: asEmu(1),
  start: { shapeHandle, connectionSiteIndex: 0 },
});
target.reorderShapes([connectorHandle, shapeHandle]);

const authored = session.source;
```

`reorderShapes` requires every top-level drawing in the target exactly once. Because the reorder
operation is applied after earlier additions, a connector can be placed behind its connection
targets without weakening endpoint validation.

### Alpha colors and gradients

Color transforms and gradient coordinates use OOXML percentages: `0` is 0% and `100000`
is 100%. Gradients use `gradientType` as their discriminator. Radial `centerX` and
`centerY` locate the center within the shape bounds.

```ts
import { addShape, asEmu, asOoxmlPercent, createPptx } from "@pptx-glimpse/document";

let source = createPptx();
declare const imageBytes: Uint8Array;
const slide = source.slides[0];
if (slide?.handle === undefined) throw new Error("Missing slide");

source = addShape(source, slide.handle, {
  geometry: { kind: "preset", preset: "ellipse" },
  offsetX: asEmu(914400),
  offsetY: asEmu(914400),
  width: asEmu(2743200),
  height: asEmu(1828800),
  fill: {
    kind: "gradient",
    gradientType: "radial",
    centerX: asOoxmlPercent(50000),
    centerY: asOoxmlPercent(50000),
    stops: [
      {
        position: asOoxmlPercent(0),
        color: {
          kind: "srgb",
          hex: "FF0000",
          transforms: [{ kind: "alpha", value: asOoxmlPercent(75000) }],
        },
      },
      {
        position: asOoxmlPercent(100000),
        color: { kind: "srgb", hex: "0000FF" },
      },
    ],
  },
  outline: {
    fill: {
      kind: "gradient",
      gradientType: "linear",
      stops: [
        { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FFFFFF" } },
        { position: asOoxmlPercent(100000), color: { kind: "srgb", hex: "000000" } },
      ],
    },
  },
});
```

### Shape and picture shadows

Shape effects accept glow, outer shadow, and inner shadow in any supported combination. Picture
effects accept outer and inner shadows. Shadow radius and distance are EMU integers from `0` to
`2147483647`; direction is an OOXML angle where one degree is 60,000. Outer-shadow alignment is one
of `tl`, `t`, `tr`, `l`, `ctr`, `r`, `bl`, `b`, or `br`.

```ts
import {
  addPicture,
  addShape,
  asEmu,
  asOoxmlAngle,
  asOoxmlPercent,
  createPptx,
} from "@pptx-glimpse/document";

let source = createPptx();
const slide = source.slides[0];
if (slide?.handle === undefined) throw new Error("Missing slide");

source = addShape(source, slide.handle, {
  geometry: { kind: "preset", preset: "rect" },
  offsetX: asEmu(914400),
  offsetY: asEmu(914400),
  width: asEmu(2743200),
  height: asEmu(1828800),
  effects: {
    outerShadow: {
      blurRadius: asEmu(40000),
      distance: asEmu(20000),
      direction: asOoxmlAngle(45 * 60000),
      alignment: "br",
      rotateWithShape: false,
      color: {
        kind: "srgb",
        hex: "000000",
        transforms: [{ kind: "alpha", value: asOoxmlPercent(40000) }],
      },
    },
  },
});

source = addPicture(source, slide.handle, {
  bytes: imageBytes,
  offsetX: asEmu(4114800),
  offsetY: asEmu(914400),
  width: asEmu(1828800),
  height: asEmu(1828800),
  effects: {
    innerShadow: {
      blurRadius: asEmu(30000),
      distance: asEmu(15000),
      direction: asOoxmlAngle(135 * 60000),
      color: { kind: "srgb", hex: "334155" },
    },
  },
});
```

### Slide master and layout authoring

`createPptx` can name and configure its initial master/layout. The same text, shape,
connector, and picture authoring helpers used for slides accept the generated master or
layout handle. A layout margin is materialized when a text-bearing shape is subsequently
authored directly on a slide that references that layout; it does not rewrite existing or
inherited shapes, and explicit per-shape margin values take precedence.

```ts
import { writeFile } from "node:fs/promises";

import {
  addEmptySlideFromLayout,
  addSlideNumber,
  addTextBox,
  asEmu,
  createPptx,
  writePptx,
} from "@pptx-glimpse/document";

let source = createPptx({
  slideMaster: {
    name: "Product Master",
    background: { kind: "solid", color: { kind: "srgb", hex: "F8FAFC" } },
  },
  slideLayout: {
    name: "Product Blank",
    margin: {
      left: asEmu(120000),
      right: asEmu(120000),
      top: asEmu(80000),
      bottom: asEmu(80000),
    },
  },
});

const master = source.slideMasters[0];
const layout = source.slideLayouts[0];
if (master?.handle === undefined || layout?.handle === undefined) {
  throw new Error("Missing template");
}

source = addTextBox(source, master.handle, {
  offsetX: asEmu(300000),
  offsetY: asEmu(180000),
  width: asEmu(3000000),
  height: asEmu(500000),
  text: "Inherited master text",
});
source = addSlideNumber(source, master.handle, {
  offsetX: asEmu(8200000),
  offsetY: asEmu(4650000),
  width: asEmu(500000),
  height: asEmu(300000),
});
source = addTextBox(source, layout.handle, {
  offsetX: asEmu(300000),
  offsetY: asEmu(900000),
  width: asEmu(3000000),
  height: asEmu(500000),
  text: "Inherited layout text",
});
source = addEmptySlideFromLayout(source, { layoutPartPath: layout.partPath });
const authoredSlide = source.slides.at(-1);
if (authoredSlide?.handle === undefined) throw new Error("Missing authored slide");
source = addTextBox(source, authoredSlide.handle, {
  offsetX: asEmu(300000),
  offsetY: asEmu(1600000),
  width: asEmu(3000000),
  height: asEmu(500000),
  text: "Uses the layout's default margins",
});

await writeFile("authored-master.pptx", writePptx(source));
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
