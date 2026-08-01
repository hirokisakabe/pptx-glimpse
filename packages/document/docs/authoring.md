# Authoring a PPTX from scratch

From-scratch authoring starts with `createPptx()`. The factory creates a valid one-slide source
model with a presentation, blank layout, slide master, theme, content types, and relationships.
Use the public authoring helpers to add content, then pass the returned source to `writePptx`.

This workflow is distinct from [editing an existing PPTX](./editing.md): it constructs supported
typed content and does not rely on preserving an input package's unknown material.

```ts
import { writeFile } from "node:fs/promises";
import { addTextBox, asEmu, createPptx, writePptx } from "@pptx-glimpse/document";

const source = createPptx();
const firstSlide = source.slides[0];

if (firstSlide?.handle === undefined) {
  throw new Error("No slide was created");
}

const authored = addTextBox(source, firstSlide.handle, {
  offsetX: asEmu(914400),
  offsetY: asEmu(914400),
  width: asEmu(3657600),
  height: asEmu(914400),
  text: "Hello from pptx-glimpse",
});

await writeFile("from-scratch.pptx", writePptx(authored));
```

Authoring operations are immutable: each function returns the next `PptxSourceModel`. Public
helpers cover supported text boxes, shapes, connectors, pictures, tables, charts, slide numbers,
backgrounds, and slide topology. Slide, layout, and master handles can be authoring targets where
the operation supports that scope.

## Theme colors and fonts

`createPptx({ theme })` customizes the single theme used by the generated master. The `theme`,
`colorScheme`, `fontScheme`, `major`, and `minor` objects are optional. Every omitted field is
filled independently from the default theme, so callers can override one brand color or font
without copying the complete scheme.

The color scheme accepts six-digit RGB strings for the twelve OOXML slots: `dk1`, `lt1`, `dk2`,
`lt2`, `accent1` through `accent6`, `hlink`, and `folHlink`. Latin fonts default to
`Aptos Display` for major and `Aptos` for minor. East Asian and complex-script fonts default to
an empty OOXML typeface, which lets PowerPoint or LibreOffice choose an appropriate font. Set
those fields when deterministic script-specific selection is required; an explicit empty string
is supported for them. Latin typefaces must be non-empty.

```ts
import { createPptx } from "@pptx-glimpse/document";

const source = createPptx({
  theme: {
    name: "Product Theme",
    colorScheme: {
      name: "Product Colors",
      dk1: "102030",
      accent1: "0067C5",
      accent2: "F59E0B",
    },
    fontScheme: {
      name: "Product Fonts",
      major: {
        latin: "Brand Display",
        eastAsian: "Noto Sans CJK JP",
        complexScript: "Noto Sans Arabic",
      },
      minor: {
        latin: "Brand Text",
        eastAsian: "Noto Sans CJK JP",
        complexScript: "Noto Sans Arabic",
      },
    },
  },
});
```

Shapes, text, and charts in the resulting package that use scheme colors or script-qualified
theme-font tokens such as `+mj-lt`, `+mn-ea`, and `+mj-cs` are resolved through these values by
the normal write/read/computed/render pipeline. This API creates one theme for the generated
master; editing themes in an existing presentation and authoring a format scheme remain outside
its scope.

## Consecutive operations

For consecutive operations, a target-scoped session retains the latest immutable source and
returns each newly created handle without collection searches:

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

The session delegates to the same immutable public authoring functions; it is not a separate
mutable document representation.

### Native groups

Call `target.groupShapes(handles)` after creating two or more consecutive sibling drawings. It
uses the generic drawing operations rather than a group-specific composite input and returns the
new native `p:grpSp` handle:

```ts
const shape = target.addShape({
  geometry: { kind: "preset", preset: "rect" },
  offsetX: asEmu(0),
  offsetY: asEmu(0),
  width: asEmu(914400),
  height: asEmu(914400),
});
const picture = target.addPicture({
  bytes: pngBytes,
  offsetX: asEmu(1143000),
  offsetY: asEmu(0),
  width: asEmu(914400),
  height: asEmu(914400),
});
const group = target.groupShapes([shape, picture]);
session.target(group).reorderShapes([picture, shape]);
```

Supported from-scratch children are generic shapes/text boxes, PNG/JPEG pictures, connectors,
native tables and supported native charts (`p:graphicFrame`), plus previously authored native
groups. The target may be the slide/layout/master or a native group, and every grouped handle must
be a consecutive direct child of that target. SmartArt, raw preserved nodes, consumer node ids,
consumer layout primitives, and connector rerouting are outside this API.

`reorderShapes` accepts a slide, layout, master, or native group target and requires every direct
child drawing in that target exactly once. It never moves a drawing across parents. Because the
operation is applied after earlier additions, a connector can be placed behind its connection
targets without weakening endpoint validation. Non-drawing children of the selected container
remain in place. Missing, duplicate, foreign-part, and non-direct handles are rejected, as are
drawing parts with duplicate node ids or `mc:AlternateContent`; validation is atomic. The
standalone `reorderShapes(source, targetHandle, orderedShapeHandles)` root export provides the same
operation without a session.

## OOXML percentages, angles, and effects

Color transforms and gradient coordinates use OOXML percentages: `0` is 0% and `100000` is 100%.
Use `asOoxmlPercent` rather than passing unbranded numbers. Gradients use `gradientType` as their
discriminator. Radial `centerX` and `centerY` locate the center within the shape bounds.

```ts
import { addShape, asEmu, asOoxmlPercent, createPptx } from "@pptx-glimpse/document";

let source = createPptx();
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

Shape effects accept glow, outer shadow, and inner shadow in supported combinations. Picture
effects accept outer and inner shadows. Shadow radius and distance are EMU integers from `0` to
`2147483647`; direction is an OOXML angle where one degree is 60,000. Outer-shadow alignment is
one of `tl`, `t`, `tr`, `l`, `ctr`, `r`, `bl`, `b`, or `br`.

```ts
import {
  addPicture,
  addShape,
  asEmu,
  asOoxmlAngle,
  asOoxmlPercent,
  createPptx,
} from "@pptx-glimpse/document";

declare const imageBytes: Uint8Array;
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

## Slide master and layout authoring

`createPptx` can name and configure its initial master/layout. The same text, shape, connector,
and picture helpers used for slides accept the generated master or layout handle. A layout margin
is materialized when a text-bearing shape is subsequently authored directly on a slide that
references that layout; it does not rewrite existing or inherited shapes, and explicit per-shape
margin values take precedence.

Additional layouts can be appended to any existing master with `addSlideLayout`, or through
`PptxAuthoringSession.addSlideLayout`. The immutable function returns the updated source, where the
appended layout is available from `source.slideLayouts`. The session method returns the new layout
handle directly; that handle is immediately accepted by the existing drawing target API and by
`addEmptySlideFromLayout` through its `partPath`. The input contract is:

- `name` is required, trimmed, non-empty, and valid in an XML attribute.
- `type` is one of the OOXML `ST_SlideLayoutType` values exposed as `SlideLayoutType`; it defaults
  to `"blank"`. See the [official Open XML SDK enumeration](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.slidelayoutvalues?view=openxml-3.0.1).
- `show` is boolean and defaults to `true`.
- `background` accepts a 6-digit sRGB solid fill or PNG/JPEG bytes.
- every `margin` side is a finite, non-negative EMU value. Margins are authoring defaults that are
  materialized into subsequently authored slide text; OOXML has no layout-level margin metadata,
  so the defaults themselves are not recovered by rereading the package.

```ts
import { asEmu, createPptx, createPptxAuthoringSession } from "@pptx-glimpse/document";

const authoring = createPptxAuthoringSession(createPptx());
const masterHandle = authoring.source.slideMasters[0]?.handle;
if (masterHandle === undefined) throw new Error("Missing master");

const productLayoutHandle = authoring.addSlideLayout(masterHandle, {
  name: "Product",
  type: "titleOnly",
  background: { kind: "solid", color: { kind: "srgb", hex: "F8FAFC" } },
  margin: {
    left: asEmu(120000),
    right: asEmu(120000),
    top: asEmu(80000),
    bottom: asEmu(80000),
  },
});
authoring.target(productLayoutHandle).addShape({
  geometry: { kind: "preset", preset: "rect" },
  offsetX: asEmu(0),
  offsetY: asEmu(5000000),
  width: asEmu(9144000),
  height: asEmu(143500),
  fill: { kind: "solid", color: { kind: "srgb", hex: "4472C4" } },
});
const productSlideHandle = authoring.addEmptySlideFromLayout({
  layoutPartPath: productLayoutHandle.partPath,
});
authoring.target(productSlideHandle).addTextBox({
  offsetX: asEmu(300000),
  offsetY: asEmu(900000),
  width: asEmu(3000000),
  height: asEmu(500000),
  text: "Uses the product layout",
});
```

### Editing existing backgrounds

`setBackground(source, handle, input)` and `clearBackground(source, handle)` are the common
background-editing API for existing slides, layouts, and masters. `setSlideBackground` remains as
a slide-only compatibility API. Direct backgrounds support solid sRGB, linear/radial gradients,
and PNG/JPEG bytes; clearing removes the target's direct `p:bg` and restores slide → layout →
master fallback.

Image edits allocate a new media part and relationship owned by the target slide/layout/master
part. Replaced or cleared image relationships and media are retained as raw package material, so
the operation does not perform an orphan sweep. The writer replaces or removes only `p:cSld/p:bg`:
unknown siblings, namespace declarations needed by retained XML, and unrelated raw sidecars remain
preserved. Unsupported content inside the replaced `p:bg` is intentionally replaced with the new
typed background.

```ts
import { clearBackground, readPptx, setBackground, writePptx } from "@pptx-glimpse/document";

const source = readPptx(templateBytes);
const masterNode = source.slideMasters[0];
const layoutNode = source.slideLayouts.find(
  (candidate) => candidate.partPath === masterNode?.layoutPartPaths[0],
);
const master = masterNode?.handle;
const layout = layoutNode?.handle;
if (master === undefined || layout === undefined || layoutNode.background === undefined) {
  throw new Error("Missing editable template hierarchy or direct layout background");
}

const withMasterBackground = setBackground(source, master, {
  kind: "solid",
  color: { kind: "srgb", hex: "F8FAFC" },
});
const edited = clearBackground(withMasterBackground, layout);
const output = writePptx(edited);
```

An existing layout can be cloned within its current master with `cloneSlideLayout`, or with
`PptxAuthoringSession.cloneSlideLayout`. The required `name` becomes the clone's author-visible
name. `insertAt` is a zero-based insertion position in that master's layout catalog; when omitted,
the clone is inserted immediately after its source. The session method returns the clone handle,
which can immediately be passed to `target`; its `partPath` can also be used to create a slide.

The clone copies the layout XML part and its complete relationship set. Internal relationship
targets such as images, charts, and embedded workbooks continue to reference the same package
resources; those resources are not deep-cloned. External relationships are copied unchanged. The
new layout part path, master relationship ID, and `p:sldLayoutId` numeric ID are allocated without
collisions. Unknown layout XML and heterogeneous child order are retained without parsing and
reordering them; only the requested `p:cSld@name` is changed. Cloning a layout with pending edits
to its XML part is rejected;
write and reread it before cloning when the edited form must be the source template.

```ts
import { readFile } from "node:fs/promises";

import { asEmu, createPptxAuthoringSession, readPptx } from "@pptx-glimpse/document";

const authoring = createPptxAuthoringSession(readPptx(await readFile("template.pptx")));
const sourceLayout = authoring.source.slideLayouts[0];
if (sourceLayout?.handle === undefined) throw new Error("Missing source layout");

const cloneHandle = authoring.cloneSlideLayout(sourceLayout.handle, {
  name: "Product Variant",
  insertAt: 1,
});
authoring.target(cloneHandle).addShape({
  geometry: { kind: "preset", preset: "rect" },
  offsetX: asEmu(0),
  offsetY: asEmu(5000000),
  width: asEmu(9144000),
  height: asEmu(143500),
});
const slideHandle = authoring.addEmptySlideFromLayout({ layoutPartPath: cloneHandle.partPath });
```

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

See the [feature support matrix](./feature-support.md) for the exact authoring subset. Constructing
`PptxSourceModel` internals or importing authoring implementation modules directly is unsupported.
