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

`editor.layoutCatalog` exposes slide masters in presentation authoring order and nests each
master's layouts in its authoring order. Master and layout `handle` values identify the OOXML root
part and can be retained as stable UI identities. Layout entries also expose `name`, `type`,
`hidden`, and `slideReferenceCount`; omitted `p:sldLayout@show` is treated as visible, and the
reference count includes slides directly assigned to that layout.

```ts
for (const master of editor.layoutCatalog) {
  for (const layout of master.layouts) {
    console.log(layout.name, layout.hidden, layout.slideReferenceCount, layout.handle);
  }
}
```

Render one catalog entry on demand with `previewLayoutCatalogTarget`. The catalog remains
metadata-only; preview SVG, diagnostics, caching, scheduling, and UI fallback are separate concerns.
The call does not add a temporary slide or change the document, selection, history, current slide
SVGs, or serialized PPTX.

```ts
const layout = editor.layoutCatalog[0]?.layouts[0];
if (layout) {
  const preview = await editor.previewLayoutCatalogTarget(layout.handle);
  if (preview.ok) {
    showThumbnail(preview.svg);
    reportDiagnostics(preview.diagnostics);
  } else if (preview.code === "preview-handle-not-found") {
    showFallbackThumbnail();
  }
}
```

Layout previews resolve the layout/master background, normal template shapes, and compatible
placeholder inheritance without reading real slide user content. Missing and ambiguous handles
return `preview-handle-not-found` and `preview-handle-ambiguous`; unsupported elements are skipped
with stable diagnostics. Renderer/runtime failures still throw `PptxEditorError` with
`render-failed`. PNG preview, cache policy, cancellation, priority scheduling, and thumbnail UI are
not part of this API.

Native group topology is available through typed `groupShapes` / `moveShapes` / `ungroupShape` commands and the
matching `PptxEditorSession` methods. A successful group selects the new group; a successful
ungroup selects its first child in document order, while cross-parent moves retain
the current selection. Exactly representable affine cross-parent moves also preserve the rendered
mapping across nested group rotation, flips, and scale; inexact mappings fail as `invalid-command`.
Undo and redo restore topology, ids, z-order,
and the corresponding selection.

Use `editor.moveShapesAcrossSlides(shapeHandles, destinationSlideHandle, options)` for a
consecutive non-placeholder shape/picture/connector block at a slide root. Both source and
destination slides rerender, moved selection follows the returned destination identity, and
save/read preserves the remapped node, relationship, and connector endpoint IDs. Authored local
OOXML is retained, so destination theme/layout/master differences can change effective appearance.

Existing theme schemes can be patched through `editor.updateThemeScheme(themeHandle, input)` or
the matching command. Omitted color/font fields and the format scheme are preserved. When masters
share that theme, the session rerenders slides under those masters only; unrelated themes remain
outside the invalidation set.

Browser applications pass PPTX and font bytes as `Uint8Array` data:

```ts
const editor = await createPptxEditorSession(new Uint8Array(pptx), {
  fonts: [{ name: "Inter", data: inter }],
  textOutput: "text",
});
```

See [Editing presentations](https://glimpse.pptx.app/docs/editing) for commands and workflow
details, including stretch-picture `setPictureCrop` / `clearPictureCrop` commands that rerender
the affected slide and persist the crop change by updating or removing `a:srcRect`. Complete
signatures are available for
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
- EMF and WMF pictures, image fills, and OLE preview pictures are converted to SVG in both Node.js
  and browser builds. The supported subset covers common GDI paths, filled/outlined primitives,
  text, and bitmap records. Unknown records, malformed streams, and conversion errors use a labeled
  placeholder and add an `image.metafile-conversion` warning instead of failing the presentation.
- Expected editor failures throw `PptxEditorError`; successful commands can still return warnings.

### EMF/WMF images

Pictures, image fills, backgrounds, and OLE preview pictures use one synchronous EMF/WMF-to-SVG
path in Node.js and browser builds. The initial subset covers window/viewport state, pen and brush
objects, rectangles and rounded rectangles, lines, polygons, Bezier paths, path clipping, EMF
extended text, WMF text, and WMF DIB bitmap records. The record stream and deterministic fixtures
follow Microsoft's [MS-EMF](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/91c257d7-c39d-4a36-9b1f-63e3f73d30ca)
and [MS-WMF](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wmf/4813e7fd-52d0-4f42-965f-228c8b7488d2)
specifications.

EMF extended-text fallback is intentionally conservative: it supports anisotropic compatible-mode
text with left/top alignment, window/viewport origin and extent mapping, LOGFONT height, weight,
italic state, escapement rotation, and per-character `offDx` advances. Text records using a font
width, incompatible orientation, non-compatible graphics mode, other map/alignment modes, or text
options that cannot be represented accurately fall back with `unsupported-record`.

An unknown record, malformed stream, or conversion error falls back for that metafile as a whole.
SVG output keeps the labeled `[EMF]` or `[WMF]` placeholder, and conversion reports include a
`renderer.image.metafile-conversion` warning whose message identifies `unsupported-record`,
`invalid-data`, or `conversion-failed`. Complete EMF/WMF and EMF+ record fidelity is not supported.
For predictable resource use, conversion rejects decoded metafiles over 8 MiB, individual records
over 4 MiB, bitmap records over 2 MiB, streams over 50,000 records or 200,000 geometry points, and
generated SVG over 100,000 nodes or 8 MiB. The per-render conversion cache retains at most 16
entries and 16 MiB of converted results; every cached SVG insertion receives a distinct,
deterministic fragment-ID namespace.

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
