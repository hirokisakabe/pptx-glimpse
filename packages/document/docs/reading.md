# Reading PPTX source data

Use `readPptx(input)` to parse PPTX bytes into a `PptxSourceModel`:

```ts
import { readFile } from "node:fs/promises";
import { readPptx } from "@pptx-glimpse/document";

const source = readPptx(await readFile("input.pptx"));

console.log(source.presentation.slideSize);
console.log(source.slides.map((slide) => slide.partPath));
console.log(source.diagnostics);
```

The input type is `Uint8Array`; Node.js `Buffer` is accepted because it is a `Uint8Array`
subclass.

## What the source model represents

`PptxSourceModel` is the canonical document representation owned by this package. It groups the
presentation, slides, layouts, masters, themes, relationships, media, content types, diagnostics,
and pending edits as PPTX source semantics. Typed values remain source-local: they describe what
was authored in a particular package part rather than applying inheritance or renderer defaults.

The reader deliberately keeps two complementary forms of source material:

- The typed representation exposes supported PPTX concepts, branded units, stable source handles,
  and relationship/part paths for inspection and supported editing.
- Raw package parts and raw sidecars retain unsupported OOXML, vendor extensions,
  `mc:AlternateContent`, and unmodeled DrawingML for structural preservation.

Typed support does not imply that every OOXML variant is modeled. `source.diagnostics` reports
detected package and reference problems; it is not an exhaustive list of unsupported features. Use
the [feature support matrix](./feature-support.md) when deciding whether a workflow is supported.

## Source values versus effective values

Reading does not resolve the slide → layout → master → theme cascade. When you need effective
backgrounds, theme colors, placeholder matches, inherited text styles, or display order, derive a
[computed view](./computed-view.md). Keep the source model for editing and writing; the computed
view is not an editable replacement for it.

## Public boundary

Import `readPptx`, `ReadPptxInput`, `PptxSourceModel`, and related source types from
`@pptx-glimpse/document`. Reader implementation modules and XML parser helpers are internal.
