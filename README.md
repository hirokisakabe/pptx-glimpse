# pptx-glimpse

[![npm](https://img.shields.io/npm/v/pptx-glimpse)](https://www.npmjs.com/package/pptx-glimpse)
[![CI](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml/badge.svg)](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js](https://img.shields.io/badge/node-%3E%3D22-brightgreen)](https://nodejs.org/)

Render and edit PowerPoint (`.pptx`) files in Node.js or the browser. pptx-glimpse converts slides
to SVG or PNG in-process, without Microsoft Office or LibreOffice.

**[Try the demo](https://glimpse.pptx.app/)** ·
**[Read the documentation](https://glimpse.pptx.app/docs)**

![pptx-glimpse browser rendering and editing demo](https://raw.githubusercontent.com/hirokisakabe/pptx-glimpse/main/docs/demo-editor.png)

## Install

```bash
npm install pptx-glimpse
```

Node.js 22 or later is required for Node.js usage and package tooling.

## Render a presentation

```typescript
import { readFile, writeFile } from "node:fs/promises";
import { convertPptxToPng } from "pptx-glimpse";

const pptx = await readFile("presentation.pptx");
const { slides, diagnostics } = await convertPptxToPng(pptx);

await writeFile("slide-1.png", slides[0].png);
console.log(diagnostics);
```

For SVG output, selected slides, browser font loading, and conversion reports, see
[Rendering presentations](https://glimpse.pptx.app/docs/rendering). Complete signatures and
options are available for
[`convertPptxToPng`](https://glimpse.pptx.app/docs/api/node/functions/convertPptxToPng),
[`convertPptxToSvg`](https://glimpse.pptx.app/docs/api/node/functions/convertPptxToSvg), and
[`ConvertOptions`](https://glimpse.pptx.app/docs/api/node/interfaces/ConvertOptions).

## Edit a presentation

`createPptxEditorSession` provides an integrated read, edit, rerender, history, and save workflow.
See [Editing presentations](https://glimpse.pptx.app/docs/editing) for supported commands and a
complete example. The API reference documents
[`createPptxEditorSession`](https://glimpse.pptx.app/docs/api/node/functions/createPptxEditorSession),
[`PptxEditorSession`](https://glimpse.pptx.app/docs/api/node/classes/PptxEditorSession), and every
[`EditorCommand`](https://glimpse.pptx.app/docs/api/node/type-aliases/EditorCommand) payload.

## Choose a package

| Package                                                                                                        | Use it when you need to…                                     |
| -------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------ |
| [`pptx-glimpse`](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/core/README.md)               | Render slides or build an integrated editing flow            |
| [`@pptx-glimpse/document`](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/README.md) | Read, author, edit, or write PPTX document data directly     |
| [`@pptx-glimpse/editor`](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/editor/README.md)     | Add validated commands and history to a headless application |

Most applications should start with `pptx-glimpse`. The
[package guide](https://glimpse.pptx.app/docs/packages) explains the package boundaries and
dependency direction.

## Documentation

- [Getting started](https://glimpse.pptx.app/docs/getting-started)
- [Why pptx-glimpse](https://glimpse.pptx.app/docs/why)
- [Rendering presentations](https://glimpse.pptx.app/docs/rendering)
- [Editing presentations](https://glimpse.pptx.app/docs/editing)
- [Using fonts](https://glimpse.pptx.app/docs/fonts)
- [Browser usage](https://glimpse.pptx.app/docs/browser)
- [Node.js usage](https://glimpse.pptx.app/docs/nodejs)
- [Feature support](https://glimpse.pptx.app/docs/feature-support)
- [API Reference](https://glimpse.pptx.app/docs/api)

## Contributing

See the [development setup guide](docs/development/setup.md) for repository setup, common
validation commands, and the local editor preview.

## License

MIT
