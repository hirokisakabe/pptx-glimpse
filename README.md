# pptx-glimpse

[![npm](https://img.shields.io/npm/v/pptx-glimpse)](https://www.npmjs.com/package/pptx-glimpse)
[![CI](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml/badge.svg)](https://github.com/hirokisakabe/pptx-glimpse/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js](https://img.shields.io/badge/node-%3E%3D22-brightgreen)](https://nodejs.org/)

No LibreOffice required — just `npm install`.

A lightweight JavaScript library that renders PowerPoint (.pptx) slides as SVG or PNG in Node.js and the browser.

**[Try the Demo](https://glimpse.pptx.app/)** | [npm](https://www.npmjs.com/package/pptx-glimpse)

![pptx-glimpse demo](https://raw.githubusercontent.com/hirokisakabe/pptx-glimpse/main/docs/demo.gif)

_Upload a .pptx file → get SVG/PNG output instantly_

|                                                   PowerPoint                                                   |                                                    pptx-glimpse                                                    |
| :------------------------------------------------------------------------------------------------------------: | :----------------------------------------------------------------------------------------------------------------: |
| ![PowerPoint](https://raw.githubusercontent.com/hirokisakabe/pptx-glimpse/main/docs/comparison-powerpoint.png) | ![pptx-glimpse](https://raw.githubusercontent.com/hirokisakabe/pptx-glimpse/main/docs/comparison-pptx-glimpse.png) |

## Motivation

pptx-glimpse is designed for two primary use cases:

- **Frontend PPTX preview** — Render slide thumbnails entirely client-side,
  without a conversion server, Microsoft Office, or LibreOffice. The SVG output
  can be embedded directly in web pages.
- **AI image recognition** — Convert slides to PNG so that vision-capable LLMs
  can understand slide content and layout.

The library focuses on accurately reproducing text, shapes, and spatial layout
rather than pixel-perfect rendering of every PowerPoint feature.

## Why not LibreOffice?

|              | LibreOffice                | pptx-glimpse                             |
| ------------ | -------------------------- | ---------------------------------------- |
| Install size | ~500 MB+                   | `npm install` (~30 MB)                   |
| Docker image | Large base image required  | Works in Node.js images and browser apps |
| Startup time | Process spawning overhead  | In-process, no spawning                  |
| Concurrency  | One process per conversion | Async, runs in your event loop           |

## Requirements

- **Node.js >= 22** for package tooling and Node.js runtime usage.
- **Browser runtime usage** should provide font bytes with the `fonts` option; see
  [Custom Font Loading](#custom-font-loading). PNG conversion in browser-like
  runtimes also requires explicit resvg WASM initialization; see
  [Browser resvg WASM Loading for PNG](#browser-resvg-wasm-loading-for-png).

## Installation

```bash
npm install pptx-glimpse
```

## Usage

```typescript
import { readFileSync, writeFileSync } from "fs";
import { convertPptxToSvg, convertPptxToPng } from "pptx-glimpse";

const pptx = readFileSync("presentation.pptx");

// Convert to SVG
const { slides: svgResults } = await convertPptxToSvg(pptx);
// [{ slideNumber: 1, svg: "<svg>...</svg>" }, ...]

// Convert to PNG
const { slides: pngResults } = await convertPptxToPng(pptx);
// [{ slideNumber: 1, png: Uint8Array, width: 960, height: 540 }, ...]

writeFileSync("slide1.png", pngResults[0].png);
```

The conversion APIs return a report object:

```typescript
const { slides, diagnostics, supportCoverage } = await convertPptxToSvg(pptx);
```

- `slides` contains the converted SVG or PNG slide results.
- `diagnostics` contains document reader, computed view, renderer adapter, and renderer warning diagnostics collected during conversion.
- `supportCoverage` summarizes support/renderability counts such as input elements, output elements, skipped elements, unresolved elements, fallback elements, and warnings. It is not a PowerPoint visual-match or pixel-accuracy score.

### Options

Both `convertPptxToSvg` and `convertPptxToPng` accept an optional `ConvertOptions` object.

```typescript
const { slides } = await convertPptxToPng(pptx, {
  slides: [1, 3], // Convert only slides 1 and 3
  width: 1920, // Output width in pixels (default: 960)
  height: 1080, // Currently ignored by public conversion; width controls PNG size
  logLevel: "warn", // Warning log level: "off" | "warn" | "debug"
  fontDirs: ["/custom/fonts"], // Additional font directories to search
  skipSystemFonts: true, // Skip OS system font directories; use fontDirs only
  fontMapping: {
    "Custom Corp Font": "Noto Sans", // Custom font name mapping
  },
});
```

#### Text output mode (`textOutput`)

By default, text is converted to `<path>` outlines, which renders consistently in any environment but bypasses the browser's native text rasterization (hinting, text-specific anti-aliasing), so glyph edges can look jagged when SVGs are displayed inline in a browser.

With `textOutput: "text"`, `convertPptxToSvg` emits native `<text>` elements along with `@font-face` rules that embed subsetted fonts as data URIs. This produces smoother text rendering in browsers, smaller SVGs for CJK-heavy slides, and selectable/copyable text. Characters not covered by the resolved font fall back to the viewer's fonts via the `font-family` chain.

```typescript
// SVG only — convertPptxToPng always uses path output
const { slides: svgResults } = await convertPptxToSvg(pptx, { textOutput: "text" });
```

Note: embedded fonts and `<text>` may not render as expected when the SVG is referenced via `<img src="...svg">` or sanitized. `convertPptxToPng` always uses path output regardless of this option.

### Advanced Usage

#### Rendering from a parsed source model

When rendering slides repeatedly, keep the parsed `PptxSourceModel` from
`@pptx-glimpse/document` and call `renderPptxSourceModelToSvg`. This avoids
unzipping and parsing the PPTX bytes again on each render.

```bash
npm install pptx-glimpse @pptx-glimpse/document
```

```typescript
import { readPptx } from "@pptx-glimpse/document";
import { renderPptxSourceModelToSvg } from "pptx-glimpse";

const source = readPptx(pptx);

const first = await renderPptxSourceModelToSvg(source, { slides: [1] });
const second = await renderPptxSourceModelToSvg(source, { slides: [2] });
```

<details>
<summary>Font Utilities</summary>

#### Collecting used fonts

`collectUsedFonts` parses a PPTX file and returns all font names used across slides — without performing a full render. Useful for pre-checking which fonts need to be installed.

```typescript
import { collectUsedFonts } from "pptx-glimpse";

const fonts = collectUsedFonts(pptx);
// {
//   theme: { majorFont: "Calibri Light", minorFont: "Calibri", majorFontEa: "...", ... },
//   fonts: ["Arial", "Calibri", "Meiryo"]
// }
```

#### Font mapping helpers

```typescript
import { DEFAULT_FONT_MAPPING, createFontMapping, getMappedFont } from "pptx-glimpse";

// Create a custom mapping (merges with defaults, user values take priority)
const mapping = createFontMapping({ Calibri: "Ubuntu" });

// Look up mapped font (case-insensitive)
getMappedFont("Meiryo", mapping); // "Noto Sans JP"
getMappedFont("calibri", mapping); // "Ubuntu"
```

</details>

<a id="custom-font-loading"></a>

<details>
<summary>Custom Font Loading</summary>

When font files are loaded outside Node.js filesystem scanning, pass the bytes directly with the `fonts` option. This is the font-loading path to use for browser or Edge Runtime integrations that fetch fonts from a URL, application bundle, File input, or any other source.

```typescript
import { convertPptxToSvg } from "pptx-glimpse";

const [pptx, inter] = await Promise.all([
  fetch("/slides/report.pptx").then((response) => response.arrayBuffer()),
  fetch("/fonts/Inter-Regular.ttf").then((response) => response.arrayBuffer()),
]);

const { slides } = await convertPptxToSvg(new Uint8Array(pptx), {
  fonts: [{ name: "Inter", data: inter }],
  fontMapping: {
    Arial: "Inter",
    Calibri: "Inter",
  },
  textOutput: "text",
});
```

For lower-level integrations, you can also build font services from font buffers using `createOpentypeSetupFromBuffers`.

```typescript
import { readFileSync } from "fs";
import { createOpentypeSetupFromBuffers } from "pptx-glimpse";

const setup = await createOpentypeSetupFromBuffers([
  { name: "Carlito", data: readFileSync("fonts/Carlito-Regular.ttf") },
  { name: "Noto Sans JP", data: readFileSync("fonts/NotoSansJP-Regular.ttf") },
]);
// setup.measurer — text width measurement
// setup.fontResolver — text-to-SVG-path conversion
```

</details>

<a id="browser-resvg-wasm-loading-for-png"></a>

<details>
<summary>Browser resvg WASM Loading for PNG</summary>

PNG conversion uses `@resvg/resvg-wasm`. In Node.js, `convertPptxToPng` initializes the bundled WASM automatically on first use. In browser-like runtimes that call the PNG conversion path, fetch or bundle the `.wasm` file yourself and pass it to `initResvgWasm` before PNG conversion so no Node.js filesystem loading is needed.

```typescript
import { convertPptxToPng, initResvgWasm } from "pptx-glimpse";

// A Response can be passed directly.
const wasmResponse = await fetch("/assets/resvg.wasm");
await initResvgWasm(wasmResponse);

const pptx = new Uint8Array(await pptxFile.arrayBuffer());
const { slides } = await convertPptxToPng(pptx, {
  fonts: loadedFonts,
  skipSystemFonts: true,
});
```

`initResvgWasm` also accepts raw WASM bytes when your bundler or application already loaded them:

```typescript
import { initResvgWasm } from "pptx-glimpse";

const wasm = await fetch("/assets/resvg.wasm").then((response) => response.arrayBuffer());
await initResvgWasm(wasm);

// Uint8Array is accepted too.
await initResvgWasm(new Uint8Array(wasm));
```

In Node.js, calling `initResvgWasm()` with no arguments preserves the default bundled WASM loading behavior.

</details>

## Fonts

### Automatic Font Loading

pptx-glimpse automatically scans system font directories and loads fonts using [opentype.js](https://opentype.js.org/). Text in SVG output is converted to `<path>` elements, ensuring consistent rendering regardless of the environment.

Default system font directories:

| OS      | Directories                                                  |
| ------- | ------------------------------------------------------------ |
| Linux   | `/usr/share/fonts`, `/usr/local/share/fonts`                 |
| macOS   | `/System/Library/Fonts`, `/Library/Fonts`, `~/Library/Fonts` |
| Windows | `C:\Windows\Fonts`                                           |

Use the `fontDirs` option to add custom font directories. To skip system font scanning entirely and use only `fontDirs` (useful in containers, serverless environments, or when you want to bundle specific fonts to reduce startup time), set `skipSystemFonts: true`.

If your application already has font bytes, use the `fonts` option instead of `fontDirs`. When `fonts` is provided, `fontDirs` and system font scanning are not used:

```typescript
const font = await fetch("/fonts/Inter-Regular.ttf").then((response) => response.arrayBuffer());

await convertPptxToSvg(pptxBytes, {
  fonts: [{ name: "Inter", data: font }],
  fontMapping: { Arial: "Inter" },
});
```

### Font Mapping

PPTX files often reference proprietary fonts (e.g., Calibri, Meiryo). pptx-glimpse maps these to open-source alternatives available on Google Fonts.

Default mapping:

| PPTX Font                           | Mapped to     |
| ----------------------------------- | ------------- |
| Calibri                             | Carlito       |
| Arial                               | Arimo         |
| Times New Roman                     | Tinos         |
| Courier New                         | Cousine       |
| Cambria                             | Caladea       |
| Meiryo / Yu Gothic / MS Gothic etc. | Noto Sans JP  |
| MS Mincho / Yu Mincho etc.          | Noto Serif JP |

You can customize the mapping via the `fontMapping` option:

```typescript
const { slides } = await convertPptxToSvg(pptx, {
  fontMapping: {
    "Custom Corp Font": "Noto Sans", // Add a new mapping
    Arial: "Inter", // Override the default
  },
});
```

## Feature Support

136 preset shapes, charts, tables, SmartArt, gradients, shadows, and more — covering the most common static PowerPoint content.

### Supported Features

#### Shapes

| Feature       | Details                                                                                                                        |
| ------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| Preset shapes | 136 types (rectangles, ellipses, arrows, flowcharts, callouts, stars, math symbols, etc.)                                      |
| Custom shapes | Arbitrary shape drawing using custom paths (moveTo, lnTo, cubicBezTo, quadBezTo, arcTo, close), adjust values / guide formulas |
| Connectors    | Straight / bent / curved connectors, arrow endpoints (headEnd/tailEnd), line style / color / width                             |
| Groups        | Shape grouping, nested groups, group rotation / flip                                                                           |
| Transforms    | Position, size, rotation, flip (flipH/flipV), adjustment values                                                                |

#### Text

| Feature              | Details                                                                                                                                     |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| Character formatting | Font size, font family (East Asian font support), bold, italic, underline, strikethrough, font color, superscript / subscript, hyperlinks   |
| Paragraph formatting | Horizontal alignment (left/center/right/justify), vertical anchor (top/center/bottom), line spacing, before/after paragraph spacing, indent |
| Bullet points        | Character bullets (buChar), auto-numbering (buAutoNum, 9 types), bullet font / color / size                                                 |
| Text boxes           | Word wrap (square/none), auto-fit (noAutofit/normAutofit/spAutofit), font scaling, margins                                                  |
| Word wrapping        | Word wrapping for English, Japanese, and CJK text, wrapping with mixed font sizes                                                           |
| Style inheritance    | Full text style inheritance chain (run → paragraph default → body lstStyle → layout → master → txStyles → defaultTextStyle → theme fonts)   |
| Tab stops / fields   | Tab stop positions, field codes (slide number, date, etc.)                                                                                  |

#### Fill

| Feature      | Details                                                          |
| ------------ | ---------------------------------------------------------------- |
| Solid color  | RGB color specification, transparency                            |
| Gradient     | Linear gradient, radial gradient, multiple gradient stops, angle |
| Image fill   | PNG/JPEG/GIF, stretch mode, cropping (srcRect)                   |
| Pattern fill | Hatching patterns (horizontal, vertical, diagonal, cross, etc.)  |
| Group fill   | Inherit fill from parent group                                   |
| No fill      | noFill specification                                             |

#### Lines & Borders

| Feature    | Details                                                           |
| ---------- | ----------------------------------------------------------------- |
| Line style | Line width, solid color, transparency, lineCap, lineJoin          |
| Dash style | solid, dash, dot, dashDot, lgDash, lgDashDot, sysDash, sysDot     |
| Arrows     | Head / tail arrow endpoints with type, width, and length settings |

#### Colors

| Feature          | Details                                                                  |
| ---------------- | ------------------------------------------------------------------------ |
| Color types      | RGB (srgbClr), theme color (schemeClr), system color (sysClr)            |
| Theme colors     | Color scheme (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink), color map |
| Color transforms | Luminance adjustment (lumMod/lumOff), tint, shade, transparency (alpha)  |

#### Effects

| Feature      | Details                                                |
| ------------ | ------------------------------------------------------ |
| Outer shadow | Blur radius, distance, direction, color / transparency |
| Inner shadow | Blur radius, distance, direction, color / transparency |
| Glow         | Radius, color / transparency                           |
| Soft edge    | Radius                                                 |

#### Images

| Feature        | Details                                                                      |
| -------------- | ---------------------------------------------------------------------------- |
| Image elements | PNG/JPEG/GIF, position / size / rotation / flip, cropping (srcRect), effects |
| Image fill     | Image fill for shapes and backgrounds                                        |

#### Tables

| Feature         | Details                                                                         |
| --------------- | ------------------------------------------------------------------------------- |
| Table structure | Row and column grid, cell merging (gridSpan/rowSpan), row height / column width |
| Cell formatting | Text, fill (solid/gradient/image), borders (top/bottom/left/right, styles)      |

#### Charts

| Feature          | Details                                                                                                   |
| ---------------- | --------------------------------------------------------------------------------------------------------- |
| Supported charts | Bar chart (vertical/horizontal), line chart, pie chart, scatter plot, area chart, doughnut, bubble, radar |
| Chart elements   | Title, legend (position), series (name/values/categories/color), category axis, value axis                |

#### SmartArt

| Feature             | Details                                                                        |
| ------------------- | ------------------------------------------------------------------------------ |
| Pre-rendered shapes | Renders SmartArt using PowerPoint's pre-rendered drawing shapes (drawingN.xml) |
| mc:AlternateContent | Handles AlternateContent fallback mechanism used by SmartArt                   |

#### Background & Slide Settings

| Feature      | Details                                                                |
| ------------ | ---------------------------------------------------------------------- |
| Background   | Solid, gradient, image. Fallback order: slide → layout → master        |
| Slide size   | 16:9, 4:3, custom sizes                                                |
| Theme        | Theme color scheme, theme fonts (majorFont/minorFont), theme font refs |
| showMasterSp | Control visibility of master shapes per slide / layout                 |

<details>
<summary>Unsupported Features</summary>

| Category       | Unsupported features                                                       |
| -------------- | -------------------------------------------------------------------------- |
| Fill           | Path gradient (rect/shape type)                                            |
| Effects        | Reflection, 3D rotation / extrusion, artistic effects                      |
| Charts         | Stock, combo, histogram, box plot, waterfall, treemap, sunburst            |
| Chart details  | Data labels, axis titles / tick marks / grid lines, error bars, trendlines |
| Text           | Individual text effects (shadow/glow), text columns                        |
| Tables         | Table style template application, diagonal borders                         |
| Shapes         | Shape operations (Union/Subtract/Intersect/Fragment)                       |
| Multimedia     | Embedded video / audio                                                     |
| Animations     | Object animations, slide transitions                                       |
| Slide elements | Slide notes, comments, headers / footers, slide numbers / dates            |
| Image formats  | EMF/WMF (parsed but not rendered)                                          |
| Other          | Macros / VBA, sections, zoom slides                                        |

</details>

## Development

The local editor preview (`npm run dev -- <pptx-file>`) includes an MVP text editing
overlay for text shapes. IME behavior is intentionally not automated in CI; verify IME
composition manually as part of the release checklist before shipping editor changes.

Browser conversion smoke tests run with Playwright and cover browser-only SVG conversion
plus PNG conversion after explicit resvg WASM initialization:

```bash
npm run test:playwright
```

<details>
<summary>Test rendering</summary>

You can specify a PPTX file to preview SVG and PNG conversion results.

```bash
npm run render -- <pptx-file> [output-dir]
```

- If `output-dir` is omitted, output is saved to `./output`

```bash
# Examples
npm run render -- presentation.pptx
npm run render -- presentation.pptx ./my-output
```

</details>

## License

MIT
