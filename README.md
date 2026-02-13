# pptx-glimpse

[![npm](https://img.shields.io/npm/v/pptx-glimpse)](https://www.npmjs.com/package/pptx-glimpse)

A TypeScript library to render PPTX slides as SVG / PNG.

[Demo](https://hirokisakabe.github.io/pptx-glimpse/) | [npm](https://www.npmjs.com/package/pptx-glimpse)

## Motivation

pptx-glimpse is designed for two primary use cases:

- **Frontend PPTX preview** — Render slide thumbnails without depending on
  Microsoft Office or LibreOffice. The SVG output can be embedded directly
  in web pages.
- **AI image recognition** — Convert slides to PNG so that vision-capable LLMs
  can understand slide content and layout.

The library focuses on accurately reproducing text, shapes, and spatial layout
rather than pixel-perfect rendering of every PowerPoint feature.

## Requirements

- **Node.js >= 20** (does not work in browser environments)
- Requires a platform supported by [sharp](https://sharp.pixelplumbing.com/), which is used for PNG conversion

## Installation

```bash
npm install pptx-glimpse
```

## Usage

```typescript
import { readFileSync } from "fs";
import { convertPptxToSvg, convertPptxToPng } from "pptx-glimpse";

const pptx = readFileSync("presentation.pptx");

// Convert to SVG
const svgResults = await convertPptxToSvg(pptx);
// [{ slideNumber: 1, svg: "<svg>...</svg>" }, ...]

// Convert to PNG
const pngResults = await convertPptxToPng(pptx);
// [{ slideNumber: 1, png: Buffer, width: 960, height: 540 }, ...]
```

## Feature Support

Supports conversion of static visual content in PowerPoint. Dynamic elements such as animations and transitions are not supported.

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

| Feature          | Details                                                                                    |
| ---------------- | ------------------------------------------------------------------------------------------ |
| Supported charts | Bar chart (vertical/horizontal), line chart, pie chart, scatter plot                       |
| Chart elements   | Title, legend (position), series (name/values/categories/color), category axis, value axis |

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

### Unsupported Features

| Category       | Unsupported features                                                                           |
| -------------- | ---------------------------------------------------------------------------------------------- |
| Fill           | Path gradient (rect/shape type)                                                                |
| Effects        | Reflection, 3D rotation / extrusion, artistic effects                                          |
| Charts         | Area, radar, doughnut, bubble, stock, combo, histogram, box plot, waterfall, treemap, sunburst |
| Chart details  | Data labels, axis titles / tick marks / grid lines, error bars, trendlines                     |
| Text           | Vertical text, individual text effects (shadow/glow), text outline, text columns               |
| Tables         | Table style template application, diagonal borders                                             |
| Shapes         | Shape operations (Union/Subtract/Intersect/Fragment)                                           |
| Multimedia     | Embedded video / audio                                                                         |
| Animations     | Object animations, slide transitions                                                           |
| Slide elements | Slide notes, comments, headers / footers, slide numbers / dates                                |
| Image formats  | EMF/WMF (parsed but not rendered)                                                              |
| Other          | Macros / VBA, sections, zoom slides                                                            |

## Test Rendering

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

## License

MIT
