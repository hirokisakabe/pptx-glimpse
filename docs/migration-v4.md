# Migrating to pptx-glimpse v4

## High-level editor API rename

The high-level editor session is environment-independent and is now exported with the same names
from the Node.js main entry and browser conditional entry. The browser-specific names were removed.

| v3                               | v4                        |
| -------------------------------- | ------------------------- |
| `BrowserPptxEditorSession`       | `PptxEditorSession`       |
| `createBrowserPptxEditorSession` | `createPptxEditorSession` |
| `BrowserEditor*` types           | `PptxEditor*` types       |

Update imports and type annotations:

```ts
import {
  createPptxEditorSession,
  type PptxEditorRenderOptions,
  type PptxEditorShapeInfo,
} from "pptx-glimpse";

const editor = await createPptxEditorSession(pptxBytes, options);
```

There is no compatibility alias for the old names. Session behavior and its `Uint8Array` input and
output contract are unchanged.

In browsers, pass font bytes with `fonts`. In Node.js, pass `skipSystemFonts: false` to opt into OS
font discovery; the session default remains `true`. Resvg WASM is not used by the SVG editor
preview and is only needed for browser PNG conversion through `convertPptxToPng`.
