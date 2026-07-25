---
"pptx-glimpse": major
---

Rename the environment-independent high-level editor API from
`BrowserPptxEditorSession` / `createBrowserPptxEditorSession` and `BrowserEditor*` types to
`PptxEditorSession` / `createPptxEditorSession` and `PptxEditor*`, and export the same API from the
Node.js and browser entries.
