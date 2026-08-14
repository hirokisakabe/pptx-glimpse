# Editor Capability Matrix

This matrix tracks how every public `@pptx-glimpse/document` root editing/authoring API
reaches the headless `EditorCommand`, public `PptxEditorSession`, and
`@pptx-glimpse/editor-react` layers. It records reachability, not rendering fidelity; see
[document feature support](../../packages/document/docs/feature-support.md) for reader, computed,
writer, edit, and preservation coverage.

## Status legend

- **S — supported**: the layer exposes the capability; the manifest records source evidence.
- **I — intentional boundary**: the capability deliberately stops before this layer for the stated
  package-responsibility or product-scope reason.
- **T — tracked**: the capability is not present at this layer and an existing issue owns the work.

The machine-readable source of truth is
[`scripts/editor-capability-manifest.ts`](../../scripts/editor-capability-manifest.ts). The
validation test compares that manifest with document root value exports, every `EditorCommand`
kind, public `PptxEditorSession` members, React command usage, and this rendered table. Therefore
adding or changing a mutation/capability requires updating the manifest and this document in the
same change.

## Matrix

| Capability | Document root API | Document | EditorCommand | Core session | React UI |
| --- | --- | :---: | :---: | :---: | :---: |
| Presentation authoring | `createPptx`<br>`createPptxAuthoringSession` | S (`packages/document/src/index.ts`) | I (From-scratch presentation construction stays in the document layer.) | I (Core edits an already opened presentation.) | I (The reusable editor UI edits a consumer-owned session.) |
| Text content and formatting | `replaceTextRunPlainText`<br>`replaceParagraphPlainText`<br>`setTextRunProperties`<br>`clearTextRunProperties`<br>`setParagraphProperties`<br>`clearParagraphProperties` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#apply`) | S (`packages/editor-react/src/PptxEditor.tsx`) |
| Shape transform | `updateShapeTransform` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#apply`) | S (`packages/editor-react/src/PptxEditor.tsx`) |
| Shape fill and outline | `setShapeFill`<br>`setShapeOutline` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#apply`) | I (The current toolbar exposes text formatting, not general shape styling.) |
| Generic drawing authoring (shape/picture/table/chart) | `addShape`<br>`addPicture`<br>`addTable`<br>`addChart` | S (`packages/document/src/index.ts`) | I (The command layer exposes focused interactive authoring operations only.) | I (Generic document construction remains a lower-level document workflow.) | I (A general-purpose drawing insertion UI is outside the reusable editor scope.) |
| Text box authoring | `addTextBox` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#addTextBox`) | S (`packages/editor-react/src/PptxEditor.tsx`) |
| Connector authoring | `addConnector` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#addConnector`) | I (The current toolbar has no connector insertion control.) |
| Delete drawing | `deleteShape` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#deleteShape`) | S (`packages/editor-react/src/PptxEditor.tsx`) |
| Group / ungroup | `groupShapes`<br>`ungroupShape` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#groupShapes`, `packages/core/src/pptx-editor-session.ts#ungroupShape`) | T ([#818](https://github.com/hirokisakabe/pptx-glimpse/issues/818): Multiple selection and group/ungroup UI are planned.) |
| Move drawings within/across containers | `moveShapes`<br>`moveShapesAcrossSlides` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#moveShapes`, `packages/core/src/pptx-editor-session.ts#moveShapesAcrossSlides`) | I (Direct manipulation currently moves one shape; container transfer has no UI.) |
| Drawing reorder | `reorderShapes` | S (`packages/document/src/index.ts`) | I (Editor uses partial move commands instead of the complete-order API.) | I (Core follows the editor partial-move contract.) | I (Drawing z-order controls are not part of the current UI.) |
| Replace picture media | `replaceImageBytes` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#apply`) | S (`packages/editor-react/src/PptxEditor.tsx`) |
| Picture crop | `setPictureCrop`<br>`clearPictureCrop` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#apply`) | I (Picture crop controls have not been adopted by the reusable UI.) |
| Table cell properties | `setTableCellProperties`<br>`clearTableCellProperties` | S (`packages/document/src/index.ts`) | T ([#846](https://github.com/hirokisakabe/pptx-glimpse/issues/846): Table cell/range commands are planned with the React table editor.) | T ([#846](https://github.com/hirokisakabe/pptx-glimpse/issues/846): Public-session table editing is part of the tracked integration.) | T ([#846](https://github.com/hirokisakabe/pptx-glimpse/issues/846): Table cell selection and editing UI are planned.) |
| Chart data | `updateChartData`<br>`updateScatterChartData`<br>`updateBubbleChartData` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#apply`) | T ([#847](https://github.com/hirokisakabe/pptx-glimpse/issues/847): Chart data and formatting UI are planned.) |
| Slide/master/layout background | `setBackground`<br>`clearBackground`<br>`setSlideBackground` | S (`packages/document/src/index.ts`) | T ([#848](https://github.com/hirokisakabe/pptx-glimpse/issues/848): Background commands are planned.) | T ([#848](https://github.com/hirokisakabe/pptx-glimpse/issues/848): Affected-slide rendering for background edits is planned.) | T ([#848](https://github.com/hirokisakabe/pptx-glimpse/issues/848): Background source display and editing UI are planned.) |
| Theme color/font scheme | `updateThemeScheme` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#updateThemeScheme`) | I (Theme scheme controls are not part of the current editor UI.) |
| Add slide from layout | `addEmptySlideFromLayout` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#apply`) | S (`packages/editor-react/src/PptxEditor.tsx`) |
| Slide duplicate/move/delete | `duplicateSlide`<br>`moveSlide`<br>`deleteSlide` | S (`packages/document/src/index.ts`) | S (`packages/editor/src/index.ts`) | S (`packages/core/src/pptx-editor-session.ts#apply`) | S (`packages/editor-react/src/PptxEditor.tsx`) |
| Layout authoring / cloning | `addSlideLayout`<br>`cloneSlideLayout` | S (`packages/document/src/index.ts`) | I (Template authoring remains a document-layer workflow.) | I (Core exposes template catalog/preview, not template construction.) | I (The layout picker consumes existing layouts only.) |
| Placeholder authoring | `addPlaceholder` | S (`packages/document/src/index.ts`) | I (Placeholder construction remains a document-layer workflow.) | I (Core does not materialize template placeholders.) | I (The reusable UI does not author master/layout placeholders.) |
| Slide-number placeholder authoring | `addSlideNumber` | S (`packages/document/src/index.ts`) | I (Template placeholder construction remains in the document layer.) | I (Core does not author template placeholders.) | I (The reusable UI does not author slide-number placeholders.) |

## Boundary notes

- `@pptx-glimpse/editor-react` continues to import production APIs only from the public `pptx-glimpse` root. This matrix does not authorize lower-layer imports.
- A tracked status names an existing issue; closed implementation issues are evidence for an
  already-supported layer, not a substitute for current source evidence.
- New document authoring/editing exports, editor commands, core session capabilities, or React UI
  operations must be classified as supported, intentional, or tracked. Unclassified additions fail
  `scripts/validate-editor-capabilities.test.ts`.
