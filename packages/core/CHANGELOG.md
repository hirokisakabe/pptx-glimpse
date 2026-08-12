# pptx-glimpse

## 5.3.0

### Minor Changes

- 0f69f37: Add typed existing bubble Chart XYZ data editing with synchronized chart formulas, caches, series
  topology, and standard three-column embedded worksheet tables.
- 347f661: Add fixed-topology column and line category combo Chart data editing with typed source identities,
  synchronized caches/formulas/workbooks, ordered axis-aware computed projection and rendering, and
  atomic rejection of unsupported topology including horizontal, stacked, empty-group, and invalid-
  axis combos. Combo rendering also supports zero-only and all-negative primary/secondary domains.
- 981df6b: Allow native charts to move between slide roots while preserving chart/workbook parts, remapping
  owner relationships and selection identity, and retaining chronological chart data edits.
- ddfa337: Add handle-based existing theme color and font scheme editing with field-level XML preservation,
  undo/redo history, and shared-theme-aware slide rerendering.
- 862a64e: Allow cross-parent drawing moves across exactly representable rotated, flipped, and scaled native
  group coordinate spaces while rejecting shear, singular, invalid, and quantization-inexact moves
  atomically.
- 8d99a47: Add identity-mapped cross-parent drawing moves within one slide, layout, or master part, including
  connector-closure validation, headless and integrated editor commands, history, and round-trip
  preservation.
- bf484fd: Add non-mutating one-target SVG previews for ordered slide master and layout catalog handles,
  including template inheritance, structured diagnostics, and matching Node/browser contracts.
- 90afc56: Render supported EMF and WMF pictures, image fills, backgrounds, and OLE preview images in Node.js
  and browser SVG/PNG conversions, with structured warning and placeholder fallback for unsupported or
  invalid metafiles.
- a6383a9: Add typed existing scatter Chart XY data editing with synchronized chart formulas, caches, series
  topology, and standard two-column embedded worksheet tables.
- 027f0a1: Add atomic slide-to-slide moves for consecutive non-placeholder shape, picture, and closed
  connector blocks with destination identity/relationship remapping, editor selection history, and
  affected-slide rerendering.
- 73f0258: Extend the public top-level drawing delete operation to pictures, native Table and Chart graphic
  frames, and native groups, with atomic connector validation and reference-safe recursive package
  cleanup for relationships, media, charts, and embedded workbooks.

### Patch Changes

- be2d83a: Support stable-handle deletion of typed drawings nested at any native-group depth while preserving parent groups, sibling topology, and reference-safe package cleanup.
- Updated dependencies [0f69f37]
- Updated dependencies [b52c6c7]
- Updated dependencies [347f661]
- Updated dependencies [981df6b]
- Updated dependencies [ddfa337]
- Updated dependencies [862a64e]
- Updated dependencies [8d99a47]
- Updated dependencies [bf484fd]
- Updated dependencies [24978d6]
- Updated dependencies [be2d83a]
- Updated dependencies [90f31d7]
- Updated dependencies [a6383a9]
- Updated dependencies [027f0a1]
- Updated dependencies [73f0258]
  - @pptx-glimpse/document@0.16.0
  - @pptx-glimpse/editor@0.6.0

## 5.2.0

### Minor Changes

- a6bf2c7: 既存 picture の stretch fill に対して typed crop の設定・解除を追加し、対象の `a:srcRect` だけを保持的に更新できるようにする。
- e6d14b2: Allow supported transform, fill, and outline edits on existing slide master and slide layout shapes.
- 466af94: integrated editor に authoring 順の slide master / layout catalog API を追加する。
- 8bc6e48: Support adding and removing series with `updateChartData` while keeping chart XML, caches, and the embedded worksheet synchronized.

### Patch Changes

- 16e912b: 共有 media part の画像置換を copy-on-write 化し、選択した picture だけを新しい media と relationship に切り替える。
- 1a9b79c: `@pptx-glimpse/document` に既存の slide、layout、master に共通の background 設定・解除 API を追加する。`pptx-glimpse` では継承 background の変更時に影響する slide だけを再描画する。
- Updated dependencies [bb52f57]
- Updated dependencies [c228c62]
- Updated dependencies [16e912b]
- Updated dependencies [a6bf2c7]
- Updated dependencies [e6d14b2]
- Updated dependencies [6ea2c31]
- Updated dependencies [bd2a2d2]
- Updated dependencies [8bc6e48]
- Updated dependencies [1a9b79c]
- Updated dependencies [1c77833]
  - @pptx-glimpse/document@0.15.0
  - @pptx-glimpse/editor@0.5.0

## 5.1.0

### Minor Changes

- 42c5dd4: Expose from-scratch native group authoring and typed headless/integrated group and ungroup commands with selection-aware undo and redo.

### Patch Changes

- Updated dependencies [9973470]
- Updated dependencies [42c5dd4]
- Updated dependencies [c492b4d]
- Updated dependencies [328de21]
  - @pptx-glimpse/document@0.14.0
  - @pptx-glimpse/editor@0.4.0

## 5.0.0

### Major Changes

- 5075df1: 非推奨の ProseMirror text-body 互換 API（`PptxEditorShapeInfo.editableTextBody`、`PptxEditorTextBodyInfo`、`PptxEditorSession.applyTextBodyDocJson()`）と `prosemirror-model` 依存を削除する。テキスト編集には `PptxEditorShapeInfo.textBody` と `PptxEditorSession.applyAll()` を使用してください。

### Minor Changes

- 2f33268: 既存 Chart の系列名、category label、数値を chart XML と embedded workbook へ一貫して反映する typed document API と editor command / convenience API を追加する。

### Patch Changes

- 38d59ad: group shape の source-local transform を保持し、nested group の描画合成と fallback 診断を固定する
- Updated dependencies [2f33268]
- Updated dependencies [38d59ad]
- Updated dependencies [93ee164]
  - @pptx-glimpse/document@0.13.0
  - @pptx-glimpse/editor@0.3.0

## 4.0.1

### Patch Changes

- 5911629: Editor command、undo、redo の適用時に、影響範囲を安全に特定できる場合は対象 slide だけを再レンダリングして SVG cache を更新するようにしました。特定できない場合は全 slide を再レンダリングします。

## 4.0.0

### Major Changes

- e057459: Rename the environment-independent high-level editor API from
  `BrowserPptxEditorSession` / `createBrowserPptxEditorSession` and `BrowserEditor*` types to
  `PptxEditorSession` / `createPptxEditorSession` and `PptxEditor*`, and export the same API from the
  Node.js and browser entries.

### Minor Changes

- 8e5b47b: Unify headless editor operation failures under a shared discriminated result and expose
  `PptxEditorError` with stable operation and read/render/write codes from both high-level runtime
  entries.

### Patch Changes

- ffa4e0c: Add package-specific English README guidance and align public package metadata with the current
  rendering, document, and editing responsibilities.
- Updated dependencies [ffa4e0c]
- Updated dependencies [f38b91a]
- Updated dependencies [8e5b47b]
  - @pptx-glimpse/document@0.12.1
  - @pptx-glimpse/editor@0.2.0

## 3.3.1

### Patch Changes

- c5d8491: 公開 workspace package の document / editor を runtime dependency として externalize し、private renderer だけを配布 bundle に含める package 境界へ整理する。

## 3.3.0

### Minor Changes

- b5efe3d: UI 非依存の headless 編集 API を `@pptx-glimpse/editor` として公開し、browser editor に command の一括適用と ProseMirror 非依存の text body view を追加する。

### Patch Changes

- Updated dependencies [b5efe3d]
  - @pptx-glimpse/editor@0.1.0

## 3.2.8

### Patch Changes

- Updated dependencies [c60442c]
  - @pptx-glimpse/document@0.12.0

## 3.2.7

### Patch Changes

- Updated dependencies [4c3cd80]
- Updated dependencies [433da8e]
  - @pptx-glimpse/document@0.11.0

## 3.2.6

### Patch Changes

- Updated dependencies [01568a2]
- Updated dependencies [5d60794]
  - @pptx-glimpse/document@0.10.0

## 3.2.5

### Patch Changes

- Updated dependencies [0034bb8]
- Updated dependencies [593b9d5]
- Updated dependencies [fd81dae]
- Updated dependencies [dd05595]
- Updated dependencies [bfb53c9]
- Updated dependencies [04b883d]
- Updated dependencies [3ad0961]
  - @pptx-glimpse/document@0.9.0

## 3.2.4

### Patch Changes

- Updated dependencies [772b4f9]
- Updated dependencies [897d4fd]
  - @pptx-glimpse/document@0.8.0

## 3.2.3

### Patch Changes

- Updated dependencies [b954ce5]
- Updated dependencies [514df94]
  - @pptx-glimpse/document@0.7.0

## 3.2.2

### Patch Changes

- Updated dependencies [f20f4b4]
- Updated dependencies [1d69ee3]
  - @pptx-glimpse/document@0.6.0

## 3.2.1

### Patch Changes

- Updated dependencies [ef9b64e]
  - @pptx-glimpse/document@0.5.0

## 3.2.0

### Minor Changes

- 27fb259: Add a headless `moveSlide` edit operation for reordering existing slides and preserving the updated slide order when writing PPTX files.
- d131e50: Add browser editor support for inserting free connector arrows and allow connector shapes to be deleted through the editing APIs.
- 675d0f0: Add paragraph property editing APIs for alignment, bullet, and paragraph level updates.
- 10c2b13: Add document-layer editing helpers for shape fill and outline styles, including srgb solid fills, line color and width, and noFill.

### Patch Changes

- Updated dependencies [27fb259]
- Updated dependencies [c35eec8]
- Updated dependencies [d131e50]
- Updated dependencies [675d0f0]
- Updated dependencies [10c2b13]
  - @pptx-glimpse/document@0.4.0

## 3.1.1

### Patch Changes

- 7f46470: Preserve numeric-like OOXML text values such as `007`, `1e5`, and `12.50` when reading and writing PPTX slides.
- c50dc1a: Unify new-content edit XML generation at edit time: `addTextBox` / `addConnector` now finalize their shape XML fragment on the edit record and derive the in-memory shape from it, and `addEmptySlideFromLayout` / `duplicateSlide` assign the new `p:sldId` numeric id at edit time. The writer no longer generates new-content XML and only applies insertion positions. The `addTextBox` / `addConnector` / `addEmptySlideFromLayout` / `duplicateSlide` edit record shapes changed accordingly.
- Updated dependencies [c5f2302]
- Updated dependencies [7f46470]
- Updated dependencies [c50dc1a]
  - @pptx-glimpse/document@0.3.0

## 3.1.0

### Minor Changes

- c93f354: Add browser editor image replacement UI support for selected picture shapes.
- 224e24f: Add a standalone `pptx-glimpse` CLI with SVG and PNG conversion commands.
- 1eed18e: Add browser editor text box insertion and selected shape deletion UI/API support.
- f4fe770: Add browser editor slide duplicate and delete controls backed by slide handles.
- f627f71: Add headless text run formatting edits for bold, italic, underline, font size, direct sRGB color, and latin typeface.
- 3d61817: Add headless image media replacement for existing pic shapes, limited to same-format media byte swaps.
- d6f238a: Add `renderPptxSourceModelToSvg` for rendering SVG slides directly from a parsed `PptxSourceModel` without re-reading PPTX bytes.

### Patch Changes

- f0136a9: Add headless slide duplicate/delete editing support with package relationship, content type, and ID management.
- c57532b: Add headless empty slide creation from a slide layout, including writer package bookkeeping and editor-core command support.
- Updated dependencies [f0136a9]
- Updated dependencies [c57532b]
- Updated dependencies [f627f71]
- Updated dependencies [3d61817]
- Updated dependencies [020f949]
- Updated dependencies [69ae720]
- Updated dependencies [8904a5c]
  - @pptx-glimpse/document@0.2.0

## 3.0.0

### Major Changes

- cc0b52b: Remove Node.js Buffer types from the public conversion and font-collection APIs in favor of Uint8Array.

### Minor Changes

- 13b554d: Expose browser PNG conversion after explicit resvg WASM initialization and add Playwright coverage for browser-only SVG/PNG conversion.
- 19f0718: Allow `initResvgWasm` to accept externally loaded WASM bytes or a `Response`, enabling browser-like runtimes to initialize PNG conversion without Node.js filesystem loading.
- 0b931dc: Add a browser-only PPTX editor session API for loading, editing, undoing, redoing, rendering, and downloading edited presentations without a Node backend.
- 7794eca: Add a `fonts` conversion option that accepts `ArrayBuffer` or `Uint8Array` font data directly for SVG and PNG rendering without Node.js font file loading.

### Patch Changes

- 27687a6: Fix browser-entry bundling in webpack-based apps so SVG-only browser viewers do not pull in Node or declaration-file artifacts from the PNG path.
- e6aeb72: Add document-path foundation support for shape transform edits, enabling internal writer round-trips for xfrm offset and extent updates.

## 2.0.0

### Major Changes

- 6fe6cd6: Change `convertPptxToSvg` and `convertPptxToPng` to return conversion report objects with `slides`, `diagnostics`, and `supportCoverage` instead of returning slide arrays directly.

### Patch Changes

- 08ceb6a: Replace the SmartArt fallback's legacy parser shape-tree dependency with the document computed diagram drawing contract.
- 2098d03: Narrow renderer/document image MIME and rectangle alignment token fields to explicit union types.

## 1.1.2

### Patch Changes

- 0a9a8e9: Improve experimental document path rendering parity with the parser path under default snapshot VRT font conditions.
- 436d171: Fix PptxSourceModel document-path text font and autofit parity with the current parser path.
- e588f05: Move the published package metadata and build output to the core workspace package.
- 29e9f8c: Switch the public conversion default for `convertPptxToSvg` and `convertPptxToPng` to the PptxSourceModel document path while keeping an explicit parser-path oracle for parity checks.
- 5a7498a: Move remaining public font collection to the PptxSourceModel path and keep legacy parser rendering scoped to the internal parity oracle.
