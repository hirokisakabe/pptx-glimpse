# @pptx-glimpse/document

## 0.16.0

### Minor Changes

- 0f69f37: Add typed existing bubble Chart XYZ data editing with synchronized chart formulas, caches, series
  topology, and standard three-column embedded worksheet tables.
- 347f661: Add fixed-topology column and line category combo Chart data editing with typed source identities,
  synchronized caches/formulas/workbooks, ordered axis-aware computed projection and rendering, and
  atomic rejection of unsupported topology including horizontal, stacked, empty-group, and invalid-
  axis combos. Combo rendering also supports zero-only and all-negative primary/secondary domains.
- ddfa337: Add handle-based existing theme color and font scheme editing with field-level XML preservation,
  undo/redo history, and shared-theme-aware slide rerendering.
- bf484fd: Add non-mutating one-target SVG previews for ordered slide master and layout catalog handles,
  including template inheritance, structured diagnostics, and matching Node/browser contracts.
- 24978d6: Add native title/body placeholder authoring for masters, title/body/centered-title/subtitle
  placeholder authoring for layouts, and layout-driven slide placeholder materialization.
- a6383a9: Add typed existing scatter Chart XY data editing with synchronized chart formulas, caches, series
  topology, and standard two-column embedded worksheet tables.

## 0.15.0

### Minor Changes

- bb52f57: Add immutable and authoring-session APIs for cloning an existing slide layout within its master while preserving raw XML and sharing related resources.
- c228c62: Allow `createPptx` callers to configure theme color and major/minor font schemes.
- a6bf2c7: 既存 picture の stretch fill に対して typed crop の設定・解除を追加し、対象の `a:srcRect` だけを保持的に更新できるようにする。
- e6d14b2: Allow supported transform, fill, and outline edits on existing slide master and slide layout shapes.
- 6ea2c31: 既存の slide master に追加の slide layout を authoring し、返された handle を drawing API と slide 作成に利用できる公開 API を追加します。
- bd2a2d2: 既存 native Table cell の fill、四辺 border、margin を typed operation で更新・解除できるようにする。
- 8bc6e48: Support adding and removing series with `updateChartData` while keeping chart XML, caches, and the embedded worksheet synchronized.
- 1a9b79c: `@pptx-glimpse/document` に既存の slide、layout、master に共通の background 設定・解除 API を追加する。`pptx-glimpse` では継承 background の変更時に影響する slide だけを再描画する。
- 1c77833: 公開 `reorderShapes` operation で native group の direct children を完全な z-order として並べ替えられるようにする。

### Patch Changes

- 16e912b: 共有 media part の画像置換を copy-on-write 化し、選択した picture だけを新しい media と relationship に切り替える。

## 0.14.0

### Minor Changes

- 9973470: Add lossless group and ungroup operations for consecutive existing sibling shapes while preserving child ids, handles, z-order, transforms, and internal connector references.
- 42c5dd4: Expose from-scratch native group authoring and typed headless/integrated group and ungroup commands with selection-aware undo and redo.
- c492b4d: Support stable nested group shape handles and recursive transform, fill, and outline editing.
- 328de21: Expose slide master and layout authoring order through the typed source model.

## 0.13.0

### Minor Changes

- 2f33268: 既存 Chart の系列名、category label、数値を chart XML と embedded workbook へ一貫して反映する typed document API と editor command / convenience API を追加する。
- 93ee164: 既存 Table のセル内 paragraph / run を、通常の shape text と同じ公開 text operation・editor command で編集して保存できるようにする。

### Patch Changes

- 38d59ad: group shape の source-local transform を保持し、nested group の描画合成と fallback 診断を固定する

## 0.12.1

### Patch Changes

- ffa4e0c: Add package-specific English README guidance and align public package metadata with the current
  rendering, document, and editing responsibilities.

## 0.12.0

### Minor Changes

- c60442c: Add a typed shape-tree reorder operation for slide, layout, and master authoring targets.

## 0.11.0

### Minor Changes

- 4c3cd80: Add target-scoped authoring sessions that retain the latest source and return new slide and drawing handles.

### Patch Changes

- 433da8e: Fix target-scoped authoring sessions after shape authoring and editing were split into separate modules.

## 0.10.0

### Minor Changes

- 5d60794: Allow authored connectors to specify horizontal and vertical transform flips.

### Patch Changes

- 01568a2: Unify part-local drawing ID and ordering allocation across shape, picture, table, and chart authoring while avoiding reserved and pending-edit IDs.

## 0.9.2

### Patch Changes

- f16b0f5: Allow native connector outlines to specify fill, width, and dash styles alongside arrow endpoints.
- 79c3c55: Strengthen native connector endpoint validation and preserve documented connection semantics across slide duplication.

## 0.9.1

### Patch Changes

- 4e00197: writePptx で shape、picture、graphicFrame、および table cell 内の run と line break の追加順を保持する

## 0.9.0

### Minor Changes

- 0034bb8: text・shape authoring の色に alpha transform を追加し、linear / radial gradient と gradient outline を typed input から生成できるようにしました。既存の linear gradient input には `gradientType: "linear"` を追加してください。
- 593b9d5: `addShape` で preset adjustment、custom geometry、水平・垂直 flip、line の width または height 0 を指定できるようにする。従来の `preset` input は `geometry: { kind: "preset", preset }` に移行する
- fd81dae: テキストボックスと図形テキストの shape auto-fit、段落インデント、割合行間、箇条書き・自動番号、明示 baseline を typed authoring API に追加します。
- dd05595: Chart authoring に、chart area、plot area、title、axis、series、marker、data point、blank display、manual layout 向けの型付き書式入力を追加します。
- bfb53c9: shape と picture の authoring API で outer shadow と inner shadow を指定できるようにする
- 04b883d: Add a typed authoring API for solid, linear gradient, radial gradient, PNG, and JPEG slide backgrounds.
- 3ad0961: Table authoring でセル余白、拡張 run 書式、全 preset border dash、run 内改行を指定できるようにする。

## 0.8.0

### Minor Changes

- 772b4f9: from-scratch writer に、スライドマスター／レイアウトの名前・背景・既定テキストマージン、オブジェクト追加、スライド番号フィールドの authoring API を追加しました。
- 897d4fd: text box / shape authoring の run に外部 HTTP(S) hyperlink を指定できるようにしました。

## 0.7.0

### Minor Changes

- b954ce5: from-scratch writer に、セル書式・結合・ハイパーリンクを含む native Table 生成 API を追加
- 514df94: from-scratch writer に native Chart と編集可能な embedded workbook の生成 API を追加しました。

## 0.6.0

### Minor Changes

- f20f4b4: Add `addPicture` for from-scratch writer flows, including PNG/JPEG media part creation, slide image relationships, content type registration, and `p:pic` XML generation.
- 1d69ee3: Add a from-scratch writer API for preset geometry shapes with fill, line, glow, rotation, and text body options.

## 0.5.0

### Minor Changes

- ef9b64e: Add from-scratch text box formatting options for runs, paragraphs, text bodies, and rotation.

## 0.4.0

### Minor Changes

- 27fb259: Add a headless `moveSlide` edit operation for reordering existing slides and preserving the updated slide order when writing PPTX files.
- c35eec8: Add `createPptx()` for constructing a minimal from-scratch `PptxSourceModel` that can be edited with `addTextBox()` and written with `writePptx()`.
- d131e50: Add browser editor support for inserting free connector arrows and allow connector shapes to be deleted through the editing APIs.
- 675d0f0: Add paragraph property editing APIs for alignment, bullet, and paragraph level updates.
- 10c2b13: Add document-layer editing helpers for shape fill and outline styles, including srgb solid fills, line color and width, and noFill.

## 0.3.0

### Minor Changes

- c50dc1a: Unify new-content edit XML generation at edit time: `addTextBox` / `addConnector` now finalize their shape XML fragment on the edit record and derive the in-memory shape from it, and `addEmptySlideFromLayout` / `duplicateSlide` assign the new `p:sldId` numeric id at edit time. The writer no longer generates new-content XML and only applies insertion positions. The `addTextBox` / `addConnector` / `addEmptySlideFromLayout` / `duplicateSlide` edit record shapes changed accordingly.

### Patch Changes

- c5f2302: Make text run replacement and shape transform updates idempotent when they do not change the source model.
- 7f46470: Preserve numeric-like OOXML text values such as `007`, `1e5`, and `12.50` when reading and writing PPTX slides.

## 0.2.0

### Minor Changes

- f627f71: Add headless text run formatting edits for bold, italic, underline, font size, direct sRGB color, and latin typeface.
- 3d61817: Add headless image media replacement for existing pic shapes, limited to same-format media byte swaps.
- 020f949: Add PptxSourceModel editing and writer support for inserting native PowerPoint connector shapes.
- 8904a5c: Expose `p:sldLayout@show` as `SourceSlideLayout.show` for detecting hidden slide layouts.

### Patch Changes

- f0136a9: Add headless slide duplicate/delete editing support with package relationship, content type, and ID management.
- c57532b: Add headless empty slide creation from a slide layout, including writer package bookkeeping and editor-core command support.
- 69ae720: Add PptxSourceModel writer operations for adding text boxes and deleting top-level shapes.

## 0.1.0

### Minor Changes

- b32b8a8: Publish the document package as an installable public 0.x package with README guidance.
