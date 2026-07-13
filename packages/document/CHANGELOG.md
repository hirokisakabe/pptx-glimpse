# @pptx-glimpse/document

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
