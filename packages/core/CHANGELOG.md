# pptx-glimpse

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
