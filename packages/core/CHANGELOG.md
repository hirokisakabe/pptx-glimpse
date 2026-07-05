# pptx-glimpse

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
