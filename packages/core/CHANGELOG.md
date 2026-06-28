# pptx-glimpse

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
