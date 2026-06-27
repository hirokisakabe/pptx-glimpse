# pptx-glimpse

## 1.1.2

### Patch Changes

- 0a9a8e9: Improve experimental document path rendering parity with the parser path under default snapshot VRT font conditions.
- 436d171: Fix CleanDoc document-path text font and autofit parity with the current parser path.
- e588f05: Move the published package metadata and build output to the core workspace package.
- 29e9f8c: Switch the public conversion default for `convertPptxToSvg` and `convertPptxToPng` to the CleanDoc document path while keeping an explicit parser-path oracle for parity checks.
- 5a7498a: Move remaining public font collection to the CleanDoc path and keep legacy parser rendering scoped to the internal parity oracle.
