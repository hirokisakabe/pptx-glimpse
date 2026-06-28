# Legacy Parser Semantics Audit

- Status: retired by [#543](https://github.com/hirokisakabe/pptx-glimpse/issues/543)
- Date: 2026-06-28

This note used to track the temporary legacy parser oracle that protected the
PptxSourceModel document-path migration. The migration is complete: public
conversion, snapshot VRT, LibreOffice VRT, and benchmark preparation now use the
document path.

## Retired Components

The following parser-oracle components were removed:

- `packages/core/src/parser-path-oracle.ts`
- `packages/core/src/pptx-data-parser.ts`
- `packages/core/src/dual-reader-structural-comparison.test.ts`
- `packages/core/src/parser-path-oracle.test.ts`
- `packages/core/src/parse-render.integration.test.ts`
- `packages/core/src/text-style-resolver.ts`
- old render parser helpers under `packages/core/src/parser/`
- `vrt/snapshot/document-path-regression.test.ts`
- `vrt/snapshot/document-path-zero-diff-gate.test.ts`
- `vrt/snapshot/document-path-cases.ts`

The core-local XML parser was moved to `packages/core/src/ooxml/xml-parser.ts`
because chart/color adapter code still needs a small OOXML parsing utility
without depending on the retired render parser subsystem.

## Current Regression Coverage

Regression coverage no longer treats the old parser output as a baseline. The
current coverage stack is:

- focused document reader/computed-view tests in `packages/document/src/`
- focused renderer adapter tests in `packages/core/src/pptx-computed-view-renderer-adapter.test.ts`
- public converter tests in `packages/core/src/converter*.test.ts`
- committed snapshot VRT in `vrt/snapshot/regression.test.ts`
- LibreOffice VRT in `vrt/libreoffice/regression.test.ts`

Historical migration notes remain in the document-boundary and dogfood-migration
docs, but they describe completed migration context rather than an active oracle.
