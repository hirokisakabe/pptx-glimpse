# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

pptx-glimpse is a TypeScript library that converts PPTX slides to SVG / PNG.
Input: `Buffer | Uint8Array`, Output: SVG string or PNG Buffer.

## Commands

```bash
npm run build          # Build with tsup (CJS + ESM + .d.ts)
npm run test           # Run all tests with vitest
npm run test -- src/utils/emu.test.ts  # Run a single test file
npm run test:watch     # Watch mode for tests
npm run lint           # ESLint check
npm run lint:fix       # ESLint auto-fix
npm run format         # Prettier formatting
npm run format:check   # Prettier check
npm run typecheck      # Type check with tsc --noEmit
npm run render         # Test rendering with tsx scripts/test-render.ts
npm run inspect        # Inspect PPTX internal XML (e.g., npm run inspect -- file.pptx slide1)
npm run dev -- file.pptx  # Live preview dev server (auto-reload on src/ changes)
```

CI consists of 3 jobs:

- **ci**: `lint` → `format:check` → `typecheck` → `test` (excluding VRT) → `build` (Node 20/22/24)
- **vrt**: Standard VRT (self-comparison)
- **libreoffice-vrt**: LibreOffice VRT (generates fixtures and reference images via Docker)

## Architecture

Data flow: **PPTX binary → Parser (ZIP extraction + XML parsing) → Intermediate model → Renderer (SVG generation) → PNG conversion (optional)**

- `src/parser/` — Builds intermediate model from PPTX via ZIP extraction (`jszip`) and XML parsing (`fast-xml-parser`)
- `src/model/` — TypeScript interfaces for the intermediate model (Slide, Shape, Fill, Text, Theme, Table, Chart, Image, Line, Effect, Presentation, etc.)
- `src/renderer/` — Generates SVG strings from the intermediate model. Includes preset shape definitions in `geometry/`, plus dedicated renderers for tables, charts, and images
- `src/color/` — Theme color resolution (schemeClr → colorMap → colorScheme) and color transformations (lumMod/tint/shade)
- `src/png/` — SVG → PNG conversion using sharp
- `src/data/` — Font metrics data (character width information extracted from OSS-compatible fonts)
- `src/utils/` — EMU ↔ pixel conversion (1 inch = 914400 EMU, 96 DPI), text width measurement, and text wrapping

Entry point: `src/index.ts` exports `convertPptxToSvg` and `convertPptxToPng`.

## Technical Constraints

- **SVG uses inline attributes only** — No CSS classes. sharp (librsvg) does not correctly interpret CSS
- **`isArray` configuration in fast-xml-parser is required** — Tags such as `sp`, `pic`, `p`, `r` must be returned as arrays even for single elements (`ARRAY_TAGS` in `xml-parser.ts`)
- **EMU units** — PPTX internal coordinates use EMU (English Metric Units). Convert with `emuToPixels()`. A 16:9 slide is 9144000×5143500 EMU = 960×540 px
- **Background fallback** — Backgrounds are resolved in order: slide → slide layout → slide master

## VRT (Visual Regression Testing)

Visual regression tests for rendering output. When modifying the parser or renderer, **always check whether VRT updates are needed**.

### Directory Structure

```
vrt/
├── compare-utils.ts                          # Shared image comparison utilities
├── internal/                                 # Standard VRT (self-comparison)
│   ├── regression.test.ts                    # Test file
│   ├── create-fixtures.ts                    # Fixture generation script
│   ├── update-snapshots.ts                   # Snapshot update script
│   ├── fixtures/                             # VRT PPTX fixtures (dynamically generated in CI)
│   └── snapshots/                            # Reference snapshot images
└── libreoffice/                              # LibreOffice VRT
    ├── regression.test.ts                    # Test file
    ├── create_fixtures.py                    # Fixture generation (Python, Docker)
    ├── update_snapshots.sh                   # Snapshot update (Docker)
    ├── fixtures/                             # Dynamically generated in CI
    └── snapshots/                            # Dynamically generated in CI
```

### VRT Update Procedure

When changes to the parser, renderer, or model affect rendering output:

1. **Update fixtures** (if adding new features or modifying existing fixtures): Edit `vrt/internal/create-fixtures.ts` and run `npm run vrt:internal:fixtures`
2. **Update snapshots**: Regenerate reference images with `npm run vrt:internal:update`
3. **Verify tests**: Confirm VRT tests pass with `npm run test`

### 3 Locations That Must Stay in Sync

When adding a new rendering feature, **all 3 of the following** must be updated:

1. **`vrt/internal/create-fixtures.ts`** — Add fixtures (PPTX) covering the new feature
2. **`vrt/internal/regression.test.ts`** — Add new test cases to the `VRT_CASES` array
3. **`vrt/internal/snapshots/`** — Regenerate snapshots with `npm run vrt:internal:update`

**Common mistake**: Modifying the parser or renderer but forgetting to update snapshots, causing VRT tests to fail. Always run `npm run vrt:internal:update` after making changes that affect rendering.

### LibreOffice VRT (Docker-based)

Renders PPTX files generated with python-pptx using LibreOffice and compares them against pptx-glimpse output. Docker ensures a consistent environment.

#### Setup

```bash
npm run vrt:lo:docker-build   # Build the Docker image
npm run vrt:lo:update          # Generate fixtures + reference images (Docker required)
npm run test                   # Run tests (including LibreOffice VRT)
```

#### Tolerance

- `PIXEL_THRESHOLD = 0.3` (per-pixel color difference tolerance)
- `MISMATCH_TOLERANCE = 0.05` (allows up to 5% pixel mismatch)

Since LibreOffice ≠ PowerPoint, differences in font rendering and anti-aliasing are tolerated. The goal is to detect obvious rendering omissions and structural errors.

#### Without Docker

LibreOffice VRT tests are automatically skipped in environments without Docker. `npm run test` will pass without issues.

## Coding Conventions

- Prettier: double quotes, semicolons, trailing commas, printWidth 100
- ESLint: unused variables with `_` prefix are allowed
- ESM (`"type": "module"`) — imports require `.js` extension
- Tests are colocated with source files (`src/parser/slide-parser.test.ts`, etc.)
