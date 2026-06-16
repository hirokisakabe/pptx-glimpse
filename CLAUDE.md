# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

pptx-glimpse is a TypeScript library that converts PPTX slides to SVG / PNG.
Input: `Buffer | Uint8Array`, Output: SVG string or PNG Buffer.

## Commands

```bash
npm run build          # Build with tsup (CJS + ESM + .d.ts)
npm run test           # Run all tests with vitest
npm run test -- packages/pptx-glimpse-renderer/src/utils/emu.test.ts  # Run a single test file
npm run test:watch     # Watch mode for tests
npm run lint           # ESLint check
npm run lint:fix       # ESLint auto-fix
npm run format         # Prettier formatting
npm run format:check   # Prettier check
npm run typecheck      # Type check with tsc --noEmit
npm run render         # Test rendering with tsx scripts/test-render.ts
npm run inspect        # Inspect PPTX internal XML (e.g., npm run inspect -- file.pptx slide1)
npm run dev -- file.pptx  # Live preview dev server (auto-reload on packages/*/src/ changes)
```

CI consists of 4 jobs:

- **lint**: `knip` → `lint` → `format:check` → `typecheck` (Node 22 only, 1回実行)
- **test**: `test` with coverage → `build` → package verification (Node 22/24, coverage report on Node 22)
- **vrt**: Snapshot VRT (Docker-based, self-comparison)
- **libreoffice-vrt**: LibreOffice VRT (generates fixtures and reference images via Docker)

## Architecture

Data flow: **PPTX binary → Parser (ZIP extraction + XML parsing) → Intermediate model → Renderer (SVG generation) → PNG conversion (optional)**

ソースは pnpm workspaces (`packages/*`) で分割されている。`pptx-glimpse` パッケージは `pptx-glimpse-renderer` を workspace 依存として参照する。

`packages/pptx-glimpse/src/` — 公開パッケージ `pptx-glimpse` の実装（パーサー + 公開 API）

- `parser/` — Builds intermediate model from PPTX via ZIP extraction (`fflate`) and XML parsing (`fast-xml-parser`)
- `color/` — Theme color resolution (schemeClr → colorMap → colorScheme) and color transformations (lumMod/tint/shade)
- `font/font-collector.ts` — PPTX から使用フォント名を収集する公開 API (`collectUsedFonts`)
- `converter.ts` — `convertPptxToSvg` / `convertPptxToPng` の実装
- `pptx-data-parser.ts`, `text-style-resolver.ts` — パーサー共通ヘルパー
- `index.ts` — 公開エントリポイント

`packages/pptx-glimpse-renderer/src/` — 内部 renderer パッケージ（private; 親 issue #340 決定）

- `renderer/` — Generates SVG strings from the intermediate model. Includes preset shape definitions in `geometry/`, plus dedicated renderers for tables, charts, and images
- `model/` — TypeScript interfaces for the intermediate model (Slide, Shape, Fill, Text, Theme, Table, Chart, Image, Line, Effect, Presentation, etc.)
- `font/` — Font loading (system font scanning), font mapping (proprietary → OSS alternatives), text measurement and text-to-SVG-path conversion via `opentype.js`
- `png/` — SVG → PNG conversion using `@resvg/resvg-wasm`
- `data/` — Font metrics data (fallback character width information)
- `utils/` — EMU ↔ pixel conversion (1 inch = 914400 EMU, 96 DPI) and text wrapping
- `warning-logger.ts` — 共有警告ロガー
- `index.ts` — pptx-glimpse が import する barrel re-export

Entry point: `packages/pptx-glimpse/src/index.ts` exports `convertPptxToSvg`, `convertPptxToPng`, warning utilities (`getWarningSummary`, `getWarningEntries`), font utilities (`collectUsedFonts`, `DEFAULT_FONT_MAPPING`, `createFontMapping`, `getMappedFont`), and related types.

ルートの `tsup.config.ts` が `packages/pptx-glimpse/src/index.ts` を bundle して `dist/` を生成し、`pptx-glimpse` パッケージとして publish する（renderer は `noExternal` で bundle 内に取り込む）。publish 経路の monorepo 対応 (`packages/pptx-glimpse` を直接 publish する形への移行) は #340 子4 で実施予定。

## Technical Constraints

- **SVG uses inline attributes only** — No CSS classes. resvg and librsvg do not correctly interpret CSS
- **`isArray` configuration in fast-xml-parser is required** — Tags such as `sp`, `pic`, `p`, `r` must be returned as arrays even for single elements (`ARRAY_TAGS` in `xml-parser.ts`)
- **EMU units & branded types** — PPTX internal coordinates use EMU (English Metric Units). Convert with `emuToPixels()`. A 16:9 slide is 9144000×5143500 EMU = 960×540 px. Model fields use branded types (`Emu`, `Pt`, `HundredthPt` in `packages/pptx-glimpse-renderer/src/utils/unit-types.ts`) to prevent unit confusion at compile time. Use `asEmu()`, `asPt()`, `asHundredthPt()` to create branded values from raw numbers
- **Background fallback** — Backgrounds are resolved in order: slide → slide layout → slide master

## VRT (Visual Regression Testing)

Visual regression tests for rendering output. When modifying the parser or renderer, **always check whether VRT updates are needed**.

### Directory Structure

```
shared-fixtures/                              # Real PPTX files shared by e2e and VRT
├── real-basic-theme.pptx
└── real-product-page.pptx
vrt/
├── compare-utils.ts                          # Shared image comparison utilities
├── snapshot/                                 # Standard VRT (self-comparison, Docker-based)
│   ├── vrt-cases.ts                          # Shared test case definitions (VRT_CASES + SHARED_FIXTURE_CASES)
│   ├── regression.test.ts                    # Test file
│   ├── create-fixtures.ts                    # Fixture generation script
│   ├── update-snapshots.ts                   # Snapshot update script
│   ├── docker-run.sh                         # Docker entrypoint (npm ci + exec)
│   ├── diffs/                                # Diff images on test failure (gitignored)
│   ├── fixtures/                             # VRT PPTX fixtures (dynamically generated)
│   └── snapshots/                            # Reference snapshot images (Docker-generated)
└── libreoffice/                              # LibreOffice VRT
    ├── regression.test.ts                    # Test file
    ├── create_fixtures.py                    # Fixture generation (Python, Docker)
    ├── update_snapshots.sh                   # Snapshot update (Docker)
    ├── diffs/                                # Diff images on test failure (gitignored)
    ├── fixtures/                             # Dynamically generated in CI
    └── snapshots/                            # Dynamically generated in CI
```

### Snapshot VRT (Docker-based)

Snapshots are generated inside a Docker container (Node.js + sharp + fonts) to ensure consistent rendering across macOS and Linux. Both snapshot generation and CI test execution use the same Docker image.

#### Setup

```bash
npm run vrt:snapshot:docker-build   # Build the Docker image
npm run vrt:snapshot:update          # Generate fixtures + snapshots (Docker required)
```

### VRT Update Procedure

When changes to the parser, renderer, or model affect rendering output:

1. **Update fixtures** (if adding new features or modifying existing fixtures): Edit `vrt/snapshot/create-fixtures.ts` and run `npm run vrt:snapshot:update`
2. **Update snapshots**: `npm run vrt:snapshot:update` regenerates both fixtures and snapshots in Docker
3. **Verify tests**: Confirm VRT tests pass in CI after pushing

### 3 Locations That Must Stay in Sync

When adding a new rendering feature, **all 3 of the following** must be updated:

1. **`vrt/snapshot/vrt-cases.ts`** — Add a new entry to the `VRT_CASES` array
2. **`vrt/snapshot/create-fixtures.ts`** — Add a fixture creator function and register it in `FIXTURE_CREATORS`
3. **`vrt/snapshot/snapshots/`** — Regenerate snapshots with `npm run vrt:snapshot:update`

`VRT_CASES` is the single source of truth shared by both `create-fixtures.ts` and `regression.test.ts`. If a case is added to `VRT_CASES` without a corresponding creator in `FIXTURE_CREATORS`, the fixture generation script will fail with an error.

**Common mistake**: Modifying the parser or renderer but forgetting to update snapshots, causing VRT tests to fail. Always run `npm run vrt:snapshot:update` after making changes that affect rendering.

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
- Each test case has an explicit `tolerance` set to its measured mismatch rate in CI × 1.2, rounded up to 0.1pt (minimum 0.3%). Measured values are printed as `[lo-vrt]` log lines during test runs — use the CI job logs to recalibrate when LibreOffice or runner fonts change
- `MISMATCH_TOLERANCE = 0.02` is the fallback for newly added cases before calibration

Since LibreOffice ≠ PowerPoint, differences in font rendering and anti-aliasing are tolerated. The goal is to detect rendering regressions, omissions, and structural errors.

#### Without Docker

LibreOffice VRT tests are automatically skipped in environments without Docker. `npm run test` will pass without issues.

## Release Workflow (Changesets)

This project uses [Changesets](https://github.com/changesets/changesets) for version management and releases.

### Adding a changeset

When a PR includes changes that affect the published package (bug fixes, new features, breaking changes), run `npx changeset` before committing and select the appropriate version bump type (patch / minor / major) with a summary of the change. This creates a markdown file in `.changeset/` that should be committed with the PR.

Changes that do NOT require a changeset: docs-only updates, CI config, test-only changes, refactoring with no public API impact.

### Release flow

1. PR with changeset is merged to main
2. `release.yml` (changesets/action) automatically creates a "Version Packages" PR that bumps `package.json` and updates `CHANGELOG.md`
3. "Version Packages" PR is reviewed and merged
4. `release.yml` runs `npx changeset publish` — creates a `v{version}` tag and publishes to npm with provenance (Trusted Publishing)

## Coding Conventions

- Prettier: double quotes, semicolons, trailing commas, printWidth 100
- ESLint: unused variables with `_` prefix are allowed
- ESM (`"type": "module"`) — imports require `.js` extension
- Tests are colocated with source files (`packages/pptx-glimpse/src/parser/slide-parser.test.ts`, etc.)
