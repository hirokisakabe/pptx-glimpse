# AGENTS.md

This file provides guidance to AI coding agents, including Codex and Claude Code, when working with code in this repository.

## Project Overview

pptx-glimpse is a TypeScript library that converts PPTX slides to SVG / PNG.
Input: `Uint8Array` (Node.js `Buffer` is accepted as a subclass), Output: SVG string or PNG `Uint8Array`.

## Commands

```bash
npm run build          # Build with tsup (CJS + ESM + .d.ts)
npm run test           # Run all tests with vitest
npm run test -- packages/renderer/src/utils/emu.test.ts  # Run a single test file
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

CI consists of 6 jobs:

- **lint**: `knip` → `lint` → `format:check` → `typecheck` (Node 22 only, runs once)
- **test**: `test` with coverage → `build` → package verification (Node 22/24/26)
- **playwright**: Browser E2E tests and demo production verification (Node 22/24/26)
- **vrt**: Snapshot VRT (Docker-based, self-comparison)
- **libreoffice-vrt**: LibreOffice VRT for renderer regressions (generates fixtures and reference images via Docker)
- **editor-validity**: LibreOffice validity checks for editing API output (shares the libreoffice-vrt Docker image, runs independently from renderer VRT)

## Architecture

Read [`docs/architecture/overview.md`](docs/architecture/overview.md) before changing package
responsibilities, workspace dependencies, public/private package status, build entries,
externalization/bundling, the document-to-renderer adapter, or Node/browser boundaries. Update
the overview in the same change when one of those boundaries changes.

The working summary is: **PPTX binary → PptxSourceModel → computed view → core adapter →
private renderer → SVG → optional PNG**. `@pptx-glimpse/document` is the lower-level OOXML
foundation; `@pptx-glimpse/editor` builds on it; public `pptx-glimpse` orchestrates document,
editor, and rendering behavior; the demo/UI consumes public packages. Lower layers must not
depend on higher layers.

### External consumer boundary

pptx-glimpse packages must not depend on external consumers or adopt consumer-specific
concepts as repository responsibilities. Do not add consumer-specific types, adapters, node
IDs, fixture generators, validation branches, or runtime/development dependencies to
document, editor, core, renderer, or other workspace packages. In particular,
`@pptx-glimpse/document` remains the lower-level PPTX foundation and must not depend on core,
renderer, editor UI, or any external consumer.

External consumers may be referenced only as usage examples, background context, or input to
compatibility checks against pptx-glimpse's public APIs and generic PPTX/OOXML contracts.
Consumer-specific conversion and end-to-end compatibility tests belong in the consumer's own
repository. Keep fixtures and repository-level tests based on public pptx-glimpse APIs and
generic, source-attributed PPTX inputs rather than consumer packages, source formats, or
semantics.

Read [`docs/editor-error-contract.md`](docs/editor-error-contract.md) before adding or changing
editor operations, operation failure codes, high-level editor error wrapping, warning transport,
or read/render/write catches in `PptxEditorSession`.

Before changing PptxSourceModel, computed-view, writer, or adapter behavior, also read the
module-level comments in `packages/document/src/source/pptx-source-model.ts`,
`packages/document/src/computed/pptx-computed-view.ts`,
`packages/document/src/writer/write-pptx.ts`, and
`packages/core/src/pptx-computed-view-renderer-adapter.ts`.

When adding or changing a `@pptx-glimpse/document` reader, computed view, from-scratch writer, existing-element edit, or round-trip preservation capability, update `packages/document/docs/feature-support.md` in the same change. Base every `S` entry on a public root API and an implementation test; use `△`, `P`, or `—` when support is constrained, preservation-only, or unverified, and keep the linked constraints/evidence current.

Read [`docs/development/type-assertions.md`](docs/development/type-assertions.md) before
adding or changing a type assertion, an `unsafe*Assertion` helper, parser/external boundary
narrowing, a branded constructor, or the ESLint assertion rules. Update the policy when the
allowed exceptions or enforcement changes; update its dated audit snapshot only when running
a deliberate assertion audit.

## Technical Constraints

- **SVG uses inline attributes only** — No CSS classes. resvg and librsvg do not correctly interpret CSS
- **`isArray` configuration in fast-xml-parser is required** — Tags such as `sp`, `pic`, `p`, `r` must be returned as arrays even for single elements (`ARRAY_TAGS` in `xml-parser.ts`)
- **EMU units & branded types** — PPTX internal coordinates use EMU (English Metric Units). Convert with `emuToPixels()`. A 16:9 slide is 9144000×5143500 EMU = 960×540 px. Model fields use branded types (`Emu`, `Pt`, `HundredthPt` in `packages/renderer/src/utils/unit-types.ts`) to prevent unit confusion at compile time. Use `asEmu()`, `asPt()`, `asHundredthPt()` to create branded values from raw numbers
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
│   ├── fixture-builder.ts                    # Shared PPTX fixture scaffolding
│   ├── fixtures-src/                         # Domain-specific fixture creator modules
│   ├── create-fixtures.ts                    # Fixture creator registry + entrypoint
│   ├── update-snapshots.ts                   # Snapshot update script
│   ├── docker-run.sh                         # Docker entrypoint (npm ci + exec)
│   ├── diffs/                                # Diff images on test failure (gitignored)
│   ├── fixtures/                             # VRT PPTX fixtures (dynamically generated)
│   └── snapshots/                            # Reference snapshot images (Docker-generated)
├── libreoffice/                              # LibreOffice VRT (renderer regressions)
│   ├── regression.test.ts                    # Test file
│   ├── create_fixtures.py                    # Fixture generation (Python, Docker)
│   ├── update_snapshots.sh                   # Snapshot update (Docker)
│   ├── diffs/                                # Diff images on test failure (gitignored)
│   ├── fixtures/                             # Dynamically generated in CI
│   └── snapshots/                            # Dynamically generated in CI
└── editor-validity/                          # Editing API output validity (LibreOffice-based)
    ├── editor-validity.test.ts               # Test file
    ├── create_fixtures.py                    # Fixture generation (Python, Docker)
    ├── diffs/                                # Diff images on test failure (gitignored)
    └── fixtures/                             # Dynamically generated in CI
```

### Snapshot VRT (Docker-based)

Snapshots are generated inside a Docker container (Node.js + sharp + fonts) to ensure consistent rendering across macOS and Linux. Both snapshot generation and CI test execution use the same Docker image.

Snapshot VRT uses `vrt/snapshot/render-options.ts` to build a small deterministic
font directory from pinned real font files. The script downloads source fonts into a
temp cache, verifies their SHA256 hashes, subsets them to the glyphs used by the VRT
fixtures, and passes only those subset fonts with `skipSystemFonts: true`. This keeps
local runs from scanning and parsing every system font on macOS while still rendering
readable Latin/CJK glyphs in standard snapshot VRT. The Docker image still fixes the
runtime and rendering toolchain; the VRT intentionally does not depend on each
developer machine's full OS font inventory.

#### Setup

```bash
npm run vrt:snapshot:docker-build   # Build the Docker image
npm run vrt:snapshot:update          # Generate fixtures + snapshots (Docker required)
npm run vrt:snapshot:update -- shapes text  # Regenerate only the named VRT cases
```

### VRT Update Procedure

When changes to the parser, renderer, or model affect rendering output:

1. **Update fixtures** (if adding new features or modifying existing fixtures): Edit the appropriate `vrt/snapshot/fixtures-src/*.ts` domain creator module and run `npm run vrt:snapshot:update`
2. **Update snapshots**: `npm run vrt:snapshot:update` regenerates both fixtures and snapshots in Docker. To update only specific cases, pass the `vrt/snapshot/vrt-cases.ts` `name` values, e.g. `npm run vrt:snapshot:update -- shapes text`.
3. **Verify tests**: Confirm VRT tests pass in CI after pushing

### 3 Locations That Must Stay in Sync

When adding a new rendering feature, **all 3 of the following** must be updated:

1. **`vrt/snapshot/vrt-cases.ts`** — Add a new entry to the `VRT_CASES` array
2. **`vrt/snapshot/fixtures-src/*.ts` + `vrt/snapshot/create-fixtures.ts`** — Add a fixture creator function to the appropriate domain module, export it through that module's creator map, and ensure the map is included in `FIXTURE_CREATORS`
3. **`vrt/snapshot/snapshots/`** — Regenerate snapshots with `npm run vrt:snapshot:update`

`VRT_CASES` is the single source of truth shared by both `create-fixtures.ts` and `regression.test.ts`. If a case is added to `VRT_CASES` without a corresponding creator exported from `fixtures-src/` and included in `FIXTURE_CREATORS`, the fixture generation script will fail with an error.

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

### Editor Validity Tests (Docker-based)

`vrt/editor-validity/` is layer 3 of the editor test strategy (LibreOffice validity): it applies editing API operations to source fixtures, renders both the edited output and an expected fixture with LibreOffice, and compares them. It is separate from the renderer regression VRT above — renderer changes affect `vrt/libreoffice/`, document/editing changes affect `vrt/editor-validity/` — and CI runs them as independent jobs sharing the same Docker image.

```bash
npm run vrt:lo:docker-build              # Build the Docker image (shared with LibreOffice VRT)
npm run vrt:editor-validity:fixtures     # Generate fixtures (Docker required)
npm run test                             # Run tests (including editor validity)
```

Fixtures are source / expected PPTX pairs generated by `vrt/editor-validity/create_fixtures.py`. There are no committed snapshots; both sides are rendered at test time. Tests are automatically skipped without Docker or fixtures.

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
- Tests are colocated with source files (`packages/core/src/converter.test.ts`, etc.)
