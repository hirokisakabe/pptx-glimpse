# Development setup

This guide covers the basic workflow for developing pptx-glimpse locally.

## Prerequisites

- Node.js 22 or later
- pnpm 11.5.0, as specified by the root `package.json`

## Install dependencies

From the repository root, install the workspace dependencies:

```bash
pnpm install
```

## Run the local editor preview

Pass a PPTX file to the development server:

```bash
pnpm run dev -- presentation.pptx
```

The server automatically reloads when files under `packages/*/src/` change.

## Validate changes

Run the checks that apply to your change before opening a pull request:

```bash
pnpm run knip
pnpm run lint
pnpm run format:check
pnpm run typecheck
pnpm run test
pnpm run build
```

To run a single test file while iterating:

```bash
pnpm run test -- packages/renderer/src/utils/emu.test.ts
```

See the root [`package.json`](../../package.json) for additional scripts, including visual
regression testing, package verification, and fixture generation.
