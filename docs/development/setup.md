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

## TypeScript compiler version

The repository and demo use TypeScript 6. Keep them on the same compiler generation so root
type-checking, declaration emit, and generated API references are evaluated consistently.

TypeScript 7 adoption is currently blocked by the supported compiler ranges of TypeDoc 0.28 and
typescript-eslint 8. Recheck both upstream peer dependencies before upgrading the compiler. The
`stableTypeOrdering` option is only for diagnosing TypeScript 6/7 output differences and must not
be committed as a permanent compiler option.

tsup 8.5.1 injects the deprecated `baseUrl` option into declaration builds even when the project
does not configure it. The workspace applies `patches/tsup.patch` until
[tsup issue #1388](https://github.com/egoist/tsup/issues/1388) is resolved in a release; remove the
patch when upgrading to a version that omits this fallback.

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

See [API reference generation](./api-reference.md) before changing public API JSDoc or updating
TypeDoc and its Markdown plugin.
