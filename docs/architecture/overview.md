# Architecture Overview

This document is the source of truth for repository-level package and layer boundaries in
pptx-glimpse. It is written for maintainers; npm users should start with the
[root README](../../README.md) and the README for the package they use.

## Data flow

The high-level conversion path is:

```text
PPTX binary (Uint8Array)
  -> @pptx-glimpse/document reader
  -> PptxSourceModel
  -> PptxComputedView
  -> core-owned renderer adapter
  -> private renderer model
  -> SVG
  -> PNG (optional resvg rasterization)
```

The reader turns the ZIP/OOXML package into `PptxSourceModel`, the canonical source
representation used by reading, editing, writing, and round-trip preservation. The
non-mutating computed view resolves effective values across slide, layout, master, theme,
and relationship layers while retaining source provenance. Core then adapts those
document semantics to the display-oriented renderer model. The renderer produces SVG;
PNG conversion rasterizes that SVG as an optional final step.

`convertPptxToSvg` and `convertPptxToPng` orchestrate this document path. Consumers that
already own a `PptxSourceModel` can render it without reading the PPTX again. Editing and
writing branch from the source model rather than from the renderer model:

```text
PptxSourceModel -> document editing operations -> document writer -> PPTX binary
        |
        +-> editor commands, selection, validation, and history
```

Node.js AI clients can enter through the optional MCP integration without creating another
rendering path:

```text
AI client -> @pptx-glimpse/mcp (stdio) -> pptx-glimpse public PNG conversion API
          -> MCP image content + structured diagnostics/support coverage
```

## Package and layer responsibilities

### `@pptx-glimpse/document`

`packages/document` is the lower-level OOXML document foundation. It owns the source model,
reader, computed view, from-scratch authoring, supported source edits, writer, typed OOXML
units and handles, diagnostics, and raw material needed for preservation. Its public API is
the root entry point in `packages/document/src/index.ts`.

Document semantics must not depend on core orchestration, editor commands, the renderer
model, SVG/PNG output, UI state, or environment-specific font behavior. Unsupported OOXML
can remain available for preservation or diagnostics without becoming a renderer contract.

### `@pptx-glimpse/editor`

`packages/editor` is a UI-independent command layer over `@pptx-glimpse/document`. It owns
validated commands, selection, warnings, and undo/redo history. It does not own PPTX
parsing/writing, rendering, or application UI. Its public API is the root entry point in
`packages/editor/src/index.ts`.

Expected editor failures, high-level typed errors, integration wrapping, warnings, and
atomicity are specified in the focused [editor error contract](../editor-error-contract.md).

### `pptx-glimpse` (core)

`packages/core` is the high-level public package. It orchestrates reading, computed-view
creation, adaptation, SVG/PNG conversion, font options, diagnostics, and the integrated
editor session exposed to applications. The adapter in
`packages/core/src/pptx-computed-view-renderer-adapter.ts` is the only boundary that turns
document computed values into the renderer's render-ready model.

Core may apply renderer-specific defaults and fallback policy, but those decisions must not
move down into `@pptx-glimpse/document`. Public exports are assembled by
`packages/core/src/index.ts` for the default Node-capable entry and
`packages/core/src/browser.ts` for the browser condition.

### `@pptx-glimpse/renderer`

`packages/renderer` is a private workspace package. It owns the display model, SVG
generation, shape/table/chart/image rendering, text measurement and paths, font mapping,
warning collection, and SVG-to-PNG adapters. It consumes an already adapted rendering
model and must not read or edit `PptxSourceModel`.

The renderer is an implementation boundary, not an npm API that applications may depend
on. Its subpath entry points isolate shared rendering code, Node-only system-font support,
and Node/browser PNG initialization.

### `@pptx-glimpse/mcp`

`packages/mcp` is a public, Node.js-only integration package for local AI clients. It owns
the stdio MCP server, tool schemas, local file loading, PNG content encoding, and MCP Registry
metadata. Its `preview_pptx` tool delegates PPTX parsing and rendering to the public
`convertPptxToPng` API from `pptx-glimpse`; it must not duplicate document or rendering
logic.

MCP SDK and schema dependencies remain private to this integration boundary. Core, document,
editor, renderer, browser entry points, and UI packages must never depend on
`@pptx-glimpse/mcp`.

### Demo and UI

`demo/` is a private Next.js application and an integration consumer of the public
`pptx-glimpse` package. UI components own browser interaction and presentation state; they
must not become dependencies of any workspace package. Reusable headless behavior belongs
in a public package only after it has an intentional public contract.

## Dependency direction

Allowed workspace dependency direction is:

```text
@pptx-glimpse/document
          ^
          |
@pptx-glimpse/editor
          ^
          |
     pptx-glimpse  ------>  @pptx-glimpse/renderer (private, bundled)
          ^
          |
          +------------  @pptx-glimpse/mcp (Node.js stdio)
          |
        demo / UI
```

Core also depends directly on `@pptx-glimpse/document`; the diagram emphasizes that both
document and editor are lower layers. In concrete terms:

- `document` has no dependency on another pptx-glimpse workspace package.
- `editor` may depend on `document`.
- `core` may depend on `document`, `editor`, and `renderer`.
- `mcp` may depend on the public `pptx-glimpse` package.
- demo/UI code may depend on public packages, normally the high-level core package.

Reverse dependencies are forbidden. In particular, `document` must not import from
`editor`, `core`, `renderer`, MCP, or demo/UI; `editor` must not import from `core`,
`renderer`, MCP, or demo/UI; `renderer` must not import document, editor, core, MCP, or
demo/UI semantics; and core must not import MCP. Shared behavior should stay in the lowest
layer that owns the concept, or be passed across an adapter, rather than introducing a
reverse import.

## Publication and build boundaries

The repository root is private and only orchestrates the pnpm workspace. Four packages are
publicly published:

- `@pptx-glimpse/document`
- `@pptx-glimpse/editor`
- `pptx-glimpse` from `packages/core`
- `@pptx-glimpse/mcp`

The reusable library packages expose intentional root entry points and build ESM, CommonJS,
and declaration output with package-specific tsup configurations. Document and editor can
be installed independently for consumers that own the lower-level workflow. The MCP package
builds an ESM-only Node.js executable and declarations because its public runtime contract is
the stdio package bin.

The `pptx-glimpse` package declares `@pptx-glimpse/document` and
`@pptx-glimpse/editor` as runtime dependencies. Its build marks them as `external`, so
their code is not copied into the core bundle and package-manager versioning/deduplication
continues to apply at runtime.

`@pptx-glimpse/renderer` is different: its package is marked `private`, core lists it as a
workspace development/build dependency, and core's tsup configuration includes it through
`noExternal`. Renderer implementation is therefore bundled into the published core
artifacts rather than installed as a public runtime package. The root tsup configuration
mirrors this boundary for manual builds; package-specific configurations define normal
publication.

`@pptx-glimpse/mcp` declares `pptx-glimpse`, the MCP SDK, and its schema library as runtime
dependencies and keeps each external in its build. This preserves the one-way integration
boundary and prevents MCP dependencies from entering core's Node or browser artifacts.

These statements must be checked together when the boundary changes:

- `packages/*/package.json` for names, public/private status, exports, files, and
  dependencies;
- `packages/*/tsup.config.ts` and the root `tsup.config.ts` for entries, external packages,
  and bundled packages;
- `packages/*/src/index.ts` plus any conditional entry such as `packages/core/src/browser.ts`
  for the actual public surface.

## Node and browser runtime policy

The document and editor layers operate on `Uint8Array` data and keep their public behavior
independent of UI and renderer concerns. Node.js `Buffer` works as a `Uint8Array` subtype,
but public binary contracts use `Uint8Array`.

The default core entry supports Node-oriented behavior such as discovering system fonts and
initializing the Node resvg adapter. Node-only imports are isolated behind dedicated
modules or dynamic imports. The `browser` export condition selects
`packages/core/src/browser.ts`, which must remain bundleable without Node built-ins.
Browser consumers provide font buffers explicitly when needed and initialize resvg WASM
before PNG conversion. SVG conversion does not require the PNG/WASM step.

Code shared by both runtimes belongs in the common core/renderer paths. Platform-specific
filesystem, system-font, or WASM-loading behavior belongs behind the corresponding Node or
browser entry point. Browser compatibility is a build-time boundary and must be covered by
the browser-entry smoke test, not inferred from a Node build succeeding.

The MCP package is intentionally outside the browser surface. It may use Node filesystem,
`Buffer`, process, and stdio APIs, but none of its source or dependencies may be imported by
the browser-capable packages.

## When to split another architecture document

Keep this overview focused on package responsibilities, data flow, dependency direction,
publication, and runtime boundaries. Create a separate file under `docs/architecture/`
when a topic:

- defines a substantial invariant or contract within one domain;
- needs lifecycle, failure-mode, or state-transition detail;
- affects several modules but would make this overview harder to scan; or
- needs independent review and change history.

Examples include an editor error contract, document round-trip preservation, or a
diagnostics model. Link the focused document from this overview and from the owning
module-level comments. Do not promote a local implementation note into repository
architecture when a concise module comment is sufficient.
