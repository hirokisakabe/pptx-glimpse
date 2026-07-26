# API reference generation

The API reference at `/docs/api` is generated from the public `pptx-glimpse` Node.js
and browser entry points. TypeScript declarations and JSDoc are the source of truth; do not edit
files under `demo/src/content/` because the directory is deleted and recreated during generation
and is not tracked by Git.

## Generate and validate

From the repository root:

```bash
pnpm run docs:api:generate
pnpm run docs:api:validate
```

The demo runs generation automatically before `dev` and production `build`. CI runs validation
and a production build so invalid JSDoc links, missing required top-level documentation, and
Markdown/MDX incompatibilities fail the change.

When changing a public API, update its JSDoc in the same change. Document runtime restrictions,
defaults, ignored conditions, errors, warnings, parameters, and return values where they are
relevant. Use `@deprecated` with a migration target for deprecated APIs and `@internal` for
symbols that must not appear in the generated reference. Add a symbol to
`intentionallyNotDocumented` or `intentionallyNotExported` in `typedoc.api.json` only when the
omission is deliberate and explainable in review.

## Updating TypeDoc

`typedoc` and `typedoc-plugin-markdown` are exact-version dev dependencies because TypeDoc minor
versions may be breaking for plugins. Before updating:

1. Check the plugin's official compatibility table and peer dependency.
2. Update TypeDoc and the Markdown plugin together with exact versions.
3. Run API generation and validation.
4. Run `npm run build` in `demo/` to verify the generated Markdown with the production Nextra
   compiler.

The generator consumes the plugin's navigation JSON, creates the task-oriented API Reference
landing page and Nextra `_meta.ts` files, and normalizes generated Markdown links to Nextra's
extensionless routes. The sidebar groups symbols into Conversion, Editing, Fonts, Reports and
diagnostics, and Browser-specific API categories. The underlying Node.js and browser entry-point
directories stay out of the sidebar while preserving a direct page for every exported function,
class, interface, type alias, and variable.

Some generated pages are also post-processed for readability. Shared conversion options are
expanded on the functions that accept them, and `EditorCommand` variants are rendered as
individual `kind` and payload blocks. Change the TypeScript/JSDoc source or the generator when
these pages need updating; never edit generated Markdown directly.
