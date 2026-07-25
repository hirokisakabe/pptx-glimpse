# Repository Documentation

The root `docs/` directory is for people who develop and maintain this repository. It
separates repository architecture and development policy from package documentation
written for npm users.

## Where documentation belongs

| Location                                                          | Audience and purpose                                                                                                                                                                                                    |
| ----------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [`architecture/`](./architecture/)                                | Repository maintainers. Describes package and layer responsibilities, dependency direction, cross-cutting invariants, and durable design decisions. Start with the [architecture overview](./architecture/overview.md). |
| [`development/`](./development/)                                  | Repository contributors. Contains development policies and repeatable procedures for linting, types, tests, and related engineering practices.                                                                          |
| [`../packages/*/README.md`](../packages/)                         | npm users choosing or starting with a public package. These files document the supported public entry points and common usage.                                                                                          |
| [`../packages/*/docs/`](../packages/)                             | npm users who need detailed package guides, feature support, or API-oriented workflows. Package documentation must describe public root exports rather than repository internals.                                       |
| Files such as `demo.gif` and `comparison-*.png` in this directory | Assets referenced by the root README. Keep existing assets here so their published raw URLs remain stable; add an asset here only when it supports the root README.                                                     |

The root and package READMEs remain the entry points for users of the published packages.
Repository architecture documents may link to those guides but must not replace their
public API and usage documentation. The existing `migration-v4.md` is a user-facing
migration guide linked from those entry points and remains in place to preserve its URL.

## Adding or updating documentation

- Add a document to `architecture/` when it defines responsibilities or invariants across
  modules or packages, explains a dependency boundary, or records a design decision that
  multiple implementation areas must follow.
- Add a document to `development/` when it defines how contributors write, validate, or
  maintain code. Keep time-sensitive audit results visibly separate from lasting policy.
- Put package-specific public API, installation, usage, compatibility, and feature-support
  material beside the owning package.
- Keep module-level comments focused on local invariants. Link them to the relevant
  architecture document instead of copying repository-wide explanations into source files.
- Update an architecture document in the same change when package responsibilities,
  dependency direction, public/private status, build externalization, bundling, or runtime
  boundaries change.
- Update package documentation in the same change when a public root export, supported
  workflow, or user-visible constraint changes.
