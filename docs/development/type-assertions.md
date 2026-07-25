# Type Assertion Policy

This document contains the lasting policy for TypeScript assertions in this repository.
The dated inventory at the end is an audit snapshot, not a source of current truth.

## Policy

This repository allows type assertions only when the type system cannot model a boundary
without a local helper. Normal implementation code should prefer control-flow narrowing,
discriminated unions, typed parser helpers, or branded constructors.

## ESLint enforcement

The CI lint path (`npm run lint`) runs these type assertion rules:

- `@typescript-eslint/no-unnecessary-type-assertion`: `error`
- `@typescript-eslint/no-unsafe-type-assertion`: `error`
- `@typescript-eslint/consistent-type-assertions`: `error`, with direct object/array
  literal assertions disallowed

`no-unsafe-type-assertion` is enforced as an error. Unsafe narrowing at XML/OOXML parser
boundaries, fixture boundaries, branded values, and external-library gaps belongs behind
local helper functions grouped by purpose (`unsafeXmlBoundaryAssertion`,
`unsafeOoxmlBoundaryAssertion`, `unsafeFixtureAssertion`, `unsafeBrandAssertion`, external
interop helpers, and adapter/script/VRT variants). Put the reasoned
`eslint-disable-next-line` at the actual assertion boundary.

`unsafeFixtureAssertion` is test-only; ESLint rejects importing it from production package
sources.

Rule references:

- [`@typescript-eslint/no-unnecessary-type-assertion`](https://typescript-eslint.io/rules/no-unnecessary-type-assertion/)
- [`@typescript-eslint/no-unsafe-type-assertion`](https://typescript-eslint.io/rules/no-unsafe-type-assertion/)
- [`@typescript-eslint/consistent-type-assertions`](https://typescript-eslint.io/rules/consistent-type-assertions/)

## Allowed exceptions

- XML and OOXML input parsing may narrow from `unknown` or record-shaped values at the
  boundary, but new code should use shared parser helpers or purpose-specific
  `unsafe*Assertion` helpers before adding an inline assertion.
- Branded numeric or string values must be created through local constructor helpers.
  Direct brand assertions are allowed inside those constructors only.
- Tests may use boundary helpers for fixture narrowing, but repeated patterns should become
  fixture builders or assertion helper functions.
- External library type gaps may use assertions inside the adapter module that owns the
  integration.
- `as const` is allowed for literal preservation. Prefer `satisfies` when shape validation
  is also needed.
- `as any`, new `as unknown as X`, and broad object or array literal assertions are not
  allowed in normal implementation code.

## Audit snapshot (non-normative)

The values in this section describe one point in time and will become stale as the source
tree changes. Run `npm run audit:type-assertions` for the current inventory; update this
snapshot only when performing a deliberate assertion audit. The policy above remains in
force regardless of these counts.

Audit date: 2026-06-27.

At that date, the linted TypeScript source tree contained 84 `as` or angle-bracket type
assertions:

| Category                                                                 | Count | Audit interpretation                                                                       |
| ------------------------------------------------------------------------ | ----: | ------------------------------------------------------------------------------------------ |
| XML / external input boundary (`XmlNode`, `XmlOrderedNode`, parsed JSON) |    11 | Allowed at parser boundaries; prefer shared XML or enum helpers when touching nearby code. |
| Test fixture / mock construction                                         |     0 | Prefer fixture builders or boundary helpers.                                               |
| `as const` literal preservation                                          |    59 | Allowed; prefer `satisfies` when shape validation is also needed.                          |
| Branded unit / handle constructors (`Emu`, `Pt`, `PartPath`, etc.)       |     0 | Callers use constructor helpers.                                                           |
| Double assertion (`as unknown as X`)                                     |     0 | Not allowed in new implementation code.                                                    |
| Object literal assertion                                                 |     0 | Banned by ESLint for direct concrete-type assertions.                                      |
| Array literal assertion                                                  |     0 | Banned by ESLint for direct concrete-type assertions.                                      |
| `as any`                                                                 |     0 | Not allowed in normal implementation code.                                                 |
| External library / platform interop                                      |     0 | Keep allowed gaps inside the owning adapter.                                               |
| Other narrow assertions                                                  |    14 | Review case by case and move repeated narrowing toward helpers.                            |

No `as any` assertions were found in linted TypeScript sources during that audit.
