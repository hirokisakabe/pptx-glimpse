# Type Assertion Policy

This repository allows type assertions only when the type system cannot model a
boundary without a local helper. Normal implementation code should prefer
control-flow narrowing, discriminated unions, typed parser helpers, or branded
constructors.

## Current Inventory

Audit date: 2026-06-27.

Command used for the broad inventory:

```bash
npm run audit:type-assertions
```

The current source tree contains 84 `as` / angle-bracket type assertions in
linted TypeScript sources. The main buckets are:

| Category | Count | Policy |
| --- | ---: | --- |
| XML / external input boundary (`XmlNode`, `XmlOrderedNode`, parsed JSON) | 11 | Allowed at parser boundaries. Prefer shared XML helpers such as `getNode`, `getNodeArray`, `getAttr`, and enum parsers when touching nearby code. Unsafe narrowing belongs in purpose-specific `unsafe*Assertion` boundary helpers, not inline parser code. |
| Test fixture / mock construction | 0 | Prefer fixture builders or boundary helpers when constructing narrow fixtures. |
| `as const` literal preservation | 59 | Allowed. Prefer `satisfies` when shape validation is also needed. |
| Branded unit / handle constructors (`Emu`, `Pt`, `PartPath`, etc.) | 0 | Callers should use constructor helpers such as `asEmu`, `asPt`, source unit helpers, and handle helpers. |
| Double assertion (`as unknown as X`) | 0 | Not allowed in new implementation code. Replace with focused adapter helpers when needed. |
| Object literal assertion | 0 | Direct object/array literal assertions to concrete types are banned by ESLint. Use typed variables, `satisfies`, or fixture factories. |
| Array literal assertion | 0 | New direct array literal assertions to concrete types are banned by ESLint. |
| `as any` | 0 | Not allowed in normal implementation code. |
| External library / platform interop | 0 | Allowed inside the adapter module that owns the integration, preferably via a boundary helper. |
| Other narrow assertions | 14 | Treat case-by-case. Enum/string-literal narrowing should move toward parser helpers such as `parseEnum`. External library gaps should stay inside adapter modules or purpose-specific `unsafe*Assertion` helpers. |

No `as any` assertions were found in linted TypeScript sources during this
audit.

## ESLint Policy

The CI lint path (`npm run lint`) now explicitly runs these type assertion
rules:

- `@typescript-eslint/no-unnecessary-type-assertion`: `error`
- `@typescript-eslint/no-unsafe-type-assertion`: `error`
- `@typescript-eslint/consistent-type-assertions`: `error`, with direct
  object/array literal assertions disallowed

`no-unsafe-type-assertion` is enforced as an error. Existing unsafe narrowing
from XML/OOXML parser boundaries, fixture narrowing, branded values, and
external-library gaps has been moved behind local helper functions grouped by
boundary purpose (`unsafeXmlBoundaryAssertion`,
`unsafeOoxmlBoundaryAssertion`, `unsafeFixtureAssertion`,
`unsafeBrandAssertion`, external interop helpers, and adapter/script/VRT
variants), each with a reasoned `eslint-disable-next-line` at the actual
assertion boundary.
`unsafeFixtureAssertion` is test-only; ESLint rejects importing it from
production package sources.

Rule references:

- [`@typescript-eslint/no-unnecessary-type-assertion`](https://typescript-eslint.io/rules/no-unnecessary-type-assertion/)
- [`@typescript-eslint/no-unsafe-type-assertion`](https://typescript-eslint.io/rules/no-unsafe-type-assertion/)
- [`@typescript-eslint/consistent-type-assertions`](https://typescript-eslint.io/rules/consistent-type-assertions/)

## Allowed Exceptions

- XML and OOXML input parsing may narrow from `unknown`/record-shaped values at
  the boundary, but new code should use shared parser helpers or
  purpose-specific `unsafe*Assertion` helpers before adding another inline
  assertion.
- Branded numeric/string values must be created through local constructor
  helpers. Direct brand assertions are allowed inside those constructors only.
- Tests may use boundary helpers for fixture narrowing, but repeated patterns
  should become fixture builders or assertion helper functions.
- External library type gaps may use assertions inside the adapter module that
  owns the integration.
- `as any`, new `as unknown as X`, and broad object literal assertions are not
  allowed in normal implementation code.
