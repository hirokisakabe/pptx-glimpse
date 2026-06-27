# Type Assertion Policy

This repository allows type assertions only when the type system cannot model a
boundary without a local helper. Normal implementation code should prefer
control-flow narrowing, discriminated unions, typed parser helpers, or branded
constructors.

## Current Inventory

Audit date: 2026-06-27.

Command used for the broad inventory:

```bash
node -e '/* TypeScript AST walk over packages/, vrt/, scripts/, bench/, e2e/ */'
```

The current source tree contains 751 `as` / angle-bracket type assertions in
linted TypeScript sources. The main buckets are:

| Category | Count | Policy |
| --- | ---: | --- |
| XML / external input boundary (`XmlNode`, `XmlOrderedNode`, parsed JSON) | 443 | Allowed at parser boundaries. Prefer shared XML helpers such as `getNode`, `getNodeArray`, `getAttr`, and enum parsers when touching nearby code. |
| Test fixture / mock construction | 95 | Allowed when constructing narrow fixtures. Prefer factory helpers when the same assertion repeats. |
| `as const` literal preservation | 54 | Allowed. Prefer `satisfies` when shape validation is also needed. |
| Branded unit / handle constructors (`Emu`, `Pt`, `PartPath`, etc.) | 12 | Allowed only inside constructor helpers such as `asEmu`, `asPt`, source unit helpers, and handle helpers. Callers should use the helper, not assert the brand directly. |
| Double assertion (`as unknown as X`) | 8 | Avoid in new implementation code. Existing uses are limited to external library gaps and test fixtures; replace with focused adapter helpers when edited. |
| Object literal assertion | 3 before this policy change | New direct object/array literal assertions are banned by ESLint. Use typed variables, `satisfies`, or fixture factories instead. |
| Other narrow assertions | 132 | Treat case-by-case. Enum/string-literal narrowing should move toward parser helpers such as `parseEnum`. External library gaps should stay inside adapter modules. |

No `as any` assertions were found in linted TypeScript sources during this
audit.

## ESLint Policy

The CI lint path (`npm run lint`) now explicitly runs these type assertion
rules:

- `@typescript-eslint/no-unnecessary-type-assertion`: `error`
- `@typescript-eslint/consistent-type-assertions`: `error`, with direct
  object/array literal assertions disallowed
- `@typescript-eslint/no-unsafe-type-assertion`: `warn`

`no-unsafe-type-assertion` is intentionally a warning for now. A trial run on
2026-06-27 reported 676 warnings, mostly from XML parser boundaries and
fixture narrowing. Turning it into an error should be done package-by-package
after the repeated XML access patterns have been moved behind helpers.

Rule references:

- [`@typescript-eslint/no-unnecessary-type-assertion`](https://typescript-eslint.io/rules/no-unnecessary-type-assertion/)
- [`@typescript-eslint/no-unsafe-type-assertion`](https://typescript-eslint.io/rules/no-unsafe-type-assertion/)
- [`@typescript-eslint/consistent-type-assertions`](https://typescript-eslint.io/rules/consistent-type-assertions/)

## Allowed Exceptions

- XML and OOXML input parsing may assert from `unknown`/record-shaped values at
  the boundary, but new code should use the shared parser helpers before adding
  another inline assertion.
- Branded numeric/string values must be created through local constructor
  helpers. Direct brand assertions are allowed inside those constructors only.
- Tests may use assertions for fixture narrowing, but repeated patterns should
  become fixture builders or assertion helper functions.
- External library type gaps may use assertions inside the adapter module that
  owns the integration.
- `as any`, new `as unknown as X`, and broad object literal assertions are not
  allowed in normal implementation code.
