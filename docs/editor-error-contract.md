# Editor Error Contract

## Audience and purpose

This document is the source of truth for maintainers who implement or extend editing
operations across `@pptx-glimpse/document`, `@pptx-glimpse/editor`, `pptx-glimpse`, and the
demo/UI. It defines ownership, classification, conversion, catching, and state guarantees.

Package READMEs and generated API documentation are for consumers: they show how to inspect a
result, catch a typed error, and process warnings. They must agree with this document, but they
must not become the repository's design specification.

## Layer ownership

- `@pptx-glimpse/document` owns OOXML parsing, source-model edits, and PPTX writing. Its edit
  functions may reject unsupported or invalid document operations by throwing at their direct
  API boundary; it does not depend on editor error types.
- `@pptx-glimpse/editor` owns validated commands, selection, history, warnings, and
  `EditorOperationErrorCode`. It converts expected command-operation rejection into
  `EditorOperationFailure` and otherwise remains independent of parsing, rendering, and writing.
- `pptx-glimpse` owns the integrated asynchronous session and `PptxEditorError`. It unwraps
  headless failures and owns the `read-failed`, `render-failed`, and `write-failed` integration
  codes.
- Demo/UI code consumes public results and errors. It may map codes to UI text or telemetry, but
  it must not redefine codes or infer them by parsing messages.

Dependencies continue to point upward from document to editor to core to demo/UI. A lower layer
must never import a higher layer's error class.

## Failure classification and transport

| Classification                       | Examples                                                                         | Transport                                                                                          |
| ------------------------------------ | -------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------- |
| Expected operation rejection         | Invalid command input or target, missing selection target, empty history stack   | Headless `{ ok: false, code, message, cause? }`; high-level `PptxEditorError` with the same fields |
| Integration/runtime failure          | PPTX read, SVG render, PPTX write or round-trip validation failure               | High-level `PptxEditorError` with `read-failed`, `render-failed`, or `write-failed`                |
| Warning                              | Replacing a shared media part                                                    | Successful result/response `warnings`; never thrown and never added to an error-code union         |
| Programmer error/invariant violation | Unknown command discriminant, impossible internal state, broken result invariant | Propagate unchanged; do not turn it into an expected result or integration code                    |

An operation rejection is expected when the caller can reasonably branch on it and continue
using the session. An integration failure is owned by an asynchronous or serialization boundary.
A warning describes a successfully committed operation. A programmer error indicates that the
implementation or an untyped caller violated an invariant rather than requesting a supported but
invalid operation.

## Conversion flow

```text
document edit rejection
  -> editor command boundary
  -> EditorOperationFailure { ok: false, code, message, cause? }
  -> core unwrapEditorOperation()
  -> PptxEditorError { name, code, message, cause, stack }
  -> application catch / isPptxEditorError()
```

The high-level unwrap is field preserving: it must not replace the headless `code`, rewrite the
`message`, or discard `cause`. Convenience methods use the same unwrap helper or construct the
matching typed operation error directly; they must not duplicate plain-`Error` conversion.

## Catch and wrap rules

Catch only at a boundary that owns a classification:

- The editor may catch a direct command execution rejection and the documented conflicting-edit
  validation. Warning collection, normalization, selection reconciliation, and unrelated
  internal logic stay outside that catch.
- Core may catch `readPptx()` in session creation, the configured renderer invocation in
  `renderCurrentSlides()`, and `writePptx()` plus output reread validation in `save()`.
- Core unwraps an explicit headless failure. It does not catch arbitrary exceptions around an
  entire public method and guess a code.

Do not add a catch around the whole command loop, render response mapping, or session method.
Such a catch would hide programmer errors and make the machine-readable code misleading. When a
new integration is added, give its owning boundary a code before wrapping it.

## Field preservation

- `code` is the stable machine-readable discriminator. Applications branch on it.
- `message` is human-readable context. Headless-to-high-level conversion preserves it exactly.
  Integration wrappers prepend boundary context and retain the original message when available.
- `cause` stores the original thrown value, including non-`Error` values. Omit it only when an
  expected rejection has no underlying thrown value.
- `name` is always `PptxEditorError` for the high-level typed class.
- `stack` is created by the standard `Error` constructor at the wrapping boundary. The original
  stack remains reachable through `cause` when the cause is an `Error`.

`PptxEditorError` remains an `Error`, and `isPptxEditorError(value)` is the supported TypeScript
guard.

## Atomicity and state

Headless expected rejection is atomic:

- `document` remains the identical model object;
- `selection` remains unchanged;
- undo and redo stacks retain their contents and depths.

`applyAll()` builds a candidate document and commits it only after every command and
cross-command validation succeeds. Failed selection and empty undo/redo checks inspect state
before mutation.

The high-level session inherits those guarantees for operation failures because it unwraps the
headless result after the headless boundary returns. A render failure occurs after a successful
headless command/history transition; it does not roll that committed transition back. The
document and history therefore describe the successful edit while the last rendered `slides`
remain unchanged. Callers may retry `renderCurrentSlides()` or recreate the session. Read failure
creates no session. Write failure does not mutate editor state.

## Code and test matrix

| Code                | Defined/created by                                                  | High-level destination      | Required evidence                                                                    |
| ------------------- | ------------------------------------------------------------------- | --------------------------- | ------------------------------------------------------------------------------------ |
| `invalid-command`   | `packages/editor/src/index.ts`: command boundary                    | Same-code `PptxEditorError` | Headless apply/applyAll rejection and atomicity; high-level field/cause preservation |
| `invalid-selection` | `packages/editor/src/index.ts`: `selectShape()`                     | Same-code `PptxEditorError` | Selection atomicity; Node/browser same-code test                                     |
| `empty-undo-stack`  | `packages/editor/src/index.ts`: `undo()`                            | Same-code `PptxEditorError` | Empty-history shape/message and unchanged state                                      |
| `empty-redo-stack`  | `packages/editor/src/index.ts`: `redo()`                            | Same-code `PptxEditorError` | Empty-history shape/message and unchanged state                                      |
| `read-failed`       | `packages/core/src/pptx-editor-session.ts`: `create()`              | Thrown directly             | Invalid input and retained cause                                                     |
| `render-failed`     | `packages/core/src/pptx-editor-session.ts`: `renderCurrentSlides()` | Thrown directly             | Initial and post-operation renderer rejection with retained cause                    |
| `write-failed`      | `packages/core/src/pptx-editor-session.ts`: `save()`                | Thrown directly             | Writer/round-trip validation rejection with retained cause                           |

Warning coverage must separately verify that `shared-media-part` remains on a successful
headless result and high-level response.

## Adding an operation or code

1. Classify the condition using the four categories above. Do not start by choosing a catch.
2. Put the code in the lowest layer that owns the recoverable condition.
3. Add it to `EditorOperationErrorCode` only for expected headless operation rejection. Add an
   integration code to `PptxEditorErrorCode` only when core owns that boundary.
4. Return the common `EditorOperationFailure` shape from every relevant headless operation.
5. Route high-level conversion through `unwrapEditorOperation()`; do not add a plain `Error`.
6. Define document, selection, and history atomicity before implementation.
7. Add headless shape/atomicity/unexpected-propagation tests, high-level conversion/cause tests,
   and Node/browser export or behavior coverage as applicable.
8. Update this matrix, both package READMEs/API docs, public entry exports, and a changeset.

## Examples

Recommended command boundary:

```ts
const result = session.apply(command);
if (!result.ok) {
  return result; // code, message, and cause remain intact
}
```

Recommended high-level handling:

```ts
try {
  await editor.apply(command);
} catch (error) {
  if (isPptxEditorError(error)) {
    handleEditorCode(error.code, error.message);
    return;
  }
  throw error;
}
```

Do not flatten unexpected failures:

```ts
// Wrong: this hides invariant failures as caller-recoverable command rejection.
try {
  runCommandAndAllPostProcessing();
} catch (cause) {
  return { ok: false, code: "invalid-command", message: "Rejected", cause };
}
```
