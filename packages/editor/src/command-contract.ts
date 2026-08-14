import type { PptxSourceModel } from "@pptx-glimpse/document";

export type EditorOperationErrorCode =
  | "invalid-command"
  | "invalid-selection"
  | "empty-undo-stack"
  | "empty-redo-stack";

export interface EditorOperationFailure<
  Code extends EditorOperationErrorCode = EditorOperationErrorCode,
> {
  readonly ok: false;
  readonly code: Code;
  readonly message: string;
  readonly cause?: unknown;
}

export interface EditorCommandWarning {
  readonly code: "shared-media-part";
  readonly message: string;
  readonly mediaPartPath: string;
  readonly referenceCount: number;
}

export type EditorApplyCommandResult =
  | {
      readonly ok: true;
      readonly document: PptxSourceModel;
      readonly warnings?: readonly EditorCommandWarning[];
    }
  | EditorOperationFailure<"invalid-command">;

export type ApplyCommandAttempt =
  | {
      readonly ok: true;
      readonly document: PptxSourceModel;
    }
  | EditorOperationFailure<"invalid-command">;

export function attemptCommand(
  expectedRejectionPrefixes: readonly string[],
  operation: () => PptxSourceModel,
): ApplyCommandAttempt {
  try {
    return { ok: true, document: operation() };
  } catch (cause) {
    return invalidCommandFailure(expectedRejectionPrefixes, cause);
  }
}

export function invalidCommandFailure(
  expectedRejectionPrefixes: readonly string[],
  cause: unknown,
): EditorOperationFailure<"invalid-command"> {
  if (
    !(cause instanceof Error) ||
    !expectedRejectionPrefixes.some((prefix) => cause.message.startsWith(prefix))
  ) {
    throw cause;
  }
  return {
    ok: false,
    code: "invalid-command",
    message: cause.message,
    cause,
  };
}
