import {
  type PptxSourceModel,
  type SourceHandle,
  updateThemeScheme,
  type UpdateThemeSchemeInput,
} from "@pptx-glimpse/document";

import { type ApplyCommandAttempt, attemptCommand } from "../command-contract.js";

/** Update selected color/font fields on one existing theme. @inline */
export interface UpdateThemeSchemeCommand extends UpdateThemeSchemeInput {
  readonly kind: "updateThemeScheme";
  readonly handle: SourceHandle;
}

export type ThemeEditorCommand = UpdateThemeSchemeCommand;

export function applyThemeCommand(
  document: PptxSourceModel,
  command: ThemeEditorCommand,
): ApplyCommandAttempt {
  return attemptCommand(["updateThemeScheme:"], () =>
    updateThemeScheme(document, command.handle, {
      ...(command.colorScheme !== undefined ? { colorScheme: command.colorScheme } : {}),
      ...(command.fontScheme !== undefined ? { fontScheme: command.fontScheme } : {}),
    }),
  );
}
