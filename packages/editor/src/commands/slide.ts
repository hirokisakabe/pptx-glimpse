import {
  addEmptySlideFromLayout,
  type AddEmptySlideFromLayoutInput,
  deleteSlide,
  duplicateSlide,
  moveSlide,
  type MoveSlideInput,
  type PptxSourceModel,
  type SourceHandle,
} from "@pptx-glimpse/document";

import { type ApplyCommandAttempt,attemptCommand } from "../command-contract.js";

/** Add an empty slide based on an existing layout. @inline */
export interface AddEmptySlideFromLayoutCommand extends AddEmptySlideFromLayoutInput {
  readonly kind: "addEmptySlideFromLayout";
}

/** Duplicate one slide immediately after its source. @inline */
export interface DuplicateSlideCommand {
  readonly kind: "duplicateSlide";
  readonly handle: SourceHandle;
}

/** Move one slide to a new zero-based array position. @inline */
export interface MoveSlideCommand extends MoveSlideInput {
  readonly kind: "moveSlide";
  readonly handle: SourceHandle;
}

/** Delete one slide. @inline */
export interface DeleteSlideCommand {
  readonly kind: "deleteSlide";
  readonly handle: SourceHandle;
}

export type SlideEditorCommand =
  | AddEmptySlideFromLayoutCommand
  | DuplicateSlideCommand
  | MoveSlideCommand
  | DeleteSlideCommand;

const EXPECTED_REJECTION_PREFIXES = [
  "addEmptySlideFromLayout:",
  "duplicateSlide:",
  "moveSlide:",
  "deleteSlide:",
] as const;

export function applySlideCommand(
  document: PptxSourceModel,
  command: SlideEditorCommand,
): ApplyCommandAttempt {
  return attemptCommand(EXPECTED_REJECTION_PREFIXES, () => {
    switch (command.kind) {
      case "addEmptySlideFromLayout":
        return addEmptySlideFromLayout(document, command);
      case "duplicateSlide":
        return duplicateSlide(document, command.handle);
      case "moveSlide":
        return moveSlide(document, command.handle, command);
      case "deleteSlide":
        return deleteSlide(document, command.handle);
    }
  });
}
