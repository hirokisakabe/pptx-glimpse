import {
  clearPictureCrop,
  type PptxSourceModel,
  replaceImageBytes,
  setPictureCrop,
  type SetPictureCropInput,
  type SourceHandle,
} from "@pptx-glimpse/document";

import { attemptCommand, type ApplyCommandAttempt } from "../command-contract.js";

/** Replace the media bytes referenced by one image shape. @inline */
export interface ReplaceImageCommand {
  readonly kind: "replaceImage";
  readonly handle: SourceHandle;
  readonly bytes: Uint8Array;
}

/** Set crop insets on one stretch-filled picture. @inline */
export interface SetPictureCropCommand extends SetPictureCropInput {
  readonly kind: "setPictureCrop";
  readonly handle: SourceHandle;
}

/** Remove the crop rectangle from one stretch-filled picture. @inline */
export interface ClearPictureCropCommand {
  readonly kind: "clearPictureCrop";
  readonly handle: SourceHandle;
}

export type MediaEditorCommand =
  | ReplaceImageCommand
  | SetPictureCropCommand
  | ClearPictureCropCommand;

const EXPECTED_REJECTION_PREFIXES = [
  "replaceImageBytes:",
  "setPictureCrop:",
  "clearPictureCrop:",
] as const;

export function applyMediaCommand(
  document: PptxSourceModel,
  command: MediaEditorCommand,
): ApplyCommandAttempt {
  return attemptCommand(EXPECTED_REJECTION_PREFIXES, () => {
    switch (command.kind) {
      case "replaceImage":
        return replaceImageBytes(document, command.handle, command.bytes);
      case "setPictureCrop":
        return setPictureCrop(document, command.handle, command);
      case "clearPictureCrop":
        return clearPictureCrop(document, command.handle);
    }
  });
}
