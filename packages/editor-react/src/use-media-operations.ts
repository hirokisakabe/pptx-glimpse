import type { PptxEditorSession, PptxEditorShapeInfo } from "pptx-glimpse";
import { useCallback } from "react";

import type { ApplyEditorCommand } from "./editor-interaction-types.js";
import type { PptxEditorController } from "./pptx-editor-controller.js";

const MAX_IMAGE_REPLACEMENT_BYTES = 5 * 1024 * 1024;

interface UseMediaOperationsOptions {
  readonly applyCommand: ApplyEditorCommand;
  readonly commitPendingEdits: () => Promise<boolean>;
  readonly controller: PptxEditorController<PptxEditorSession>;
  readonly selectedShape: PptxEditorShapeInfo | null;
}

export function useMediaOperations({
  applyCommand,
  commitPendingEdits,
  controller,
  selectedShape,
}: UseMediaOperationsOptions) {
  const replaceImage = useCallback(
    async (file: File) => {
      if (!(await commitPendingEdits()) || selectedShape?.handle === undefined) return;
      const replacement = selectedShape.editableImageReplacement;
      if (replacement === undefined) return;
      if (file.size > MAX_IMAGE_REPLACEMENT_BYTES) {
        controller.setError("Replacement image must be 5 MB or smaller.");
        return;
      }
      if (file.type !== "" && file.type !== replacement.contentType) {
        controller.setError(`Replacement image must use ${replacement.contentType}.`);
        return;
      }
      const bytes = new Uint8Array(await file.arrayBuffer());
      const detectedContentType = detectImageContentType(bytes);
      if (detectedContentType !== replacement.contentType) {
        controller.setError(`Replacement image must use ${replacement.contentType}.`);
        return;
      }
      await applyCommand(
        { kind: "replaceImage", handle: selectedShape.handle, bytes },
        "Image replaced",
      );
    },
    [applyCommand, commitPendingEdits, controller, selectedShape],
  );

  return { replaceImage };
}

function detectImageContentType(bytes: Uint8Array): string | undefined {
  if (startsWithBytes(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
    return "image/png";
  }
  if (startsWithBytes(bytes, [0xff, 0xd8, 0xff])) return "image/jpeg";
  if (
    startsWithBytes(bytes, [0x47, 0x49, 0x46, 0x38, 0x37, 0x61]) ||
    startsWithBytes(bytes, [0x47, 0x49, 0x46, 0x38, 0x39, 0x61])
  ) {
    return "image/gif";
  }
  if (startsWithBytes(bytes, [0x42, 0x4d])) return "image/bmp";
  if (
    startsWithBytes(bytes, [0x49, 0x49, 0x2a, 0x00]) ||
    startsWithBytes(bytes, [0x4d, 0x4d, 0x00, 0x2a])
  ) {
    return "image/tiff";
  }
  if (
    startsWithBytes(bytes, [0x52, 0x49, 0x46, 0x46]) &&
    bytes[8] === 0x57 &&
    bytes[9] === 0x45 &&
    bytes[10] === 0x42 &&
    bytes[11] === 0x50
  ) {
    return "image/webp";
  }
  return undefined;
}

function startsWithBytes(bytes: Uint8Array, prefix: readonly number[]): boolean {
  if (bytes.length < prefix.length) return false;
  return prefix.every((value, index) => bytes[index] === value);
}
