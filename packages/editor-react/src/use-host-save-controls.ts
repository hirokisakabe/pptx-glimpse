import type { PptxEditorSession } from "pptx-glimpse";
import { useCallback, useMemo } from "react";

import type { PptxEditorController } from "./pptx-editor-controller.js";
import type { PptxEditorHostControls } from "./PptxEditor.js";

interface UseHostSaveControlsOptions {
  readonly busy: boolean;
  readonly commitPendingEdits: () => Promise<boolean>;
  readonly controller: PptxEditorController<PptxEditorSession>;
  readonly dirty: boolean;
  readonly directTextEditing: boolean;
  readonly message: string;
}

export function useHostSaveControls({
  busy,
  commitPendingEdits,
  controller,
  dirty,
  directTextEditing,
  message,
}: UseHostSaveControlsOptions): PptxEditorHostControls {
  const hasUnsavedChanges = useCallback(
    () => controller.getSnapshot().dirty || directTextEditing,
    [controller, directTextEditing],
  );
  const markSaved = useCallback(
    (savedHistory: ReturnType<PptxEditorSession["save"]>["history"], savedMessage: string) =>
      controller.markSaved(savedHistory, savedMessage),
    [controller],
  );
  const save = useCallback(async () => {
    if (!(await commitPendingEdits()) || controller.getSnapshot().busy) return undefined;
    return controller.save();
  }, [commitPendingEdits, controller]);
  const setError = useCallback((error: string) => controller.setError(error), [controller]);

  return useMemo(
    () => ({
      busy,
      dirty,
      message,
      resetScope: controller,
      commitPendingEdits,
      hasUnsavedChanges,
      markSaved,
      save,
      setError,
    }),
    [
      busy,
      commitPendingEdits,
      controller,
      dirty,
      hasUnsavedChanges,
      markSaved,
      message,
      save,
      setError,
    ],
  );
}
