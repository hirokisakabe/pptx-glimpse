"use client";

import { useMemo, useSyncExternalStore } from "react";
import type { PptxEditorSession } from "pptx-glimpse";

import { EditorController, type EditorControllerState } from "./editor-controller";

export type { PreferredSlideIndex } from "./editor-controller";

export function useEditorController(editor: PptxEditorSession): EditorControllerState & {
  readonly controller: EditorController<PptxEditorSession>;
} {
  const controller = useMemo(() => new EditorController(editor), [editor]);
  const state = useSyncExternalStore(
    controller.subscribe,
    controller.getSnapshot,
    controller.getSnapshot,
  );
  return useMemo(() => ({ ...state, controller }), [controller, state]);
}
