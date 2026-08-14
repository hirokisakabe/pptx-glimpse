import type { EditorCommand, PptxEditorSession } from "pptx-glimpse";

import type { PreferredSlideIndex } from "./pptx-editor-controller.js";

type EditorSession = PptxEditorSession;

export type RunEditorOperation = (
  operation: (session: EditorSession) => Promise<string | void> | string | void,
  success: string,
  preferredIndex?: PreferredSlideIndex<EditorSession>,
  historyAction?: "mutation" | "undo" | "redo",
) => Promise<boolean>;

export type ApplyEditorCommand = (command: EditorCommand, success: string) => Promise<boolean>;
