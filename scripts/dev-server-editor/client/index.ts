import { DEV_EDITOR_BOOTSTRAP_SCRIPT } from "./bootstrap.js";
import {
  createDevEditorImageReplacementScript,
  DEV_EDITOR_PANEL_COMMANDS_SCRIPT,
  DEV_EDITOR_SHAPE_OPTIONS_SCRIPT,
} from "./editor-panel.js";
import { DEV_EDITOR_HTTP_SCRIPT } from "./http.js";
import { DEV_EDITOR_SELECTION_SCRIPT, DEV_EDITOR_TRANSFORM_SCRIPT } from "./selection-transform.js";
import { DEV_EDITOR_SLIDE_OPERATIONS_SCRIPT } from "./slide-operations.js";
import {
  createDevEditorInitialStateScript,
  DEV_EDITOR_CORE_SCRIPT,
  DEV_EDITOR_RESPONSE_SCRIPT,
} from "./state.js";
import { DEV_EDITOR_TEXT_EDITING_SCRIPT } from "./text-editing.js";

interface DevEditorClientScriptOptions {
  readonly slides: readonly unknown[];
  readonly slideCount: number;
  readonly emuPerPixel: number;
  readonly maxImageReplacementBytes: number;
}

export function createDevEditorClientScript(options: DevEditorClientScriptOptions): string {
  return [
    createDevEditorInitialStateScript(options),
    DEV_EDITOR_CORE_SCRIPT,
    DEV_EDITOR_SHAPE_OPTIONS_SCRIPT,
    DEV_EDITOR_RESPONSE_SCRIPT,
    DEV_EDITOR_SELECTION_SCRIPT,
    DEV_EDITOR_TEXT_EDITING_SCRIPT,
    DEV_EDITOR_TRANSFORM_SCRIPT,
    DEV_EDITOR_PANEL_COMMANDS_SCRIPT,
    DEV_EDITOR_SLIDE_OPERATIONS_SCRIPT,
    createDevEditorImageReplacementScript(options),
    DEV_EDITOR_HTTP_SCRIPT,
    DEV_EDITOR_BOOTSTRAP_SCRIPT,
  ].join("");
}
