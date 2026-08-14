import type { PptxSourceModel } from "@pptx-glimpse/document";

import type { ApplyCommandAttempt } from "../command-contract.js";
import type { EditorCommand } from "../index.js";
import { applyChartCommand } from "./chart.js";
import { applyDrawingCommand } from "./drawing.js";
import { applyMediaCommand } from "./media.js";
import { applySlideCommand } from "./slide.js";
import { applyTextCommand } from "./text.js";
import { applyThemeCommand } from "./theme.js";

export type * from "./chart.js";
export type * from "./drawing.js";
export type * from "./media.js";
export type * from "./slide.js";
export type * from "./text.js";
export type * from "./theme.js";

/**
 * Runtime acceptance, domain dispatch, and expected-rejection classification share this one
 * exhaustive switch. The never assignment after the switch makes a missing union member fail
 * type checking while the terminal TypeError still rejects unknown kinds from untyped callers.
 */
export function applyCommandToDocument(
  document: PptxSourceModel,
  command: EditorCommand,
): ApplyCommandAttempt {
  const runtimeKind = command.kind;
  switch (command.kind) {
    case "replaceTextRunPlainText":
    case "replaceParagraphPlainText":
    case "setTextRunProperties":
    case "clearTextRunProperties":
    case "setParagraphProperties":
    case "clearParagraphProperties":
      return applyTextCommand(document, command);
    case "moveShape":
    case "resizeShape":
    case "setShapeTransform":
    case "setShapeFill":
    case "setShapeOutline":
    case "addTextBox":
    case "addConnector":
    case "deleteShape":
    case "groupShapes":
    case "moveShapes":
    case "moveShapesAcrossSlides":
    case "ungroupShape":
      return applyDrawingCommand(document, command);
    case "replaceImage":
    case "setPictureCrop":
    case "clearPictureCrop":
      return applyMediaCommand(document, command);
    case "updateChartData":
    case "updateScatterChartData":
    case "updateBubbleChartData":
      return applyChartCommand(document, command);
    case "addEmptySlideFromLayout":
    case "duplicateSlide":
    case "moveSlide":
    case "deleteSlide":
      return applySlideCommand(document, command);
    case "updateThemeScheme":
      return applyThemeCommand(document, command);
  }
  const exhaustiveCommand: never = command;
  void exhaustiveCommand;
  throw new TypeError(`EditorSession: unsupported command kind '${String(runtimeKind)}'`);
}
