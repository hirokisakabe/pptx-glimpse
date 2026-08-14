import {
  addConnector,
  type AddConnectorInput,
  addTextBox,
  type AddTextBoxInput,
  deleteShape,
  type EditableShapeFill,
  type EditableShapeOutline,
  findShapeNodeBySourceHandle,
  groupShapes,
  moveShapes,
  moveShapesAcrossSlides,
  type Emu,
  type PptxSourceModel,
  setShapeFill,
  setShapeOutline,
  type SourceChart,
  type SourceConnector,
  type SourceGroup,
  type SourceHandle,
  type SourceImage,
  type SourceShape,
  type SourceShapeNode,
  type SourceTable,
  type SourceTransform,
  ungroupShape,
  updateShapeTransform,
} from "@pptx-glimpse/document";

import { attemptCommand, type ApplyCommandAttempt } from "../command-contract.js";

/** Move one shape without resizing it. @inline */
export interface MoveShapeCommand {
  readonly kind: "moveShape";
  readonly handle: SourceHandle;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
}

/** Resize one shape without moving its origin. @inline */
export interface ResizeShapeCommand {
  readonly kind: "resizeShape";
  readonly handle: SourceHandle;
  readonly width: Emu;
  readonly height: Emu;
}

/** Replace a shape's position and size together. @inline */
export interface SetShapeTransformCommand {
  readonly kind: "setShapeTransform";
  readonly handle: SourceHandle;
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
}

/** Replace a shape's fill properties. @inline */
export interface SetShapeFillCommand {
  readonly kind: "setShapeFill";
  readonly handle: SourceHandle;
  readonly fill: EditableShapeFill;
}

/** Update the specified outline properties while preserving omitted properties. @inline */
export interface SetShapeOutlineCommand {
  readonly kind: "setShapeOutline";
  readonly handle: SourceHandle;
  readonly outline: EditableShapeOutline;
}

/** Add a text box to one slide. @inline */
export interface AddTextBoxCommand extends AddTextBoxInput {
  readonly kind: "addTextBox";
  readonly slideHandle: SourceHandle;
}

/** Add a connector to one slide. @inline */
export interface AddConnectorCommand extends AddConnectorInput {
  readonly kind: "addConnector";
  readonly slideHandle: SourceHandle;
}

/** Delete one supported drawing at the slide root or inside a native group. @inline */
export interface DeleteShapeCommand {
  readonly kind: "deleteShape";
  readonly handle: SourceHandle;
}

/** Group two or more consecutive sibling drawings into one native DrawingML group. @inline */
export interface GroupShapesCommand {
  readonly kind: "groupShapes";
  readonly shapeHandles: readonly SourceHandle[];
}

/** Move consecutive sibling drawings to a root/native-group destination in the same part. @inline */
export interface MoveShapesCommand {
  readonly kind: "moveShapes";
  readonly shapeHandles: readonly SourceHandle[];
  readonly destinationHandle: SourceHandle;
  readonly beforeShapeHandle?: SourceHandle;
}

/** Move consecutive slide-root typed drawings to another slide root. @inline */
export interface MoveShapesAcrossSlidesCommand {
  readonly kind: "moveShapesAcrossSlides";
  readonly shapeHandles: readonly SourceHandle[];
  readonly destinationSlideHandle: SourceHandle;
  readonly beforeShapeHandle?: SourceHandle;
}

/** Expand one losslessly ungroupable native DrawingML group. @inline */
export interface UngroupShapeCommand {
  readonly kind: "ungroupShape";
  readonly groupHandle: SourceHandle;
}

/** Source drawing kinds supported by the typed group convenience method. */
export type GroupableSourceShape =
  | SourceShape
  | SourceConnector
  | SourceImage
  | SourceTable
  | SourceChart
  | SourceGroup;

export type DrawingEditorCommand =
  | MoveShapeCommand
  | ResizeShapeCommand
  | SetShapeTransformCommand
  | SetShapeFillCommand
  | SetShapeOutlineCommand
  | AddTextBoxCommand
  | AddConnectorCommand
  | DeleteShapeCommand
  | GroupShapesCommand
  | MoveShapesCommand
  | MoveShapesAcrossSlidesCommand
  | UngroupShapeCommand;

const EXPECTED_REJECTION_PREFIXES = [
  "moveShape:",
  "resizeShape:",
  "setShapeTransform:",
  "setShapeFill:",
  "setShapeOutline:",
  "addTextBox:",
  "addConnector:",
  "deleteShape:",
  "groupShapes:",
  "moveShapes:",
  "moveShapesAcrossSlides:",
  "ungroupShape:",
  "updateShapeTransform:",
] as const;

export function applyDrawingCommand(
  document: PptxSourceModel,
  command: DrawingEditorCommand,
): ApplyCommandAttempt {
  return attemptCommand(EXPECTED_REJECTION_PREFIXES, () =>
    executeDrawingCommand(document, command),
  );
}

function executeDrawingCommand(
  document: PptxSourceModel,
  command: DrawingEditorCommand,
): PptxSourceModel {
  switch (command.kind) {
    case "moveShape": {
      requireFiniteEmu(command.offsetX, "moveShape", "offsetX");
      requireFiniteEmu(command.offsetY, "moveShape", "offsetY");
      const current = requireEditableShapeTransform(document, command.handle, "moveShape");
      return updateShapeTransform(document, command.handle, {
        offsetX: command.offsetX,
        offsetY: command.offsetY,
        width: current.width,
        height: current.height,
      });
    }
    case "resizeShape": {
      requirePositiveFiniteEmu(command.width, "resizeShape", "width");
      requirePositiveFiniteEmu(command.height, "resizeShape", "height");
      const current = requireEditableShapeTransform(document, command.handle, "resizeShape");
      return updateShapeTransform(document, command.handle, {
        offsetX: current.offsetX,
        offsetY: current.offsetY,
        width: command.width,
        height: command.height,
      });
    }
    case "setShapeTransform":
      requireFiniteEmu(command.offsetX, "setShapeTransform", "offsetX");
      requireFiniteEmu(command.offsetY, "setShapeTransform", "offsetY");
      requirePositiveFiniteEmu(command.width, "setShapeTransform", "width");
      requirePositiveFiniteEmu(command.height, "setShapeTransform", "height");
      requireEditableShapeTransform(document, command.handle, "setShapeTransform");
      return updateShapeTransform(document, command.handle, {
        offsetX: command.offsetX,
        offsetY: command.offsetY,
        width: command.width,
        height: command.height,
      });
    case "setShapeFill":
      validateShapeFill(command.fill, "setShapeFill");
      return setShapeFill(document, command.handle, command.fill);
    case "setShapeOutline":
      validateShapeOutline(command.outline);
      return setShapeOutline(document, command.handle, command.outline);
    case "addTextBox":
      requireFiniteEmu(command.offsetX, "addTextBox", "offsetX");
      requireFiniteEmu(command.offsetY, "addTextBox", "offsetY");
      requirePositiveFiniteEmu(command.width, "addTextBox", "width");
      requirePositiveFiniteEmu(command.height, "addTextBox", "height");
      if (command.text !== undefined && typeof command.text !== "string") {
        throw new Error("addTextBox: text must be a string");
      }
      if (command.name !== undefined && command.name.trim() === "") {
        throw new Error("addTextBox: name must be a non-empty string when provided");
      }
      return addTextBox(document, command.slideHandle, command);
    case "addConnector":
      requireFiniteEmu(command.offsetX, "addConnector", "offsetX");
      requireFiniteEmu(command.offsetY, "addConnector", "offsetY");
      requirePositiveFiniteEmu(command.width, "addConnector", "width");
      requirePositiveFiniteEmu(command.height, "addConnector", "height");
      return addConnector(document, command.slideHandle, command);
    case "deleteShape":
      return deleteShape(document, command.handle);
    case "groupShapes":
      for (const handle of command.shapeHandles) {
        const shape = findCommandShape(document, handle, "groupShapes");
        if (shape !== undefined && !isGroupableSourceShape(shape)) {
          throw new Error(`groupShapes: shape kind '${shape.kind}' is not supported`);
        }
      }
      return groupShapes(document, command.shapeHandles);
    case "moveShapes":
      return moveShapes(document, command.shapeHandles, command.destinationHandle, {
        ...(command.beforeShapeHandle !== undefined
          ? { beforeShapeHandle: command.beforeShapeHandle }
          : {}),
      });
    case "moveShapesAcrossSlides":
      return moveShapesAcrossSlides(
        document,
        command.shapeHandles,
        command.destinationSlideHandle,
        {
          ...(command.beforeShapeHandle !== undefined
            ? { beforeShapeHandle: command.beforeShapeHandle }
            : {}),
        },
      ).document;
    case "ungroupShape": {
      const group = findCommandShape(document, command.groupHandle, "ungroupShape");
      if (group?.kind === "group" && group.children.length === 0) {
        throw new Error("ungroupShape: group must contain at least one child");
      }
      return ungroupShape(document, command.groupHandle);
    }
  }
}

function findCommandShape(
  document: PptxSourceModel,
  handle: SourceHandle,
  operation: "groupShapes" | "ungroupShape",
): SourceShapeNode | undefined {
  try {
    return findShapeNodeBySourceHandle(document, handle);
  } catch (cause) {
    if (cause instanceof Error && cause.message.startsWith("findShapeNodeBySourceHandle:")) {
      throw new Error(`${operation}: ${cause.message}`, { cause });
    }
    throw cause;
  }
}

function isGroupableSourceShape(value: SourceShapeNode): value is GroupableSourceShape {
  return (
    value.kind === "shape" ||
    value.kind === "connector" ||
    value.kind === "group" ||
    value.kind === "image" ||
    value.kind === "table" ||
    value.kind === "chart"
  );
}

function validateShapeOutline(outline: EditableShapeOutline): void {
  if (outline.width === undefined && outline.fill === undefined) {
    throw new Error("setShapeOutline: outline must set width or fill");
  }
  if (outline.width !== undefined) {
    requirePositiveFiniteEmu(outline.width, "setShapeOutline", "width");
  }
  if (outline.fill !== undefined) validateShapeFill(outline.fill, "setShapeOutline");
}

function validateShapeFill(
  fill: EditableShapeFill,
  commandName: "setShapeFill" | "setShapeOutline",
): void {
  if (fill.kind === "none") return;
  if (fill.kind !== "solid") {
    throw new Error(`${commandName}: only solid and none fills are supported`);
  }
  if (fill.color.kind !== "srgb") {
    throw new Error(`${commandName}: only srgb solid fill colors are supported`);
  }
  if (!/^[0-9A-Fa-f]{6}$/.test(fill.color.hex)) {
    throw new Error(`${commandName}: color.hex must be a 6-digit hex value`);
  }
}

function requireEditableShapeTransform(
  document: PptxSourceModel,
  handle: SourceHandle,
  commandName: "moveShape" | "resizeShape" | "setShapeTransform",
): SourceTransform {
  const shape = findShapeNodeBySourceHandle(document, handle);
  if (shape === undefined) {
    throw new Error(`${commandName}: shape handle was not found in PptxSourceModel source`);
  }
  if (shape.kind === "raw" || shape.transform === undefined) {
    throw new Error(`${commandName}: shape handle does not reference a shape with xfrm`);
  }
  return shape.transform;
}

function requireFiniteEmu(value: Emu, commandName: string, fieldName: string): void {
  if (!Number.isFinite(value)) {
    throw new Error(`${commandName}: ${fieldName} must be a finite EMU value`);
  }
}

function requirePositiveFiniteEmu(value: Emu, commandName: string, fieldName: string): void {
  if (!Number.isFinite(value) || value <= 0) {
    throw new Error(`${commandName}: ${fieldName} must be a finite positive EMU value`);
  }
}
