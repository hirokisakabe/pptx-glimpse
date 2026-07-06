/**
 * Public editing surface for PptxSourceModel.
 *
 * The operation implementations live in responsibility-focused modules so text edits,
 * shape edits, slide topology edits, and image replacement can evolve independently
 * while preserving the historical ./editing.js import path.
 */

export { replaceImageBytes } from "./image-replacement.js";
export type {
  AddConnectorConnectionEndpointInput,
  AddConnectorInput,
  AddConnectorOutlineInput,
  AddTextBoxInput,
  UpdateShapeTransformInput,
} from "./shape-editing.js";
export {
  addConnector,
  addTextBox,
  deleteShape,
  findShapeNodeBySourceHandle,
  updateShapeTransform,
} from "./shape-editing.js";
export type { AddEmptySlideFromLayoutInput } from "./slide-topology.js";
export { addEmptySlideFromLayout, deleteSlide, duplicateSlide } from "./slide-topology.js";
export {
  clearParagraphProperties,
  clearTextRunProperties,
  findParagraphBySourceHandle,
  findTextRunBySourceHandle,
  replaceParagraphPlainText,
  replaceTextRunPlainText,
  setParagraphProperties,
  setTextRunProperties,
} from "./text-editing.js";
