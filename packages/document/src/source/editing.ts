/**
 * Public editing surface for PptxSourceModel.
 *
 * The operation implementations live in responsibility-focused modules so text edits,
 * shape edits, picture authoring, slide topology edits, and image replacement can
 * evolve independently while preserving the historical ./editing.js import path.
 */

export { replaceImageBytes } from "./image-replacement.js";
export type { AddPictureCropInput, AddPictureInput } from "./picture-authoring.js";
export { addPicture } from "./picture-authoring.js";
export type {
  AddConnectorConnectionEndpointInput,
  AddConnectorInput,
  AddConnectorOutlineInput,
  AddTextBoxBaselineInput,
  AddTextBoxBodyPropertiesInput,
  AddTextBoxColorInput,
  AddTextBoxGlowInput,
  AddTextBoxGradientFillInput,
  AddTextBoxInput,
  AddTextBoxOutlineInput,
  AddTextBoxParagraphInput,
  AddTextBoxParagraphPropertiesInput,
  AddTextBoxRunInput,
  AddTextBoxRunPropertiesInput,
  AddTextBoxUnderlineInput,
  AddTextBoxUnderlineStyle,
  UpdateShapeTransformInput,
} from "./shape-editing.js";
export {
  addConnector,
  addTextBox,
  deleteShape,
  findShapeNodeBySourceHandle,
  setShapeFill,
  setShapeOutline,
  updateShapeTransform,
} from "./shape-editing.js";
export type { AddEmptySlideFromLayoutInput } from "./slide-topology.js";
export type { MoveSlideInput } from "./slide-topology.js";
export {
  addEmptySlideFromLayout,
  deleteSlide,
  duplicateSlide,
  moveSlide,
} from "./slide-topology.js";
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
