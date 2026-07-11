/**
 * Public editing surface for PptxSourceModel.
 *
 * The operation implementations live in responsibility-focused modules so text edits,
 * shape edits, picture authoring, slide topology edits, and image replacement can
 * evolve independently while preserving the historical ./editing.js import path.
 */

export type {
  AddChartAxisInput,
  AddChartInput,
  AddChartPlotLayoutInput,
  AddChartSeriesInput,
  NativeChartLegendPosition,
  NativeChartType,
  NativeRadarStyle,
} from "./chart-authoring.js";
export { addChart } from "./chart-authoring.js";
export { replaceImageBytes } from "./image-replacement.js";
export type { AddPictureCropInput, AddPictureInput } from "./picture-authoring.js";
export { addPicture } from "./picture-authoring.js";
export type {
  AddConnectorConnectionEndpointInput,
  AddConnectorInput,
  AddConnectorOutlineInput,
  AddShapeBodyPropertiesInput,
  AddShapeColorInput,
  AddShapeEffectsInput,
  AddShapeFillInput,
  AddShapeGlowInput,
  AddShapeGradientFillInput,
  AddShapeInput,
  AddShapeOutlineInput,
  AddShapeParagraphInput,
  AddShapeParagraphPropertiesInput,
  AddShapeRunInput,
  AddShapeRunPropertiesInput,
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
  addShape,
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
export type {
  AddTableBorderInput,
  AddTableCellInput,
  AddTableInput,
  AddTableRowInput,
  AddTableRunInput,
  AddTableRunPropertiesInput,
} from "./table-authoring.js";
export { addTable } from "./table-authoring.js";
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
