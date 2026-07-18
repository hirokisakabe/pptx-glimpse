import type { XmlNode } from "../reader/xml.js";
import { editDirtyPartPath } from "../source/edit-descriptors.js";
import type { PartPath, PptxSourceModel, PptxSourceModelEdit } from "../source/index.js";
import {
  applyShapeFillEdit,
  applyShapeOutlineEdit,
  applyShapeTransformEdit,
} from "./shape-property-edits.js";
import {
  applyAddChartEdit,
  applyAddConnectorEdit,
  applyAddPictureEdit,
  applyAddShapeEdit,
  applyAddTableEdit,
  applyAddTextBoxEdit,
  applyDeleteShapeEdit,
  applySetSlideBackgroundEdit,
} from "./shape-tree-edits.js";
import {
  applyParagraphPropertiesEdit,
  applyParagraphTextEdit,
  applyTextRunEdit,
  applyTextRunPropertiesEdit,
} from "./text-paragraph-edits.js";
import {
  buildXmlPreservingChildOrder,
  encodeXml,
  parseXmlForEditing,
  textDecoder,
  XML_DECLARATION,
} from "./xml-serialization.js";

export function serializeDirtyXmlPart(
  source: PptxSourceModel,
  partPath: PartPath,
  edits: readonly PptxSourceModelEdit[],
): Uint8Array {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === partPath);
  if (rawPart === undefined) {
    throw new Error(`writePptx: dirty part '${partPath}' has no preserved raw package material`);
  }
  if (rawPart.kind !== "binary") {
    throw new Error(`writePptx: dirty XML tree part '${partPath}' patching is not implemented`);
  }

  const root = parseXmlForEditing(textDecoder.decode(rawPart.bytes));
  // Edits are applied in chronological order. This relies on the editing API
  // invariant that deleteShape drops earlier edits targeting the deleted shape,
  // so no stale-target edit can follow its delete within one part. Hand-built
  // edit arrays that violate the invariant fail fast in the apply functions.
  for (const edit of edits) {
    if (editDirtyPartPath(edit) !== partPath) continue;
    applyDirtyPartEdit(root, edit);
  }
  return encodeXml(XML_DECLARATION + buildXmlPreservingChildOrder(root));
}

/**
 * The writer-side apply switch: the one place that maps an edit kind to its XML
 * patch. Kinds that never dirty an XML part (see `editDirtyPartPath`) throw so
 * a descriptor/apply mismatch fails fast instead of being silently skipped.
 */
function applyDirtyPartEdit(root: XmlNode, edit: PptxSourceModelEdit): void {
  switch (edit.kind) {
    case "replaceTextRunPlainText":
      applyTextRunEdit(root, edit);
      return;
    case "updateTextRunProperties":
      applyTextRunPropertiesEdit(root, edit);
      return;
    case "updateParagraphProperties":
      applyParagraphPropertiesEdit(root, edit);
      return;
    case "replaceParagraphPlainText":
      applyParagraphTextEdit(root, edit);
      return;
    case "updateShapeTransform":
      applyShapeTransformEdit(root, edit);
      return;
    case "updateShapeFill":
      applyShapeFillEdit(root, edit);
      return;
    case "updateShapeOutline":
      applyShapeOutlineEdit(root, edit);
      return;
    case "addTextBox":
      applyAddTextBoxEdit(root, edit);
      return;
    case "addShape":
      applyAddShapeEdit(root, edit);
      return;
    case "addTable":
      applyAddTableEdit(root, edit);
      return;
    case "addConnector":
      applyAddConnectorEdit(root, edit);
      return;
    case "addPicture":
      applyAddPictureEdit(root, edit);
      return;
    case "addChart":
      applyAddChartEdit(root, edit);
      return;
    case "deleteShape":
      applyDeleteShapeEdit(root, edit);
      return;
    case "setSlideBackground":
      applySetSlideBackgroundEdit(root, edit);
      return;
    case "replaceImage":
    case "addEmptySlideFromLayout":
    case "duplicateSlide":
    case "moveSlide":
    case "deleteSlide":
      throw new Error(`writePptx: edit kind '${edit.kind}' does not patch a dirty XML part`);
  }
}
