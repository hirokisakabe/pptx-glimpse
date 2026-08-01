import { getChild, type XmlNode } from "../reader/xml.js";
import type { PptxSourceModelReplaceImageEdit } from "../source/index.js";
import { getDrawingPartRoot } from "./dirty-part-xml-helpers.js";
import { locateShapeTreeNode, parseShapeLocator } from "./xml-locators.js";

/** Patches only the selected picture's embedded relationship for copy-on-write. */
export function applyReplaceImageEdit(root: XmlNode, edit: PptxSourceModelReplaceImageEdit): void {
  if (edit.mode !== "copyOnWrite" || edit.replacementRelationshipId === undefined) {
    throw new Error("writePptx: in-place image replacement must not dirty a drawing part");
  }
  const locator = parseShapeLocator(edit.handle, "image replacement edit");
  const spTree = getChild(getChild(getDrawingPartRoot(root), "cSld"), "spTree");
  const picture = locateShapeTreeNode(spTree, locator);
  const blip = getChild(getChild(picture, "blipFill"), "blip");
  if (picture === undefined || blip === undefined) {
    throw new Error(
      `writePptx: image replacement handle '${locator.nodeId}' no longer matches p:pic XML with a:blip`,
    );
  }
  blip["@_r:embed"] = edit.replacementRelationshipId;
}
