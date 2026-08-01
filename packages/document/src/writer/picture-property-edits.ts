import { getChild, localName, type XmlNode } from "../reader/xml.js";
import type { PptxSourceModelPictureCropEdit } from "../source/index.js";
import { insertChildByOrder, qualifiedSiblingName } from "./dirty-part-xml-helpers.js";
import { locateShapeTreeNodeLocation, parseShapeLocator } from "./xml-locators.js";
import { deleteChild } from "./xml-node-utils.js";

const CROP_ATTRIBUTE_LOCAL_NAMES = new Set(["l", "t", "r", "b"]);

export function applyPictureCropEdit(root: XmlNode, edit: PptxSourceModelPictureCropEdit): void {
  const locator = parseShapeLocator(edit.handle, "picture crop edit");
  const spTree = getChild(getChild(getChild(root, "sld"), "cSld"), "spTree");
  const location = locateShapeTreeNodeLocation(spTree, locator);
  if (location === undefined || location.nodeKind !== "pic") {
    throw new Error(
      `writePptx: picture crop handle '${String(edit.handle.nodeId)}' does not reference p:pic`,
    );
  }
  const blipFill = getChild(location.node, "blipFill");
  if (blipFill === undefined) {
    throw new Error("writePptx: picture crop target has no p:blipFill");
  }
  if (countChildren(blipFill, "stretch") !== 1 || countChildren(blipFill, "tile") !== 0) {
    throw new Error(
      "writePptx: picture crop target must contain exactly one a:stretch and no a:tile",
    );
  }
  if (countChildren(blipFill, "srcRect") > 1) {
    throw new Error("writePptx: picture crop target contains multiple a:srcRect elements");
  }

  if (edit.crop === undefined) {
    deleteChild(blipFill, "srcRect");
    return;
  }

  const srcRect = getChild(blipFill, "srcRect");
  if (srcRect === undefined) {
    const stretchKey = Object.keys(blipFill).find(
      (key) => !key.startsWith("@_") && localName(key) === "stretch",
    );
    if (stretchKey === undefined) {
      throw new Error("writePptx: picture crop target has no qualified stretch element");
    }
    insertChildByOrder(
      blipFill,
      qualifiedSiblingName(stretchKey, "srcRect"),
      cropAttributes(edit.crop),
      (name) => name === "tile" || name === "stretch",
    );
    return;
  }
  for (const key of Object.keys(srcRect)) {
    if (isCropAttribute(key)) delete srcRect[key];
  }
  Object.assign(srcRect, cropAttributes(edit.crop));
}

function cropAttributes(crop: NonNullable<PptxSourceModelPictureCropEdit["crop"]>): XmlNode {
  return {
    ...(crop.left !== undefined ? { "@_l": String(crop.left) } : {}),
    ...(crop.top !== undefined ? { "@_t": String(crop.top) } : {}),
    ...(crop.right !== undefined ? { "@_r": String(crop.right) } : {}),
    ...(crop.bottom !== undefined ? { "@_b": String(crop.bottom) } : {}),
  };
}

function countChildren(parent: XmlNode, childName: string): number {
  return Object.entries(parent).reduce((count, [key, value]) => {
    if (key.startsWith("@_") || localName(key) !== childName) return count;
    return count + (Array.isArray(value) ? value.length : 1);
  }, 0);
}

function isCropAttribute(key: string): boolean {
  return key.startsWith("@_") && CROP_ATTRIBUTE_LOCAL_NAMES.has(key.slice(2));
}
