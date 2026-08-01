import {
  getChild,
  getChildArray,
  getNamespacedAttr,
  localName,
  type XmlNode,
} from "../reader/xml.js";
import type { PptxSourceModelAddSlideLayoutEdit } from "../source/index.js";
import {
  namespacedAttributeKey,
  namespacedChildKey,
  replaceNodeEntries,
} from "./xml-node-utils.js";

export function applyAddSlideLayoutEdit(
  root: XmlNode,
  edit: PptxSourceModelAddSlideLayoutEdit,
): void {
  const master = getChild(root, "sldMaster");
  if (master === undefined) {
    throw new Error("writePptx: slide master part does not contain p:sldMaster root");
  }
  const layoutIdList = getChild(master, "sldLayoutIdLst") ?? createLayoutIdList(master);
  const items = getChildArray(layoutIdList, "sldLayoutId");
  if (
    items.some(
      (item) =>
        getNamespacedAttr(item, "id") === edit.newRelationshipId ||
        Number(item["@_id"]) === edit.newLayoutNumericId,
    )
  ) {
    throw new Error("writePptx: slide layout relationship or numeric ID already exists");
  }
  const itemKey = namespacedChildKey(layoutIdList, "p:sldLayoutId", "sldLayoutId");
  const relationshipAttributeKey =
    items[0] === undefined ? "@_r:id" : namespacedAttributeKey(items[0], "r:id", "id");
  layoutIdList[itemKey] = [
    ...items,
    {
      "@_id": String(edit.newLayoutNumericId),
      [relationshipAttributeKey]: edit.newRelationshipId,
    },
  ];
}

function createLayoutIdList(master: XmlNode): XmlNode {
  const key = namespacedChildKey(master, "p:sldLayoutIdLst", "sldLayoutIdLst");
  const created: XmlNode = {};
  const entries: [string, unknown][] = [];
  let inserted = false;
  for (const entry of Object.entries(master)) {
    if (
      !inserted &&
      !entry[0].startsWith("@_") &&
      (localName(entry[0]) === "txStyles" || localName(entry[0]) === "extLst")
    ) {
      entries.push([key, created]);
      inserted = true;
    }
    entries.push(entry);
  }
  if (!inserted) entries.push([key, created]);
  replaceNodeEntries(master, entries);
  return created;
}
