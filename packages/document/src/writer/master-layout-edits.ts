import {
  getChild,
  getChildArray,
  getNamespacedAttr,
  localName,
  type XmlNode,
} from "../reader/xml.js";
import type {
  PptxSourceModelAddSlideLayoutEdit,
  PptxSourceModelCloneSlideLayoutEdit,
} from "../source/index.js";
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
  const existingLayoutIdList = getChild(master, "sldLayoutIdLst");
  const layoutIdList = existingLayoutIdList ?? createLayoutIdList(master);
  if (existingLayoutIdList === undefined) {
    appendLayoutIds(layoutIdList, edit.initialLayoutEntries);
  }
  insertLayoutId(layoutIdList, edit, getChildArray(layoutIdList, "sldLayoutId").length);
}

export function applyCloneSlideLayoutEdit(
  root: XmlNode,
  edit: PptxSourceModelCloneSlideLayoutEdit,
): void {
  const master = getChild(root, "sldMaster");
  if (master === undefined) {
    throw new Error("writePptx: slide master part does not contain p:sldMaster root");
  }
  const existingLayoutIdList = getChild(master, "sldLayoutIdLst");
  const layoutIdList = existingLayoutIdList ?? createLayoutIdList(master);
  if (existingLayoutIdList === undefined) {
    appendLayoutIds(layoutIdList, edit.initialLayoutEntries);
  }
  insertLayoutId(layoutIdList, edit, edit.insertAt);
}

type SlideLayoutInsertionEdit =
  | PptxSourceModelAddSlideLayoutEdit
  | PptxSourceModelCloneSlideLayoutEdit;

function insertLayoutId(
  layoutIdList: XmlNode,
  edit: SlideLayoutInsertionEdit,
  insertAt: number,
): void {
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
  const inserted = [
    ...items.slice(0, insertAt),
    {
      "@_id": String(edit.newLayoutNumericId),
      [relationshipAttributeKey]: edit.newRelationshipId,
    },
    ...items.slice(insertAt),
  ];
  replaceNodeEntries(
    layoutIdList,
    Object.entries(layoutIdList).map(([key, value]) =>
      key === itemKey ? [key, inserted] : [key, value],
    ),
  );
}

function appendLayoutIds(
  layoutIdList: XmlNode,
  entries: SlideLayoutInsertionEdit["initialLayoutEntries"],
): void {
  if (entries.length === 0) return;
  const itemKey = namespacedChildKey(layoutIdList, "p:sldLayoutId", "sldLayoutId");
  layoutIdList[itemKey] = entries.map((entry) => ({
    "@_id": String(entry.numericId),
    "@_r:id": entry.relationshipId,
  }));
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
      ["transition", "timing", "hf", "txStyles", "extLst"].includes(localName(entry[0]))
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
