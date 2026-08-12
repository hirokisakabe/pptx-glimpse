import {
  getShapeTreeContainerPropertyKeys,
  isShapeTreeZOrderChildKey,
} from "../reader/shape-tree-child-classification.js";
import { getAttr, getChild, localName, type XmlNode } from "../reader/xml.js";
import type {
  PptxSourceModelAddChartEdit,
  PptxSourceModelAddConnectorEdit,
  PptxSourceModelAddPictureEdit,
  PptxSourceModelAddShapeEdit,
  PptxSourceModelAddTableEdit,
  PptxSourceModelAddTextBoxEdit,
  PptxSourceModelDeleteShapeEdit,
  PptxSourceModelGroupShapesEdit,
  PptxSourceModelMoveShapesEdit,
  PptxSourceModelReorderShapesEdit,
  PptxSourceModelSetBackgroundEdit,
  PptxSourceModelSetSlideBackgroundEdit,
  PptxSourceModelUngroupShapeEdit,
} from "../source/index.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import {
  appendShapeTreeNodeAtEnd,
  elementPrefix,
  ensurePictureNamespaces,
  getDrawingPartRoot,
  parseShapeFragmentXml,
  preserveNamespaceDeclarations,
  qualifiedSiblingName,
  remapElementPrefix,
} from "./dirty-part-xml-helpers.js";
import {
  deleteShapeXml,
  locateShapeTreeNode,
  locateShapeTreeNodeLocation,
  parseShapeLocator,
} from "./xml-locators.js";
import { replaceNodeEntries } from "./xml-node-utils.js";
import { getXmlChildOrder, parseXmlForEditing, setXmlChildOrder } from "./xml-serialization.js";

export function applyAddTextBoxEdit(root: XmlNode, edit: PptxSourceModelAddTextBoxEdit): void {
  applyAddSpEdit(root, edit);
}

export function applyAddShapeEdit(root: XmlNode, edit: PptxSourceModelAddShapeEdit): void {
  applyAddSpEdit(root, edit);
}

export function applyAddTableEdit(root: XmlNode, edit: PptxSourceModelAddTableEdit): void {
  const slide = getChild(root, "sld");
  if (slide !== undefined) ensurePictureNamespaces(slide);
  const spTree = getChild(getChild(slide, "cSld"), "spTree");
  if (spTree === undefined)
    throw new Error(`writePptx: slide '${edit.slidePartPath}' has no spTree`);
  assertShapeIdAvailable(spTree, edit.shapeId);
  appendShapeTreeNodeAtEnd(
    spTree,
    "p:graphicFrame",
    parseShapeFragmentXml(edit.xml, "graphicFrame"),
  );
}

export function applyAddConnectorEdit(root: XmlNode, edit: PptxSourceModelAddConnectorEdit): void {
  const spTree = getShapeTree(root, edit.slidePartPath);
  assertShapeIdAvailable(spTree, edit.shapeId);
  if (
    edit.startShapeId !== undefined &&
    locateShapeTreeNode(spTree, { nodeId: edit.startShapeId }) === undefined
  ) {
    throw new Error(`writePptx: connector start shape '${edit.startShapeId}' was not found`);
  }
  if (
    edit.endShapeId !== undefined &&
    locateShapeTreeNode(spTree, { nodeId: edit.endShapeId }) === undefined
  ) {
    throw new Error(`writePptx: connector end shape '${edit.endShapeId}' was not found`);
  }
  appendShapeTreeNodeAtEnd(spTree, "p:cxnSp", parseShapeFragmentXml(edit.xml, "cxnSp"));
}

export function applyAddPictureEdit(root: XmlNode, edit: PptxSourceModelAddPictureEdit): void {
  const drawingPart = getDrawingPartRoot(root);
  if (drawingPart !== undefined) ensurePictureNamespaces(drawingPart);
  const spTree = getShapeTree(root, edit.slidePartPath);
  assertShapeIdAvailable(spTree, edit.shapeId);
  appendShapeTreeNodeAtEnd(spTree, "p:pic", parseShapeFragmentXml(edit.xml, "pic"));
}

export function applyAddChartEdit(root: XmlNode, edit: PptxSourceModelAddChartEdit): void {
  const slide = getChild(root, "sld");
  if (slide !== undefined) ensurePictureNamespaces(slide);
  const spTree = getChild(getChild(slide, "cSld"), "spTree");
  if (spTree === undefined)
    throw new Error(`writePptx: slide '${edit.slidePartPath}' has no spTree`);
  assertShapeIdAvailable(spTree, edit.shapeId);
  appendShapeTreeNodeAtEnd(
    spTree,
    "p:graphicFrame",
    parseShapeFragmentXml(edit.xml, "graphicFrame"),
  );
}

export function applyDeleteShapeEdit(root: XmlNode, edit: PptxSourceModelDeleteShapeEdit): void {
  const locator = parseShapeLocator(edit.handle, "shape delete edit");
  const spTree = getChild(getChild(getChild(root, "sld"), "cSld"), "spTree");
  if (!deleteShapeXml(spTree, locator.nodeId)) {
    throw new Error(
      `writePptx: drawing delete handle '${locator.nodeId}' no longer matches a supported drawing`,
    );
  }
}

export function applyReorderShapesEdit(
  root: XmlNode,
  edit: PptxSourceModelReorderShapesEdit,
): void {
  const spTree = getShapeTree(root, edit.targetPartPath);
  const container = groupParentContainer(spTree, edit.parentGroupId, "reorder");
  const current = completeRememberedChildOrder(container);
  const shapeById = new Map<string, { key: string; value: unknown }>();
  for (const entry of current) {
    const nodeId = shapeTreeEntryNodeId(entry.value);
    if (nodeId !== undefined) {
      if (shapeById.has(nodeId)) {
        throw new Error(`writePptx: duplicate shape id '${nodeId}' in reordered container`);
      }
      shapeById.set(nodeId, entry);
    }
  }
  if (shapeById.size !== edit.shapeIds.length) {
    throw new Error("writePptx: reordered shape ids must contain every shape exactly once");
  }
  const orderedShapes = edit.shapeIds.map((shapeId) => {
    const entry = shapeById.get(shapeId);
    if (entry === undefined) {
      throw new Error(`writePptx: reordered shape '${shapeId}' was not found`);
    }
    return entry;
  });
  let orderedShapeIndex = 0;
  setXmlChildOrder(
    container,
    current.map((entry) =>
      shapeTreeEntryNodeId(entry.value) === undefined ? entry : orderedShapes[orderedShapeIndex++],
    ),
  );
}

export function applyMoveShapesEdit(root: XmlNode, edit: PptxSourceModelMoveShapesEdit): void {
  if (edit.shapeIds.length === 0) {
    throw new Error("writePptx: moved shape ids must not be empty");
  }
  const spTree = getShapeTree(root, edit.targetPartPath);
  if (edit.crossParent === true) {
    applyCrossParentMoveShapesEdit(spTree, edit);
    return;
  }
  const container = groupParentContainer(spTree, edit.parentGroupId, "move");
  const current = completeRememberedChildOrder(container);
  const containerPropertyKeys = getShapeTreeContainerPropertyKeys(container);
  const drawingEntries = current.filter((entry) =>
    isShapeTreeZOrderChildKey(entry.key, containerPropertyKeys),
  );
  const shapeById = new Map<string, { key: string; value: unknown }>();
  for (const entry of drawingEntries) {
    const nodeId = shapeTreeEntryNodeId(entry.value);
    if (nodeId === undefined) continue;
    if (shapeById.has(nodeId)) {
      throw new Error(`writePptx: duplicate shape id '${nodeId}' in moved container`);
    }
    shapeById.set(nodeId, entry);
  }

  const movedIds = new Set(edit.shapeIds);
  if (movedIds.size !== edit.shapeIds.length) {
    throw new Error("writePptx: moved shape ids contain a duplicate shape");
  }
  const movedEntries = edit.shapeIds.map((shapeId) => {
    const entry = shapeById.get(shapeId);
    if (entry === undefined) {
      throw new Error(`writePptx: moved shape '${shapeId}' was not found as a direct child`);
    }
    return entry;
  });
  const firstDrawingIndex = drawingEntries.findIndex((entry) => entry === movedEntries[0]);
  if (
    firstDrawingIndex < 0 ||
    movedEntries.some((entry, index) => drawingEntries[firstDrawingIndex + index] !== entry)
  ) {
    throw new Error("writePptx: moved shapes are not consecutive direct-child siblings");
  }
  if (edit.beforeShapeId !== undefined && movedIds.has(edit.beforeShapeId)) {
    throw new Error("writePptx: move anchor must not be inside the moved block");
  }

  const movedEntrySet = new Set(movedEntries);
  const remaining = drawingEntries.filter((entry) => !movedEntrySet.has(entry));
  const insertionIndex =
    edit.beforeShapeId === undefined
      ? remaining.length
      : remaining.findIndex((entry) => shapeTreeEntryNodeId(entry.value) === edit.beforeShapeId);
  if (insertionIndex < 0) {
    throw new Error(
      `writePptx: move anchor '${edit.beforeShapeId}' was not found as a direct child`,
    );
  }
  const reorderedDrawings = [
    ...remaining.slice(0, insertionIndex),
    ...movedEntries,
    ...remaining.slice(insertionIndex),
  ];
  let drawingIndex = 0;
  setXmlChildOrder(
    container,
    current.map((entry) =>
      isShapeTreeZOrderChildKey(entry.key, containerPropertyKeys)
        ? reorderedDrawings[drawingIndex++]
        : entry,
    ),
  );
}

function collectXmlDrawingNodeIds(value: unknown, output: Set<string>): void {
  if (Array.isArray(value)) {
    for (const child of value) collectXmlDrawingNodeIds(child, output);
    return;
  }
  if (typeof value !== "object" || value === null) return;
  const node = unsafeOoxmlBoundaryAssertion<XmlNode>(value);
  const nodeId = shapeTreeNodeId(node);
  if (nodeId !== undefined) output.add(nodeId);
  for (const child of Object.values(node)) collectXmlDrawingNodeIds(child, output);
}

function assertXmlConnectorBoundary(root: XmlNode, movedIds: ReadonlySet<string>): void {
  const visit = (node: XmlNode): void => {
    for (const [key, value] of Object.entries(node)) {
      if (key.startsWith("@_")) continue;
      const values = Array.isArray(value) ? value : [value];
      for (const childValue of values) {
        if (typeof childValue !== "object" || childValue === null || Array.isArray(childValue)) {
          continue;
        }
        const child = unsafeOoxmlBoundaryAssertion<XmlNode>(childValue);
        if (localName(key) === "cxnSp") {
          const nonVisual = getChild(child, "nvCxnSpPr");
          const connectorId = getAttr(getChild(nonVisual, "cNvPr"), "id");
          const connectorMoved = connectorId !== undefined && movedIds.has(connectorId);
          const connectionProperties = getChild(nonVisual, "cNvCxnSpPr");
          for (const endpointId of [
            getAttr(getChild(connectionProperties, "stCxn"), "id"),
            getAttr(getChild(connectionProperties, "endCxn"), "id"),
          ]) {
            if (endpointId !== undefined && connectorMoved !== movedIds.has(endpointId)) {
              throw new Error("writePptx: connector endpoint crosses the moved block boundary");
            }
          }
        }
        visit(child);
      }
    }
  };
  visit(root);
}

function applyCrossParentMoveShapesEdit(
  spTree: XmlNode,
  edit: PptxSourceModelMoveShapesEdit,
): void {
  assertIdentityMappedGroupChainXml(spTree, edit.parentGroupId, "move source");
  assertIdentityMappedGroupChainXml(spTree, edit.destinationParentGroupId, "move destination");
  const sourceContainer = groupParentContainer(spTree, edit.parentGroupId, "move source");
  const destinationContainer = groupParentContainer(
    spTree,
    edit.destinationParentGroupId,
    "move destination",
  );
  if (sourceContainer === destinationContainer) {
    throw new Error("writePptx: cross-parent move requires different containers");
  }

  const sourceOrder = completeRememberedChildOrder(sourceContainer);
  const sourceProperties = getShapeTreeContainerPropertyKeys(sourceContainer);
  const sourceDrawings = sourceOrder.filter((entry) =>
    isShapeTreeZOrderChildKey(entry.key, sourceProperties),
  );
  const sourceById = uniqueDirectDrawingEntries(sourceDrawings, "move source");
  const movedEntries = edit.shapeIds.map((shapeId) => {
    const entry = sourceById.get(shapeId);
    if (entry === undefined) {
      throw new Error(`writePptx: moved shape '${shapeId}' was not found as a source direct child`);
    }
    return entry;
  });
  const movedSubtreeIds = new Set<string>();
  for (const entry of movedEntries) collectXmlDrawingNodeIds(entry.value, movedSubtreeIds);
  assertXmlConnectorBoundary(spTree, movedSubtreeIds);
  const firstDrawingIndex = sourceDrawings.findIndex((entry) => entry === movedEntries[0]);
  if (
    firstDrawingIndex < 0 ||
    movedEntries.some((entry, index) => sourceDrawings[firstDrawingIndex + index] !== entry)
  ) {
    throw new Error("writePptx: moved shapes are not consecutive source direct-child siblings");
  }

  const movedEntrySet = new Set(movedEntries);
  if (xmlOrderContainsAnyEntry(movedEntries, destinationContainer)) {
    throw new Error("writePptx: move destination must not be inside the moved block");
  }
  const destinationOrder = completeRememberedChildOrder(destinationContainer);
  const destinationProperties = getShapeTreeContainerPropertyKeys(destinationContainer);
  const destinationDrawings = destinationOrder.filter((entry) =>
    isShapeTreeZOrderChildKey(entry.key, destinationProperties),
  );
  uniqueDirectDrawingEntries(destinationDrawings, "move destination");
  if (edit.beforeShapeId !== undefined && edit.shapeIds.includes(edit.beforeShapeId)) {
    throw new Error("writePptx: move anchor must not be inside the moved block");
  }

  const insertionIndex =
    edit.beforeShapeId === undefined
      ? crossParentEndInsertionIndex(destinationOrder, destinationProperties)
      : destinationOrder.findIndex(
          (entry) =>
            isShapeTreeZOrderChildKey(entry.key, destinationProperties) &&
            shapeTreeEntryNodeId(entry.value) === edit.beforeShapeId,
        );
  if (insertionIndex < 0) {
    throw new Error(
      `writePptx: move anchor '${edit.beforeShapeId}' was not found as a destination direct child`,
    );
  }

  replaceContainerChildren(
    sourceContainer,
    sourceOrder.filter((entry) => !movedEntrySet.has(entry)),
  );
  replaceContainerChildren(destinationContainer, [
    ...destinationOrder.slice(0, insertionIndex),
    ...movedEntries,
    ...destinationOrder.slice(insertionIndex),
  ]);
}

function assertIdentityMappedGroupChainXml(
  spTree: XmlNode,
  groupId: string | undefined,
  context: string,
): void {
  if (groupId === undefined) return;
  const path = findGroupPathXml(spTree, groupId);
  if (path === undefined) {
    throw new Error(`writePptx: ${context} group '${groupId}' was not found`);
  }
  for (const group of path) {
    const xfrm = getChild(getChild(group, "grpSpPr"), "xfrm");
    const off = getChild(xfrm, "off");
    const ext = getChild(xfrm, "ext");
    const childOff = getChild(xfrm, "chOff");
    const childExt = getChild(xfrm, "chExt");
    const width = getAttr(ext, "cx");
    const height = getAttr(ext, "cy");
    const childWidth = getAttr(childExt, "cx");
    const childHeight = getAttr(childExt, "cy");
    const offsetX = getAttr(off, "x");
    const offsetY = getAttr(off, "y");
    const childOffsetX = getAttr(childOff, "x");
    const childOffsetY = getAttr(childOff, "y");
    if (
      xfrm === undefined ||
      offsetX === undefined ||
      offsetY === undefined ||
      childOffsetX === undefined ||
      childOffsetY === undefined ||
      width === undefined ||
      height === undefined ||
      childWidth === undefined ||
      childHeight === undefined ||
      Number(offsetX) !== Number(childOffsetX) ||
      Number(offsetY) !== Number(childOffsetY) ||
      Number(width) !== Number(childWidth) ||
      Number(height) !== Number(childHeight) ||
      Number(width) <= 0 ||
      Number(height) <= 0 ||
      Number(childWidth) <= 0 ||
      Number(childHeight) <= 0 ||
      Number(getAttr(xfrm, "rot") ?? 0) !== 0 ||
      isTrueXmlAttribute(getAttr(xfrm, "flipH")) ||
      isTrueXmlAttribute(getAttr(xfrm, "flipV"))
    ) {
      throw new Error(
        `writePptx: ${context} ancestor '${shapeTreeNodeId(group) ?? "unknown"}' is not identity-mapped`,
      );
    }
  }
}

function findGroupPathXml(
  container: XmlNode,
  groupId: string,
  ancestors: readonly XmlNode[] = [],
): readonly XmlNode[] | undefined {
  for (const entry of completeRememberedChildOrder(container)) {
    if (localName(entry.key) !== "grpSp") continue;
    if (typeof entry.value !== "object" || entry.value === null || Array.isArray(entry.value)) {
      continue;
    }
    const group = unsafeOoxmlBoundaryAssertion<XmlNode>(entry.value);
    const path = [...ancestors, group];
    if (shapeTreeNodeId(group) === groupId) return path;
    const nested = findGroupPathXml(group, groupId, path);
    if (nested !== undefined) return nested;
  }
  return undefined;
}

function uniqueDirectDrawingEntries(
  drawings: readonly { readonly key: string; readonly value: unknown }[],
  context: string,
): ReadonlyMap<string, { readonly key: string; readonly value: unknown }> {
  const byId = new Map<string, { key: string; value: unknown }>();
  for (const entry of drawings) {
    const nodeId = shapeTreeEntryNodeId(entry.value);
    if (nodeId === undefined) continue;
    if (byId.has(nodeId)) {
      throw new Error(`writePptx: duplicate shape id '${nodeId}' in ${context} container`);
    }
    byId.set(nodeId, entry);
  }
  return byId;
}

function crossParentEndInsertionIndex(
  order: readonly { readonly key: string; readonly value: unknown }[],
  propertyKeys: ReadonlySet<string>,
): number {
  let lastDrawing = -1;
  for (let index = 0; index < order.length; index += 1) {
    const entry = order[index];
    if (entry !== undefined && isShapeTreeZOrderChildKey(entry.key, propertyKeys)) {
      lastDrawing = index;
    }
  }
  if (lastDrawing >= 0) return lastDrawing + 1;
  const trailingExtension = order.findIndex((entry) => localName(entry.key) === "extLst");
  return trailingExtension >= 0 ? trailingExtension : order.length;
}

function xmlOrderContainsAnyEntry(
  entries: readonly { readonly value: unknown }[],
  candidate: XmlNode,
): boolean {
  const targets = new Set(entries.map((entry) => entry.value));
  const visit = (value: unknown): boolean => {
    if (targets.has(value)) return value === candidate || xmlValueContainsNode(value, candidate);
    return xmlValueContainsNode(value, candidate);
  };
  return entries.some((entry) => visit(entry.value));
}

function xmlValueContainsNode(value: unknown, candidate: XmlNode): boolean {
  if (value === candidate) return true;
  if (Array.isArray(value)) return value.some((child) => xmlValueContainsNode(child, candidate));
  if (typeof value !== "object" || value === null) return false;
  const node = unsafeOoxmlBoundaryAssertion<XmlNode>(value);
  return Object.values(node).some((child) => xmlValueContainsNode(child, candidate));
}

function completeRememberedChildOrder(container: XmlNode): { key: string; value: unknown }[] {
  const remembered = [...getXmlChildOrder(container)];
  const remaining = Object.entries(container).flatMap(([key, grouped]) => {
    if (key.startsWith("@_") || key === "?xml") return [];
    const values = Array.isArray(grouped)
      ? unsafeOoxmlBoundaryAssertion<unknown[]>(grouped)
      : [grouped];
    return values.map((value) => ({ key, value }));
  });
  for (const entry of remembered) {
    const index = remaining.findIndex(
      (candidate) => candidate.key === entry.key && candidate.value === entry.value,
    );
    if (index >= 0) remaining.splice(index, 1);
  }
  return [...remembered, ...remaining];
}

export function applyGroupShapesEdit(root: XmlNode, edit: PptxSourceModelGroupShapesEdit): void {
  const spTree = getShapeTree(root, edit.targetPartPath);
  assertShapeIdAvailable(spTree, edit.groupId);
  const parent = groupParentContainer(spTree, edit.parentGroupId, "group");
  const order = getXmlChildOrder(parent);
  const selectedIds = new Set(edit.shapeIds);
  if (selectedIds.size !== edit.shapeIds.length) {
    throw new Error("writePptx: grouped shape ids contain a duplicate shape");
  }
  const selectedEntries = order.filter((entry) => {
    const nodeId = shapeTreeEntryNodeId(entry.value);
    return nodeId !== undefined && selectedIds.has(nodeId);
  });
  if (
    selectedEntries.length !== edit.shapeIds.length ||
    selectedEntries.some(
      (entry, index) => shapeTreeEntryNodeId(entry.value) !== edit.shapeIds[index],
    )
  ) {
    throw new Error("writePptx: grouped shapes were not found as ordered siblings");
  }
  const drawingEntries = order.filter((entry) => shapeTreeEntryNodeId(entry.value) !== undefined);
  const firstDrawingIndex = drawingEntries.findIndex((entry) => entry === selectedEntries[0]);
  if (
    firstDrawingIndex < 0 ||
    selectedEntries.some((entry, index) => drawingEntries[firstDrawingIndex + index] !== entry)
  ) {
    throw new Error("writePptx: grouped shapes are not consecutive siblings");
  }

  const group = parseShapeFragmentXml(edit.xml, "grpSp");
  if (shapeTreeNodeId(group) !== edit.groupId) {
    throw new Error(`writePptx: grouped XML id does not match '${edit.groupId}'`);
  }
  if (getXmlChildOrder(group).some(isShapeTreeEntry)) {
    throw new Error("writePptx: grouped XML must be an empty group shell");
  }
  for (const entry of selectedEntries) {
    appendShapeTreeNodeAtEnd(group, entry.key, unsafeOoxmlBoundaryAssertion<XmlNode>(entry.value));
  }
  const firstOrderIndex = order.findIndex((entry) => entry === selectedEntries[0]);
  const selectedEntrySet = new Set(selectedEntries);
  const groupKey = qualifiedSiblingName(selectedEntries[0]?.key ?? "p:sp", "grpSp");
  const nextOrder = order.flatMap((entry, index) => {
    if (index === firstOrderIndex) return [{ key: groupKey, value: group }];
    return selectedEntrySet.has(entry) ? [] : [entry];
  });
  replaceContainerChildren(parent, nextOrder);
}

export function applyUngroupShapeEdit(root: XmlNode, edit: PptxSourceModelUngroupShapeEdit): void {
  const spTree = getShapeTree(root, edit.targetPartPath);
  const location = locateShapeTreeNodeLocation(spTree, { nodeId: edit.groupId });
  if (location === undefined || location.nodeKind !== "grpSp") {
    throw new Error(`writePptx: ungroup target '${edit.groupId}' is not a group shape`);
  }
  assertIdentityGroupXml(location.node, edit.groupId);
  const children = getXmlChildOrder(location.node)
    .filter(isShapeTreeEntry)
    .map((entry) => ({
      key: entry.key,
      value: preserveNamespaceDeclarations(location.node, entry.value),
    }));
  const parentOrder = getXmlChildOrder(location.parentContainer);
  const groupIndex = parentOrder.findIndex((entry) => entry.value === location.node);
  if (groupIndex < 0) {
    throw new Error(`writePptx: ungroup target '${edit.groupId}' has no parent slot`);
  }
  const nextOrder = parentOrder.flatMap((entry, index) =>
    index === groupIndex ? children : entry.value === location.node ? [] : [entry],
  );
  replaceContainerChildren(location.parentContainer, nextOrder);
}

export function applySetBackgroundEdit(
  root: XmlNode,
  edit: PptxSourceModelSetBackgroundEdit | PptxSourceModelSetSlideBackgroundEdit,
): void {
  const targetPartPath = edit.kind === "setBackground" ? edit.targetPartPath : edit.slidePartPath;
  const drawingPart = getDrawingPartRoot(root);
  const cSldKey =
    drawingPart === undefined
      ? undefined
      : Object.keys(drawingPart).find((key) => !key.startsWith("@_") && localName(key) === "cSld");
  const cSld = getChild(drawingPart, "cSld");
  if (drawingPart === undefined || cSldKey === undefined || cSld === undefined) {
    throw new Error(`writePptx: drawing part '${targetPartPath}' has no p:cSld`);
  }
  const spTreeKey = Object.keys(cSld).find(
    (key) => !key.startsWith("@_") && localName(key) === "spTree",
  );
  if (spTreeKey === undefined) {
    throw new Error(`writePptx: drawing part '${targetPartPath}' has no p:spTree`);
  }
  const existingBackgroundKey = Object.keys(cSld).find(
    (key) => !key.startsWith("@_") && localName(key) === "bg",
  );
  const existingBackground = getChild(cSld, "bg");
  if (edit.xml === undefined) {
    replaceNodeEntries(
      cSld,
      Object.entries(cSld).filter(([key]) => key.startsWith("@_") || localName(key) !== "bg"),
    );
    return;
  }

  drawingPart["@_xmlns:a"] ??= "http://schemas.openxmlformats.org/drawingml/2006/main";
  if (edit.relationshipId !== undefined) {
    drawingPart["@_xmlns:r"] ??=
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  }
  const parsedBackground = getChild(parseXmlForEditing(edit.xml), "bg");
  if (parsedBackground === undefined) {
    throw new Error("writePptx: background edit XML fragment does not contain a p:bg root element");
  }
  const backgroundKey = existingBackgroundKey ?? qualifiedSiblingName(cSldKey, "bg");
  const remappedBackground = remapElementPrefix(
    parsedBackground,
    "p",
    elementPrefix(backgroundKey),
  );
  const background = preserveNamespaceDeclarations(existingBackground, remappedBackground);

  const entries: [string, unknown][] = [];
  let inserted = false;
  for (const [key, value] of Object.entries(cSld)) {
    if (!key.startsWith("@_") && localName(key) === "bg") continue;
    if (!inserted && key === spTreeKey) {
      entries.push([backgroundKey, background]);
      inserted = true;
    }
    entries.push([key, value]);
  }
  replaceNodeEntries(cSld, entries);
}

function applyAddSpEdit(
  root: XmlNode,
  edit: PptxSourceModelAddTextBoxEdit | PptxSourceModelAddShapeEdit,
): void {
  const spTree = getShapeTree(root, edit.slidePartPath);
  assertShapeIdAvailable(spTree, edit.shapeId);
  appendShapeTreeNodeAtEnd(spTree, "p:sp", parseShapeFragmentXml(edit.xml, "sp"));
}

function getShapeTree(root: XmlNode, partPath: string): XmlNode {
  const spTree = getChild(getChild(getDrawingPartRoot(root), "cSld"), "spTree");
  if (spTree === undefined) throw new Error(`writePptx: drawing part '${partPath}' has no spTree`);
  return spTree;
}

function assertShapeIdAvailable(spTree: XmlNode, shapeId: string): void {
  if (locateShapeTreeNode(spTree, { nodeId: shapeId }) !== undefined) {
    throw new Error(`writePptx: shape id '${shapeId}' already exists in source XML`);
  }
}

function groupParentContainer(
  spTree: XmlNode,
  parentGroupId: string | undefined,
  operation: string,
): XmlNode {
  if (parentGroupId === undefined) return spTree;
  const location = locateShapeTreeNodeLocation(spTree, { nodeId: parentGroupId });
  if (location === undefined || location.nodeKind !== "grpSp") {
    throw new Error(`writePptx: ${operation} parent '${parentGroupId}' is not a group shape`);
  }
  return location.node;
}

function assertIdentityGroupXml(group: XmlNode, groupId: string): void {
  const groupProperties = getChild(group, "grpSpPr");
  const xfrm = getChild(groupProperties, "xfrm");
  const off = getChild(xfrm, "off");
  const ext = getChild(xfrm, "ext");
  const childOff = getChild(xfrm, "chOff");
  const childExt = getChild(xfrm, "chExt");
  const offX = getAttr(off, "x");
  const offY = getAttr(off, "y");
  const extentWidth = getAttr(ext, "cx");
  const extentHeight = getAttr(ext, "cy");
  const identity =
    xfrm !== undefined &&
    offX !== undefined &&
    offY !== undefined &&
    extentWidth !== undefined &&
    extentHeight !== undefined &&
    Number(getAttr(xfrm, "rot") ?? 0) === 0 &&
    !isTrueXmlAttribute(getAttr(xfrm, "flipH")) &&
    !isTrueXmlAttribute(getAttr(xfrm, "flipV")) &&
    offX === getAttr(childOff, "x") &&
    offY === getAttr(childOff, "y") &&
    extentWidth === getAttr(childExt, "cx") &&
    extentHeight === getAttr(childExt, "cy");
  const unsupportedProperties = getXmlChildOrder(groupProperties ?? {}).some(
    (entry) => localName(entry.key) !== "xfrm",
  );
  const unsupportedNonVisual = !hasMinimalGroupNonVisualProperties(group);
  const unsupportedGroupChildren = getXmlChildOrder(group).some((entry) => {
    const name = localName(entry.key);
    return name !== "nvGrpSpPr" && name !== "grpSpPr" && !isShapeTreeEntry(entry);
  });
  if (!identity || unsupportedProperties || unsupportedNonVisual || unsupportedGroupChildren) {
    throw new Error(`writePptx: group '${groupId}' cannot be losslessly expanded`);
  }
}

function hasMinimalGroupNonVisualProperties(group: XmlNode): boolean {
  const nonVisual = getChild(group, "nvGrpSpPr");
  const cNvPr = getChild(nonVisual, "cNvPr");
  const cNvGrpSpPr = getChild(nonVisual, "cNvGrpSpPr");
  const nvPr = getChild(nonVisual, "nvPr");
  if (
    nonVisual === undefined ||
    cNvPr === undefined ||
    cNvGrpSpPr === undefined ||
    nvPr === undefined
  ) {
    return false;
  }
  const nonVisualOrder = getXmlChildOrder(nonVisual);
  if (
    nonVisualOrder.some(
      (entry) => !["cNvPr", "cNvGrpSpPr", "nvPr"].includes(localName(entry.key)),
    ) ||
    ["cNvPr", "cNvGrpSpPr", "nvPr"].some(
      (name) => nonVisualOrder.filter((entry) => localName(entry.key) === name).length !== 1,
    )
  ) {
    return false;
  }
  return (
    xmlNodeAttributesAreAllowed(nonVisual, new Set()) &&
    xmlNodeHasOnlyAttributes(cNvPr, new Set(["id", "name"])) &&
    xmlNodeHasOnlyAttributes(cNvGrpSpPr, new Set()) &&
    xmlNodeHasOnlyAttributes(nvPr, new Set())
  );
}

function xmlNodeHasOnlyAttributes(node: XmlNode, allowed: ReadonlySet<string>): boolean {
  return Object.keys(node).every((key) => {
    if (key === "#text") return String(node[key]).trim().length === 0;
    if (!key.startsWith("@_")) return false;
    const qualifiedName = key.slice(2);
    return (
      qualifiedName === "xmlns" ||
      qualifiedName.startsWith("xmlns:") ||
      (!qualifiedName.includes(":") && allowed.has(qualifiedName))
    );
  });
}

function xmlNodeAttributesAreAllowed(node: XmlNode, allowed: ReadonlySet<string>): boolean {
  return Object.keys(node)
    .filter((key) => key.startsWith("@_"))
    .every((key) => {
      const qualifiedName = key.slice(2);
      return (
        qualifiedName === "xmlns" ||
        qualifiedName.startsWith("xmlns:") ||
        (!qualifiedName.includes(":") && allowed.has(qualifiedName))
      );
    });
}

function isShapeTreeEntry(entry: { readonly key: string; readonly value: unknown }): boolean {
  return ["sp", "pic", "cxnSp", "graphicFrame", "grpSp"].includes(localName(entry.key));
}

function replaceContainerChildren(
  container: XmlNode,
  order: readonly { readonly key: string; readonly value: unknown }[],
): void {
  const grouped = new Map<string, unknown[]>();
  for (const entry of order) {
    const values = grouped.get(entry.key) ?? [];
    values.push(entry.value);
    grouped.set(entry.key, values);
  }
  const entries: [string, unknown][] = Object.entries(container).filter(([key]) =>
    key.startsWith("@_"),
  );
  for (const [key, values] of grouped) entries.push([key, values]);
  replaceNodeEntries(container, entries);
  setXmlChildOrder(container, order);
}

function isTrueXmlAttribute(value: string | undefined): boolean {
  return value === "1" || value === "true";
}

function shapeTreeNodeId(node: XmlNode): string | undefined {
  const nonVisualProperties =
    getChild(node, "nvSpPr") ??
    getChild(node, "nvPicPr") ??
    getChild(node, "nvCxnSpPr") ??
    getChild(node, "nvGrpSpPr") ??
    getChild(node, "nvGraphicFramePr");
  return getAttr(getChild(nonVisualProperties, "cNvPr"), "id");
}

function shapeTreeEntryNodeId(value: unknown): string | undefined {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return undefined;
  return shapeTreeNodeId(unsafeOoxmlBoundaryAssertion<XmlNode>(value));
}
