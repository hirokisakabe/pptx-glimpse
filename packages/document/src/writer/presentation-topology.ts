import { getChild, getNamespacedAttr, parseXml, type XmlNode } from "../reader/xml.js";
import type { SlideTopologyOperation } from "../source/edit-descriptors.js";
import type { PptxSourceModel, RelationshipId } from "../source/index.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import {
  namespacedAttributeKey,
  namespacedChildKey,
  stripXmlProcessingInstruction,
} from "./xml-node-utils.js";
import { encodeXml, textDecoder, XML_DECLARATION, xmlBuilder } from "./xml-serialization.js";

export function serializePresentationWithSlideTopologyEdits(
  source: PptxSourceModel,
  operations: readonly SlideTopologyOperation[],
): Uint8Array {
  const rawPart = source.packageGraph.rawParts?.find(
    (part) => part.partPath === source.presentation.partPath,
  );
  if (rawPart === undefined) {
    throw new Error(
      `writePptx: presentation part '${source.presentation.partPath}' has no preserved raw package material`,
    );
  }
  if (rawPart.kind !== "binary") {
    throw new Error("writePptx: presentation XML tree part patching is not implemented");
  }

  const root = parseXml(textDecoder.decode(rawPart.bytes));
  const presentation = getChild(root, "presentation");
  if (presentation === undefined) {
    throw new Error("writePptx: presentation part does not contain p:presentation root");
  }
  const sldIdLst = ensureSlideIdList(presentation);

  for (const operation of operations) {
    switch (operation.kind) {
      case "appendSlide":
        appendSlideId(sldIdLst, operation.newRelationshipId, operation.newSlideNumericId);
        break;
      case "insertSlideAfter":
        insertSlideIdAfter(
          sldIdLst,
          operation.sourceRelationshipId,
          operation.newRelationshipId,
          operation.newSlideNumericId,
        );
        break;
      case "removeSlide":
        removeSlideId(sldIdLst, operation.relationshipId);
        break;
      case "moveSlide":
        moveSlideId(sldIdLst, operation.relationshipId, operation.toIndex);
        break;
    }
  }

  return encodeXml(XML_DECLARATION + xmlBuilder.build(stripXmlProcessingInstruction(root)));
}

function ensureSlideIdList(presentation: XmlNode): XmlNode {
  const existing = getChild(presentation, "sldIdLst");
  if (existing !== undefined) return existing;
  const key = namespacedChildKey(presentation, "p:sldIdLst", "sldIdLst");
  const created: XmlNode = {};
  presentation[key] = created;
  return created;
}

function appendSlideId(
  sldIdLst: XmlNode,
  newRelationshipId: RelationshipId,
  newSlideNumericId: number,
): void {
  const { key, items } = slideIdEntries(sldIdLst);
  if (items.some((item) => getRelationshipAttr(item) === newRelationshipId)) return;
  const relationshipAttrKey =
    items[0] === undefined ? "@_r:id" : namespacedAttributeKey(items[0], "r:id", "id");
  const newNode: XmlNode = {
    "@_id": String(newSlideNumericId),
    [relationshipAttrKey]: newRelationshipId,
  };
  sldIdLst[key] = [...items, newNode];
}

function insertSlideIdAfter(
  sldIdLst: XmlNode,
  sourceRelationshipId: RelationshipId,
  newRelationshipId: RelationshipId,
  newSlideNumericId: number,
): void {
  const { key, items } = slideIdEntries(sldIdLst);
  const sourceIndex = items.findIndex((item) => getRelationshipAttr(item) === sourceRelationshipId);
  if (sourceIndex === -1) {
    throw new Error(
      `writePptx: slide relationship '${sourceRelationshipId}' was not found in p:sldIdLst`,
    );
  }
  if (items.some((item) => getRelationshipAttr(item) === newRelationshipId)) return;
  const relationshipAttrKey = namespacedAttributeKey(items[sourceIndex], "r:id", "id");
  const newNode: XmlNode = {
    "@_id": String(newSlideNumericId),
    [relationshipAttrKey]: newRelationshipId,
  };
  sldIdLst[key] = [...items.slice(0, sourceIndex + 1), newNode, ...items.slice(sourceIndex + 1)];
}

function removeSlideId(sldIdLst: XmlNode, relationshipId: RelationshipId): void {
  const { key, items } = slideIdEntries(sldIdLst);
  sldIdLst[key] = items.filter((item) => getRelationshipAttr(item) !== relationshipId);
}

function moveSlideId(sldIdLst: XmlNode, relationshipId: RelationshipId, toIndex: number): void {
  const { key, items } = slideIdEntries(sldIdLst);
  const fromIndex = items.findIndex((item) => getRelationshipAttr(item) === relationshipId);
  if (fromIndex === -1) {
    throw new Error(
      `writePptx: slide relationship '${relationshipId}' was not found in p:sldIdLst`,
    );
  }
  if (toIndex < 0 || toIndex >= items.length) {
    throw new Error(`writePptx: slide move target index '${toIndex}' is out of range`);
  }
  if (fromIndex === toIndex) return;

  const moved = [...items];
  const [item] = moved.splice(fromIndex, 1);
  if (item === undefined) return;
  moved.splice(toIndex, 0, item);
  sldIdLst[key] = moved;
}

function slideIdEntries(sldIdLst: XmlNode): { readonly key: string; readonly items: XmlNode[] } {
  const key = namespacedChildKey(sldIdLst, "p:sldId", "sldId");
  const value = sldIdLst[key];
  if (value === undefined || value === null) return { key, items: [] };
  return {
    key,
    items: Array.isArray(value)
      ? unsafeOoxmlBoundaryAssertion<XmlNode[]>(value)
      : [unsafeOoxmlBoundaryAssertion<XmlNode>(value)],
  };
}

function getRelationshipAttr(node: XmlNode): string | undefined {
  return getNamespacedAttr(node, "id");
}
