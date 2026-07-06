/**
 * `writePptx(source)` - the first round-trip slice of the PptxSourceModel source writer.
 *
 * The writer targets structural round-trip preservation rather than byte equality. It is
 * not a package patcher that preserves XML attribute order, namespace prefix placement,
 * ZIP metadata, or defaulted OOXML values. Content types and relationships are
 * regenerated structurally from `packageGraph`; media bytes, unknown parts, and
 * non-bookkeeping raw parts prefer the raw package material preserved by the reader.
 * Only dirty scopes are updated according to supported PptxSourceModel operations.
 *
 * The current slice supports plain text-run replacement, paragraph replacement,
 * top-level shape transform offset/extent edits, and slide duplicate/delete topology
 * edits. Dirty slide XML parts are reserialized by replacing only targeted values via
 * stable source handles; slide duplicate/delete patches presentation bookkeeping while
 * preserving unrelated raw package material.
 * New-content edits (new slides, text boxes, connectors) finalize their XML and id
 * numbering at edit time in `source/editing.ts` / `source/shape-xml.ts`; this writer
 * never generates new-content XML and only applies insertion positions (`p:spTree`
 * splice and `p:sldIdLst` patching).
 * Node-level XML splicing and precise unsupported raw-sidecar invalidation belong to
 * later writer slices, but the API and dirty-scope tracking remain shaped for that
 * extension path.
 */

import { XMLBuilder } from "fast-xml-parser";
import { zipSync } from "fflate";

import {
  getAttr,
  getChild,
  getChildArray,
  getNamespacedAttr,
  localName,
  parseXml,
  type XmlNode,
} from "../reader/xml.js";
import type {
  EditableTextRunProperties,
  PartPath,
  PartRelationships,
  PptxSourceModel,
  PptxSourceModelAddConnectorEdit,
  PptxSourceModelAddEmptySlideFromLayoutEdit,
  PptxSourceModelAddTextBoxEdit,
  PptxSourceModelDeleteShapeEdit,
  PptxSourceModelDeleteSlideEdit,
  PptxSourceModelDuplicateSlideEdit,
  PptxSourceModelEdit,
  PptxSourceModelParagraphTextEdit,
  PptxSourceModelShapeTransformEdit,
  PptxSourceModelTextRunEdit,
  PptxSourceModelTextRunPropertiesEdit,
  RawOoxmlNode,
  RawPackagePart,
  RelationshipId,
} from "../source/index.js";
import { isRelationshipPart, relationshipsPartPath } from "../source/package-paths.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";

/** `writePptx` output. */
export type WritePptxOutput = Uint8Array;

const CONTENT_TYPES_PART = "[Content_Types].xml";
const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: true,
});

/**
 * Writes a PptxSourceModel source back to PPTX package bytes.
 *
 * This initial round-trip writer prefers unedited package material
 * to create preserved output. If the raw bytes needed to patch a dirty part are unavailable, or
 * a non-bookkeeping part lacks required raw bytes,
 * it throws instead of regenerating content implicitly.
 */
export function writePptx(source: PptxSourceModel): WritePptxOutput {
  const textRunEdits = source.edits?.filter(isTextRunEdit) ?? [];
  const textRunPropertiesEdits = source.edits?.filter(isTextRunPropertiesEdit) ?? [];
  const paragraphTextEdits = source.edits?.filter(isParagraphTextEdit) ?? [];
  const shapeTransformEdits = source.edits?.filter(isShapeTransformEdit) ?? [];
  const addTextBoxEdits = source.edits?.filter(isAddTextBoxEdit) ?? [];
  const addConnectorEdits = source.edits?.filter(isAddConnectorEdit) ?? [];
  const deleteShapeEdits = source.edits?.filter(isDeleteShapeEdit) ?? [];
  const addEmptySlideFromLayoutEdits = source.edits?.filter(isAddEmptySlideFromLayoutEdit) ?? [];
  const duplicateSlideEdits = source.edits?.filter(isDuplicateSlideEdit) ?? [];
  const deleteSlideEdits = source.edits?.filter(isDeleteSlideEdit) ?? [];
  validateEdits(
    textRunEdits,
    textRunPropertiesEdits,
    paragraphTextEdits,
    shapeTransformEdits,
    addTextBoxEdits,
    addConnectorEdits,
    deleteShapeEdits,
  );
  const dirtyPartPaths = new Set([
    ...textRunEdits.map((edit) => edit.handle.partPath),
    ...textRunPropertiesEdits.map((edit) => edit.handle.partPath),
    ...paragraphTextEdits.map((edit) => edit.handle.partPath),
    ...shapeTransformEdits.map((edit) => edit.handle.partPath),
    ...deleteShapeEdits.map((edit) => edit.handle.partPath),
    ...addTextBoxEdits.map((edit) => edit.slidePartPath),
    ...addConnectorEdits.map((edit) => edit.slidePartPath),
  ]);
  const hasSlideTopologyEdits =
    addEmptySlideFromLayoutEdits.length > 0 ||
    duplicateSlideEdits.length > 0 ||
    deleteSlideEdits.length > 0;
  const files: Record<string, Uint8Array> = {
    [CONTENT_TYPES_PART]: encodeXml(serializeContentTypes(source.packageGraph.contentTypes)),
  };

  const written = new Set<string>([CONTENT_TYPES_PART]);

  for (const relationships of source.packageGraph.relationships) {
    const relsPath = relationshipsPartPath(relationships.sourcePartPath);
    files[relsPath] = encodeXml(serializeRelationships(relationships));
    written.add(relsPath);
  }

  for (const media of source.packageGraph.media) {
    files[media.partPath] = media.bytes;
    written.add(media.partPath);
  }

  for (const rawPart of source.packageGraph.rawParts ?? []) {
    if (hasSlideTopologyEdits && rawPart.partPath === source.presentation.partPath) continue;
    if (dirtyPartPaths.has(rawPart.partPath)) continue;
    files[rawPart.partPath] = serializeRawPackagePart(rawPart);
    written.add(rawPart.partPath);
  }

  for (const partPath of dirtyPartPaths) {
    files[partPath] = serializeDirtyXmlPart(
      source,
      partPath,
      textRunEdits,
      textRunPropertiesEdits,
      paragraphTextEdits,
      shapeTransformEdits,
      addTextBoxEdits,
      addConnectorEdits,
      deleteShapeEdits,
    );
    written.add(partPath);
  }

  if (hasSlideTopologyEdits) {
    files[source.presentation.partPath] = serializePresentationWithSlideTopologyEdits(
      source,
      mergeSlideTopologyEdits(source.edits ?? []),
    );
    written.add(source.presentation.partPath);
  }

  for (const part of source.packageGraph.parts) {
    if (written.has(part.partPath)) continue;
    if (part.contentType === RELS_CONTENT_TYPE || isRelationshipPart(part.partPath)) continue;
    throw new Error(
      `writePptx: no preserved package material for part '${part.partPath}'; ` +
        "edited part generation is not implemented in the no-edit writer",
    );
  }

  return zipSync(files);
}

function isTextRunEdit(edit: PptxSourceModelEdit): edit is PptxSourceModelTextRunEdit {
  return edit.kind === "replaceTextRunPlainText";
}

function isTextRunPropertiesEdit(
  edit: PptxSourceModelEdit,
): edit is PptxSourceModelTextRunPropertiesEdit {
  return edit.kind === "updateTextRunProperties";
}

function isParagraphTextEdit(edit: PptxSourceModelEdit): edit is PptxSourceModelParagraphTextEdit {
  return edit.kind === "replaceParagraphPlainText";
}

function isShapeTransformEdit(
  edit: PptxSourceModelEdit,
): edit is PptxSourceModelShapeTransformEdit {
  return edit.kind === "updateShapeTransform";
}

function isAddTextBoxEdit(edit: PptxSourceModelEdit): edit is PptxSourceModelAddTextBoxEdit {
  return edit.kind === "addTextBox";
}

function isAddConnectorEdit(edit: PptxSourceModelEdit): edit is PptxSourceModelAddConnectorEdit {
  return edit.kind === "addConnector";
}

function isDeleteShapeEdit(edit: PptxSourceModelEdit): edit is PptxSourceModelDeleteShapeEdit {
  return edit.kind === "deleteShape";
}

function isAddEmptySlideFromLayoutEdit(
  edit: PptxSourceModelEdit,
): edit is PptxSourceModelAddEmptySlideFromLayoutEdit {
  return edit.kind === "addEmptySlideFromLayout";
}

function isDuplicateSlideEdit(
  edit: PptxSourceModelEdit,
): edit is PptxSourceModelDuplicateSlideEdit {
  return edit.kind === "duplicateSlide";
}

function isDeleteSlideEdit(edit: PptxSourceModelEdit): edit is PptxSourceModelDeleteSlideEdit {
  return edit.kind === "deleteSlide";
}

function serializeDirtyXmlPart(
  source: PptxSourceModel,
  partPath: PartPath,
  textRunEdits: readonly PptxSourceModelTextRunEdit[],
  textRunPropertiesEdits: readonly PptxSourceModelTextRunPropertiesEdit[],
  paragraphTextEdits: readonly PptxSourceModelParagraphTextEdit[],
  shapeTransformEdits: readonly PptxSourceModelShapeTransformEdit[],
  addTextBoxEdits: readonly PptxSourceModelAddTextBoxEdit[],
  addConnectorEdits: readonly PptxSourceModelAddConnectorEdit[],
  deleteShapeEdits: readonly PptxSourceModelDeleteShapeEdit[],
): Uint8Array {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === partPath);
  if (rawPart === undefined) {
    throw new Error(`writePptx: dirty part '${partPath}' has no preserved raw package material`);
  }
  if (rawPart.kind !== "binary") {
    throw new Error(`writePptx: dirty XML tree part '${partPath}' patching is not implemented`);
  }

  const root = parseXml(textDecoder.decode(rawPart.bytes));
  for (const edit of addTextBoxEdits.filter((edit) => edit.slidePartPath === partPath)) {
    applyAddTextBoxEdit(root, edit);
  }
  for (const edit of addConnectorEdits.filter((edit) => edit.slidePartPath === partPath)) {
    applyAddConnectorEdit(root, edit);
  }
  for (const edit of paragraphTextEdits.filter((edit) => edit.handle.partPath === partPath)) {
    applyParagraphTextEdit(root, edit);
  }
  for (const edit of textRunEdits.filter((edit) => edit.handle.partPath === partPath)) {
    applyTextRunEdit(root, edit);
  }
  for (const edit of textRunPropertiesEdits.filter((edit) => edit.handle.partPath === partPath)) {
    applyTextRunPropertiesEdit(root, edit);
  }
  for (const edit of shapeTransformEdits.filter((edit) => edit.handle.partPath === partPath)) {
    applyShapeTransformEdit(root, edit);
  }
  for (const edit of deleteShapeEdits.filter((edit) => edit.handle.partPath === partPath)) {
    applyDeleteShapeEdit(root, edit);
  }
  return encodeXml(XML_DECLARATION + xmlBuilder.build(stripXmlProcessingInstruction(root)));
}

type SlideTopologyEdit =
  | PptxSourceModelAddEmptySlideFromLayoutEdit
  | PptxSourceModelDuplicateSlideEdit
  | PptxSourceModelDeleteSlideEdit;

function mergeSlideTopologyEdits(
  edits: readonly PptxSourceModelEdit[],
): readonly SlideTopologyEdit[] {
  return edits.filter(
    (edit): edit is SlideTopologyEdit =>
      edit.kind === "addEmptySlideFromLayout" ||
      edit.kind === "duplicateSlide" ||
      edit.kind === "deleteSlide",
  );
}

function serializePresentationWithSlideTopologyEdits(
  source: PptxSourceModel,
  edits: readonly SlideTopologyEdit[],
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

  for (const edit of edits) {
    if (edit.kind === "addEmptySlideFromLayout") {
      appendSlideId(sldIdLst, edit.newRelationshipId, edit.newSlideNumericId);
    } else if (edit.kind === "duplicateSlide") {
      insertSlideIdAfter(
        sldIdLst,
        edit.sourceRelationshipId,
        edit.newRelationshipId,
        edit.newSlideNumericId,
      );
    } else {
      removeSlideId(sldIdLst, edit.relationshipId);
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

function namespacedChildKey(node: XmlNode, fallback: string, local: string): string {
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    if (localName(key) === local) return key;
  }
  return fallback;
}

function namespacedAttributeKey(node: XmlNode, fallback: string, local: string): string {
  for (const key of Object.keys(node)) {
    if (!key.startsWith("@_")) continue;
    const name = key.slice(2);
    const colon = name.indexOf(":");
    if (colon !== -1 && name.slice(colon + 1) === local) return key;
  }
  return `@_${fallback}`;
}

function stripXmlProcessingInstruction(root: XmlNode): XmlNode {
  const stripped = { ...root };
  delete stripped["?xml"];
  return stripped;
}

function applyTextRunEdit(root: XmlNode, edit: PptxSourceModelTextRunEdit): void {
  const locator = parseTextRunLocator(edit.handle.nodeId);
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  const shape = locateShape(spTree, locator);
  const paragraphs = getChildArray(getChild(shape, "txBody"), "p");
  const paragraph = paragraphs[locator.paragraphIndex];
  const run = getChildArray(paragraph, "r")[locator.runIndex];

  if (run === undefined) {
    throw new Error(
      `writePptx: text run handle '${edit.handle.nodeId}' no longer matches source XML`,
    );
  }

  setChildText(run, "t", edit.text);
}

function applyTextRunPropertiesEdit(
  root: XmlNode,
  edit: PptxSourceModelTextRunPropertiesEdit,
): void {
  assertTextRunPropertiesEdit(edit);
  const run = locateTextRun(root, edit.handle.nodeId);

  if (run === undefined) {
    throw new Error(
      `writePptx: text run properties handle '${edit.handle.nodeId}' no longer matches source XML`,
    );
  }

  const set = edit.set ?? {};
  const hasSet = hasTextRunPropertiesSetValues(set);
  const existingRunProperties = getChild(run, "rPr");
  if (existingRunProperties === undefined && !hasSet) return;

  const rPr = existingRunProperties ?? ensureRunProperties(run);
  let cleared = false;
  for (const property of edit.clear ?? []) {
    cleared = clearRunProperty(rPr, property) || cleared;
  }
  if (set.bold !== undefined) rPr["@_b"] = booleanOoxmlValue(set.bold);
  if (set.italic !== undefined) rPr["@_i"] = booleanOoxmlValue(set.italic);
  if (set.underline !== undefined) rPr["@_u"] = set.underline ? "sng" : "none";
  if (set.fontSize !== undefined) rPr["@_sz"] = String(Math.round(set.fontSize * 100));
  if (set.typeface !== undefined) ensureChild(rPr, "latin")["@_typeface"] = set.typeface;
  if (set.color !== undefined) {
    replaceChild(rPr, "solidFill", {
      "a:srgbClr": {
        "@_val": set.color.hex.toUpperCase(),
      },
    });
  }
  if (!hasSet && cleared && xmlNodeIsEmpty(rPr)) deleteChild(run, "rPr");
}

function applyParagraphTextEdit(root: XmlNode, edit: PptxSourceModelParagraphTextEdit): void {
  const locator = parseParagraphLocator(edit.handle.nodeId);
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  const shape = locateShape(spTree, locator);
  const paragraphs = getChildArray(getChild(shape, "txBody"), "p");
  const paragraph = locatePhysicalParagraphForTextEdit(paragraphs, locator, edit.handle.nodeId);

  if (paragraph === undefined) {
    throw new Error(
      `writePptx: paragraph handle '${edit.handle.nodeId}' no longer matches source XML`,
    );
  }

  replaceParagraphRunsWithSingleTextRun(paragraph, edit.text);
}

function locateTextRun(
  root: XmlNode,
  nodeId: PptxSourceModelTextRunEdit["handle"]["nodeId"],
): XmlNode | undefined {
  const locator = parseTextRunLocator(nodeId);
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  const shape = locateShape(spTree, locator);
  const paragraphs = getChildArray(getChild(shape, "txBody"), "p");
  const paragraph = paragraphs[locator.paragraphIndex];
  return getChildArray(paragraph, "r")[locator.runIndex];
}

function applyShapeTransformEdit(root: XmlNode, edit: PptxSourceModelShapeTransformEdit): void {
  const locator = parseShapeLocator(edit.handle);
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  const shape = locateShapeTreeNode(spTree, locator);
  const xfrm = getShapeTransformNode(shape);

  if (xfrm === undefined) {
    throw new Error(
      `writePptx: shape transform handle '${String(
        edit.handle.nodeId ?? edit.handle.orderingSlot ?? "",
      )}' no longer matches source XML with xfrm`,
    );
  }

  const off = getChild(xfrm, "off");
  const ext = getChild(xfrm, "ext");
  if (off === undefined || ext === undefined) {
    throw new Error("writePptx: shape transform xfrm must contain off and ext");
  }

  off["@_x"] = String(edit.offsetX);
  off["@_y"] = String(edit.offsetY);
  ext["@_cx"] = String(edit.width);
  ext["@_cy"] = String(edit.height);
}

function applyAddTextBoxEdit(root: XmlNode, edit: PptxSourceModelAddTextBoxEdit): void {
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  if (spTree === undefined) {
    throw new Error(`writePptx: slide '${edit.slidePartPath}' has no spTree`);
  }
  if (locateShapeTreeNode(spTree, { nodeId: edit.shapeId }) !== undefined) {
    throw new Error(`writePptx: shape id '${edit.shapeId}' already exists in source XML`);
  }

  appendChild(spTree, "p:sp", parseShapeFragmentXml(edit.xml, "sp"));
}

function applyAddConnectorEdit(root: XmlNode, edit: PptxSourceModelAddConnectorEdit): void {
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  if (spTree === undefined) {
    throw new Error(`writePptx: slide '${edit.slidePartPath}' has no spTree`);
  }
  if (locateShapeTreeNode(spTree, { nodeId: edit.shapeId }) !== undefined) {
    throw new Error(`writePptx: shape id '${edit.shapeId}' already exists in source XML`);
  }
  if (locateShapeTreeNode(spTree, { nodeId: edit.startShapeId }) === undefined) {
    throw new Error(`writePptx: connector start shape '${edit.startShapeId}' was not found`);
  }
  if (locateShapeTreeNode(spTree, { nodeId: edit.endShapeId }) === undefined) {
    throw new Error(`writePptx: connector end shape '${edit.endShapeId}' was not found`);
  }

  appendChild(spTree, "p:cxnSp", parseShapeFragmentXml(edit.xml, "cxnSp"));
}

function applyDeleteShapeEdit(root: XmlNode, edit: PptxSourceModelDeleteShapeEdit): void {
  const locator = parseShapeLocator(edit.handle);
  const slide = getChild(root, "sld");
  const cSld = getChild(slide, "cSld");
  const spTree = getChild(cSld, "spTree");
  if (!deleteShapeXml(spTree, locator.nodeId)) {
    throw new Error(`writePptx: shape delete handle '${locator.nodeId}' no longer matches p:sp`);
  }
}

/**
 * Parse a shape XML fragment finalized at edit time and return its single root
 * element. The writer does not generate shape XML content; it only splices the
 * pre-serialized fragment into the target `p:spTree`.
 */
function parseShapeFragmentXml(xml: string, rootLocalName: "sp" | "cxnSp"): XmlNode {
  const node = getChild(parseXml(xml), rootLocalName);
  if (node === undefined) {
    throw new Error(
      `writePptx: shape edit XML fragment does not contain a '${rootLocalName}' root element`,
    );
  }
  return node;
}

interface TextRunLocator {
  readonly shapeNodeId?: string;
  readonly shapeOrderingSlot?: number;
  readonly paragraphIndex: number;
  readonly runIndex: number;
}

type ParagraphTextLocator = Omit<TextRunLocator, "runIndex">;

interface ShapeLocator {
  readonly nodeId: string;
}

function parseShapeLocator(handle: PptxSourceModelShapeTransformEdit["handle"]): ShapeLocator {
  if (handle.nodeId !== undefined) return { nodeId: String(handle.nodeId) };
  throw new Error("writePptx: shape transform edit requires nodeId in handle");
}

function parseTextRunLocator(
  nodeId: PptxSourceModelTextRunEdit["handle"]["nodeId"],
): TextRunLocator {
  const value = String(nodeId ?? "");
  const byShapeId = /^text:shape:(.+):p:(\d+):r:(\d+)$/.exec(value);
  if (byShapeId !== null) {
    return {
      shapeNodeId: byShapeId[1],
      paragraphIndex: Number(byShapeId[2]),
      runIndex: Number(byShapeId[3]),
    };
  }

  const byShapeSlot = /^text:shapeSlot:(\d+):p:(\d+):r:(\d+)$/.exec(value);
  if (byShapeSlot !== null) {
    return {
      shapeOrderingSlot: Number(byShapeSlot[1]),
      paragraphIndex: Number(byShapeSlot[2]),
      runIndex: Number(byShapeSlot[3]),
    };
  }

  throw new Error(`writePptx: unsupported text run handle '${value}'`);
}

function parseParagraphLocator(
  nodeId: PptxSourceModelParagraphTextEdit["handle"]["nodeId"],
): ParagraphTextLocator {
  const value = String(nodeId ?? "");
  const byShapeId = /^text:shape:(.+):p:(\d+)$/.exec(value);
  if (byShapeId !== null) {
    return {
      shapeNodeId: byShapeId[1],
      paragraphIndex: Number(byShapeId[2]),
    };
  }

  const byShapeSlot = /^text:shapeSlot:(\d+):p:(\d+)$/.exec(value);
  if (byShapeSlot !== null) {
    return {
      shapeOrderingSlot: Number(byShapeSlot[1]),
      paragraphIndex: Number(byShapeSlot[2]),
    };
  }

  throw new Error(`writePptx: unsupported paragraph handle '${value}'`);
}

function locateShape(
  spTree: XmlNode | undefined,
  locator: TextRunLocator | ParagraphTextLocator,
): XmlNode | undefined {
  const shapes = getChildArray(spTree, "sp");
  if (locator.shapeNodeId !== undefined) {
    return shapes.find(
      (shape) =>
        getAttr(getChild(getChild(shape, "nvSpPr"), "cNvPr"), "id") === locator.shapeNodeId,
    );
  }
  if (locator.shapeOrderingSlot === undefined) return undefined;
  return getShapeByOrderingSlot(spTree, locator.shapeOrderingSlot);
}

function locateShapeTreeNode(
  spTree: XmlNode | undefined,
  locator: ShapeLocator,
): XmlNode | undefined {
  const shapeKeys = new Set(["sp", "pic", "cxnSp", "graphicFrame", "grpSp"]);
  if (!spTree) return undefined;
  for (const key of Object.keys(spTree)) {
    if (key.startsWith("@_") || !shapeKeys.has(localName(key))) continue;
    const value = spTree[key];
    const items = Array.isArray(value) ? unsafeOoxmlBoundaryAssertion<unknown[]>(value) : [value];
    const found = items.find(
      (item) => getShapeTreeNodeId(unsafeOoxmlBoundaryAssertion<XmlNode>(item)) === locator.nodeId,
    );
    if (found !== undefined) return unsafeOoxmlBoundaryAssertion<XmlNode>(found);
  }
  return undefined;
}

function deleteShapeXml(spTree: XmlNode | undefined, nodeId: string): boolean {
  if (spTree === undefined) return false;
  const entry = Object.entries(spTree).find(
    ([key]) => !key.startsWith("@_") && localName(key) === "sp",
  );
  if (entry === undefined) return false;

  const [key, value] = entry;
  const shapes = Array.isArray(value) ? unsafeOoxmlBoundaryAssertion<unknown[]>(value) : [value];
  const nextShapes = shapes.filter(
    (shape) => getShapeTreeNodeId(unsafeOoxmlBoundaryAssertion<XmlNode>(shape)) !== nodeId,
  );
  if (nextShapes.length === shapes.length) return false;
  if (nextShapes.length === 0) delete spTree[key];
  else spTree[key] = Array.isArray(value) ? nextShapes : nextShapes[0];
  return true;
}

function getShapeByOrderingSlot(
  spTree: XmlNode | undefined,
  orderingSlot: number,
): XmlNode | undefined {
  if (!spTree) return undefined;

  let currentSlot = 0;
  for (const key of Object.keys(spTree)) {
    if (key.startsWith("@_")) continue;
    const local = localName(key);
    if (local === "nvGrpSpPr" || local === "grpSpPr") continue;

    const value = spTree[key];
    const items = Array.isArray(value) ? value : [value];
    for (const item of items) {
      if (currentSlot === orderingSlot) {
        return local === "sp" ? unsafeOoxmlBoundaryAssertion<XmlNode>(item) : undefined;
      }
      currentSlot++;
    }
  }
  return undefined;
}

function getShapeTreeNodeId(node: XmlNode): string | undefined {
  const nonVisualProperties =
    getChild(node, "nvSpPr") ??
    getChild(node, "nvPicPr") ??
    getChild(node, "nvCxnSpPr") ??
    getChild(node, "nvGrpSpPr") ??
    getChild(node, "nvGraphicFramePr");
  return getAttr(getChild(nonVisualProperties, "cNvPr"), "id");
}

function getShapeTransformNode(shape: XmlNode | undefined): XmlNode | undefined {
  if (shape === undefined) return undefined;
  return (
    getChild(getChild(shape, "spPr"), "xfrm") ??
    getChild(getChild(shape, "grpSpPr"), "xfrm") ??
    getChild(shape, "xfrm")
  );
}

function setChildText(node: XmlNode, name: string, text: string): void {
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    if (localName(key) !== name) continue;
    const value = node[key];
    if (Array.isArray(value)) {
      value[0] = textElementValue(value[0], text);
      return;
    }
    node[key] = textElementValue(value, text);
    return;
  }
  node[`a:${name}`] = textRequiresPreserve(text)
    ? { "@_xml:space": "preserve", "#text": text }
    : text;
}

function appendChild(node: XmlNode, preferredKey: string, value: XmlNode): void {
  const local = localName(preferredKey);
  const existingKey = Object.keys(node).find(
    (key) => !key.startsWith("@_") && localName(key) === local,
  );
  if (existingKey === undefined) {
    node[preferredKey] = [value];
    return;
  }

  const current = node[existingKey];
  const currentItems = Array.isArray(current)
    ? unsafeOoxmlBoundaryAssertion<unknown[]>(current)
    : [current];
  node[existingKey] = [...currentItems, value];
}

function ensureRunProperties(run: XmlNode): XmlNode {
  const existing = getChild(run, "rPr");
  if (existing !== undefined) return existing;

  const entries: [string, unknown][] = [];
  let inserted = false;
  for (const [key, value] of Object.entries(run)) {
    if (!inserted && !key.startsWith("@_")) {
      entries.push(["a:rPr", {}]);
      inserted = true;
    }
    entries.push([key, value]);
  }
  if (!inserted) entries.push(["a:rPr", {}]);
  replaceNodeEntries(run, entries);
  return getChild(run, "rPr") ?? {};
}

function clearRunProperty(
  rPr: XmlNode,
  property: NonNullable<PptxSourceModelTextRunPropertiesEdit["clear"]>[number],
): boolean {
  switch (property) {
    case "bold":
      if (rPr["@_b"] === undefined) return false;
      delete rPr["@_b"];
      return true;
    case "italic":
      if (rPr["@_i"] === undefined) return false;
      delete rPr["@_i"];
      return true;
    case "underline":
      if (rPr["@_u"] === undefined) return false;
      delete rPr["@_u"];
      return true;
    case "fontSize":
      if (rPr["@_sz"] === undefined) return false;
      delete rPr["@_sz"];
      return true;
    case "typeface": {
      const latin = getChild(rPr, "latin");
      if (latin?.["@_typeface"] === undefined) return false;
      delete latin["@_typeface"];
      if (xmlNodeIsEmpty(latin)) deleteChild(rPr, "latin");
      return true;
    }
    case "color":
      return deleteChild(rPr, "solidFill");
  }
}

function ensureChild(node: XmlNode, name: string): XmlNode {
  const existing = getChild(node, name);
  if (existing !== undefined) return existing;
  node[`a:${name}`] = {};
  return unsafeOoxmlBoundaryAssertion<XmlNode>(node[`a:${name}`]);
}

function replaceChild(node: XmlNode, name: string, value: XmlNode): void {
  const entries: [string, unknown][] = [];
  let replaced = false;
  for (const [key, entryValue] of Object.entries(node)) {
    if (!key.startsWith("@_") && localName(key) === name) {
      if (!replaced) entries.push([key, value]);
      replaced = true;
      continue;
    }
    entries.push([key, entryValue]);
  }
  if (!replaced) entries.push([`a:${name}`, value]);
  replaceNodeEntries(node, entries);
}

function deleteChild(node: XmlNode, name: string): boolean {
  let deleted = false;
  replaceNodeEntries(
    node,
    Object.entries(node).filter(([key]) => {
      const keep = key.startsWith("@_") || localName(key) !== name;
      if (!keep) deleted = true;
      return keep;
    }),
  );
  return deleted;
}

function booleanOoxmlValue(value: boolean): string {
  return value ? "1" : "0";
}

function hasTextRunPropertiesSetValues(properties: EditableTextRunProperties): boolean {
  return (
    properties.bold !== undefined ||
    properties.italic !== undefined ||
    properties.underline !== undefined ||
    properties.fontSize !== undefined ||
    properties.color !== undefined ||
    properties.typeface !== undefined
  );
}

function assertTextRunPropertiesEdit(edit: PptxSourceModelTextRunPropertiesEdit): void {
  const clear = edit.clear ?? [];
  if (!hasTextRunPropertiesSetValues(edit.set ?? {}) && clear.length === 0) {
    throw new Error("writePptx: text run properties edit must set or clear at least one property");
  }
}

function xmlNodeIsEmpty(node: XmlNode): boolean {
  return Object.keys(node).length === 0;
}

function replaceParagraphRunsWithSingleTextRun(paragraph: XmlNode, text: string): void {
  const firstRunProperties = getChild(getFirstRunLikeNode(paragraph), "rPr");
  const replacementRun: XmlNode = {
    ...(firstRunProperties !== undefined ? { "a:rPr": cloneXmlNode(firstRunProperties) } : {}),
    "a:t": textRequiresPreserve(text) ? { "@_xml:space": "preserve", "#text": text } : text,
  };
  const attrs: [string, unknown][] = [];
  const paragraphProperties: [string, unknown][] = [];
  const middleChildren: [string, unknown][] = [];
  const endProperties: [string, unknown][] = [];

  for (const [key, value] of Object.entries(paragraph)) {
    if (key.startsWith("@_")) {
      attrs.push([key, value]);
      continue;
    }

    const local = localName(key);
    if (isRunLikeLocalName(local)) continue;
    if (local === "pPr") paragraphProperties.push([key, value]);
    else if (local === "endParaRPr") endProperties.push([key, value]);
    else middleChildren.push([key, value]);
  }

  replaceNodeEntries(paragraph, [
    ...attrs,
    ...paragraphProperties,
    ["a:r", replacementRun],
    ...middleChildren,
    ...endProperties,
  ]);
}

function getFirstRunLikeNode(paragraph: XmlNode): XmlNode | undefined {
  for (const key of Object.keys(paragraph)) {
    if (key.startsWith("@_")) continue;
    if (!isRunLikeLocalName(localName(key))) continue;
    const value = paragraph[key];
    return Array.isArray(value)
      ? unsafeOoxmlBoundaryAssertion<XmlNode | undefined>(value[0])
      : unsafeOoxmlBoundaryAssertion<XmlNode | undefined>(value);
  }
  return undefined;
}

function isRunLikeLocalName(name: string): boolean {
  return name === "r" || name === "fld" || name === "br";
}

function locatePhysicalParagraphForTextEdit(
  paragraphs: readonly XmlNode[],
  locator: ParagraphTextLocator,
  handleNodeId: PptxSourceModelParagraphTextEdit["handle"]["nodeId"],
): XmlNode | undefined {
  let logicalParagraphIndex = 0;
  for (const paragraph of paragraphs) {
    const logicalCount = getLogicalParagraphCount(paragraph);
    if (
      locator.paragraphIndex >= logicalParagraphIndex &&
      locator.paragraphIndex < logicalParagraphIndex + logicalCount
    ) {
      if (logicalCount > 1) {
        throw new Error(
          `writePptx: paragraph handle '${handleNodeId}' references an interleaved bullet paragraph split by the reader; paragraph replacement is not supported for this source XML`,
        );
      }
      return paragraph;
    }
    logicalParagraphIndex += logicalCount;
  }
  return undefined;
}

function getLogicalParagraphCount(paragraph: XmlNode): number {
  const bulletParagraphProperties = getChildArray(paragraph, "pPr").filter(
    (properties) =>
      getChild(properties, "buChar") !== undefined ||
      getChild(properties, "buAutoNum") !== undefined,
  );
  return Math.max(1, bulletParagraphProperties.length);
}

function replaceNodeEntries(node: XmlNode, entries: readonly [string, unknown][]): void {
  for (const key of Object.keys(node)) delete node[key];
  for (const [key, value] of entries) node[key] = value;
}

function cloneXmlNode(node: XmlNode): XmlNode {
  return Object.fromEntries(
    Object.entries(node).map(([key, value]) => [key, cloneXmlValue(value)]),
  );
}

function cloneXmlValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(cloneXmlValue);
  if (typeof value === "object" && value !== null) {
    return cloneXmlNode(unsafeOoxmlBoundaryAssertion<XmlNode>(value));
  }
  return value;
}

function textElementValue(existing: unknown, text: string): unknown {
  if (typeof existing === "object" && existing !== null && !Array.isArray(existing)) {
    const next: XmlNode = { ...unsafeOoxmlBoundaryAssertion<XmlNode>(existing), "#text": text };
    if (textRequiresPreserve(text)) next["@_xml:space"] = "preserve";
    else delete next["@_xml:space"];
    return next;
  }
  return textRequiresPreserve(text) ? { "@_xml:space": "preserve", "#text": text } : text;
}

function textRequiresPreserve(text: string): boolean {
  return text.startsWith(" ") || text.endsWith(" ");
}

function validateEdits(
  textRunEdits: readonly PptxSourceModelTextRunEdit[],
  textRunPropertiesEdits: readonly PptxSourceModelTextRunPropertiesEdit[],
  paragraphTextEdits: readonly PptxSourceModelParagraphTextEdit[],
  shapeTransformEdits: readonly PptxSourceModelShapeTransformEdit[],
  addTextBoxEdits: readonly PptxSourceModelAddTextBoxEdit[],
  addConnectorEdits: readonly PptxSourceModelAddConnectorEdit[],
  deleteShapeEdits: readonly PptxSourceModelDeleteShapeEdit[],
): void {
  const addedShapeKeys = new Set<string>();
  for (const edit of [...addTextBoxEdits, ...addConnectorEdits]) {
    const key = [edit.slidePartPath, edit.shapeId].join("\u0000");
    if (addedShapeKeys.has(key)) {
      throw new Error(`writePptx: conflicting shape additions for shape id '${edit.shapeId}'`);
    }
    addedShapeKeys.add(key);
  }

  const runKeys = new Set<string>();
  for (const edit of textRunEdits) {
    const key = editHandleNodeKey(edit);
    if (runKeys.has(key)) {
      throw new Error(`writePptx: conflicting text run edits for handle '${edit.handle.nodeId}'`);
    }
    runKeys.add(key);
  }

  const paragraphKeys = new Set<string>();
  for (const edit of paragraphTextEdits) {
    const key = editHandleNodeKey(edit);
    if (paragraphKeys.has(key)) {
      throw new Error(
        `writePptx: conflicting paragraph text edits for handle '${edit.handle.nodeId}'`,
      );
    }
    paragraphKeys.add(key);
  }

  for (const runEdit of textRunEdits) {
    const paragraphKey = textRunParagraphEditKey(runEdit);
    if (paragraphKey !== undefined && paragraphKeys.has(paragraphKey)) {
      throw new Error(
        `writePptx: conflicting text run and paragraph edits for handle '${runEdit.handle.nodeId}'`,
      );
    }
  }
  for (const runPropertiesEdit of textRunPropertiesEdits) {
    const paragraphKey = textRunParagraphEditKey(runPropertiesEdit);
    if (paragraphKey !== undefined && paragraphKeys.has(paragraphKey)) {
      throw new Error(
        `writePptx: conflicting text run properties and paragraph edits for handle '${runPropertiesEdit.handle.nodeId}'`,
      );
    }
  }

  const shapeKeys = new Set<string>();
  for (const edit of shapeTransformEdits) {
    const key = editHandleNodeKey(edit);
    if (shapeKeys.has(key)) {
      throw new Error(
        `writePptx: conflicting shape transform edits for handle '${String(edit.handle.nodeId)}'`,
      );
    }
    shapeKeys.add(key);
  }

  const deletedShapeKeys = new Set<string>();
  for (const edit of deleteShapeEdits) {
    const key = editHandleNodeKey(edit);
    if (deletedShapeKeys.has(key)) {
      throw new Error(
        `writePptx: conflicting shape delete edits for handle '${String(edit.handle.nodeId)}'`,
      );
    }
    deletedShapeKeys.add(key);
  }
}

function editHandleNodeKey(edit: {
  readonly handle: PptxSourceModelTextRunEdit["handle"];
}): string {
  return [edit.handle.partPath, edit.handle.nodeId ?? "", edit.handle.relationshipId ?? ""].join(
    "\u0000",
  );
}

function textRunParagraphEditKey(
  edit: PptxSourceModelTextRunEdit | PptxSourceModelTextRunPropertiesEdit,
): string | undefined {
  const nodeId = String(edit.handle.nodeId ?? "");
  const byShapeId = /^(text:shape:.+:p:\d+):r:\d+$/.exec(nodeId);
  const byShapeSlot = /^(text:shapeSlot:\d+:p:\d+):r:\d+$/.exec(nodeId);
  const paragraphNodeId = byShapeId?.[1] ?? byShapeSlot?.[1];
  if (paragraphNodeId === undefined) return undefined;
  return [edit.handle.partPath, paragraphNodeId, edit.handle.relationshipId ?? ""].join("\u0000");
}

function serializeContentTypes(
  contentTypes: PptxSourceModel["packageGraph"]["contentTypes"],
): string {
  const defaults = contentTypes.defaults
    .map(
      (entry) =>
        `<Default Extension="${escapeAttribute(entry.extension)}" ` +
        `ContentType="${escapeAttribute(entry.contentType)}"/>`,
    )
    .join("");
  const overrides = contentTypes.overrides
    .map(
      (entry) =>
        `<Override PartName="/${escapeAttribute(entry.partName)}" ` +
        `ContentType="${escapeAttribute(entry.contentType)}"/>`,
    )
    .join("");
  return (
    XML_DECLARATION +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    defaults +
    overrides +
    `</Types>`
  );
}

function serializeRelationships(partRelationships: PartRelationships): string {
  const relationships = partRelationships.relationships
    .map((relationship) => {
      const targetMode =
        relationship.targetMode === undefined
          ? ""
          : ` TargetMode="${escapeAttribute(relationship.targetMode)}"`;
      return (
        `<Relationship Id="${escapeAttribute(relationship.id)}" ` +
        `Type="${escapeAttribute(relationship.type)}" ` +
        `Target="${escapeAttribute(relationship.target)}"${targetMode}/>`
      );
    })
    .join("");
  return (
    XML_DECLARATION +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    relationships +
    `</Relationships>`
  );
}

function serializeRawPackagePart(rawPart: RawPackagePart): Uint8Array {
  if (rawPart.kind === "binary") return rawPart.bytes;
  return encodeXml(XML_DECLARATION + serializeRawNode(rawPart.xml));
}

function serializeRawNode(node: RawOoxmlNode): string {
  const attributes =
    node.attributes === undefined
      ? ""
      : Object.entries(node.attributes)
          .map(([name, value]) => ` ${name}="${escapeAttribute(value)}"`)
          .join("");
  const text = node.text === undefined ? "" : escapeText(node.text);
  const children = node.children?.map((child) => serializeRawNode(child)).join("") ?? "";
  if (text !== "" && children !== "") {
    throw new Error(
      `writePptx: raw XML part '${node.name}' contains mixed text/element content; ` +
        "ordered mixed-content serialization is not implemented in the no-edit writer",
    );
  }
  if (text === "" && children === "") return `<${node.name}${attributes}/>`;
  return `<${node.name}${attributes}>${text}${children}</${node.name}>`;
}

function encodeXml(xml: string): Uint8Array {
  return textEncoder.encode(xml);
}

function escapeAttribute(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeText(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}
