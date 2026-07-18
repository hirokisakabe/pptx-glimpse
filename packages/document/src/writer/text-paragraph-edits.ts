import { getChild, getChildArray, localName, type XmlNode } from "../reader/xml.js";
import type {
  EditableParagraphProperties,
  EditableTextRunProperties,
  PptxSourceModelParagraphPropertiesEdit,
  PptxSourceModelParagraphTextEdit,
  PptxSourceModelTextRunEdit,
  PptxSourceModelTextRunPropertiesEdit,
} from "../source/index.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import {
  locateShape,
  type ParagraphTextLocator,
  parseParagraphLocator,
  parseTextRunLocator,
} from "./xml-locators.js";
import {
  cloneXmlNode,
  deleteChild,
  ensureChild,
  replaceChild,
  replaceNodeEntries,
  setChildText,
  textRequiresPreserve,
  xmlNodeIsEmpty,
} from "./xml-node-utils.js";

export function applyTextRunEdit(root: XmlNode, edit: PptxSourceModelTextRunEdit): void {
  const locator = parseTextRunLocator(edit.handle.nodeId);
  const shape = locateTextShape(root, locator);
  const paragraph = getChildArray(getChild(shape, "txBody"), "p")[locator.paragraphIndex];
  const run = getChildArray(paragraph, "r")[locator.runIndex];
  if (run === undefined) {
    throw new Error(
      `writePptx: text run handle '${edit.handle.nodeId}' no longer matches source XML`,
    );
  }
  setChildText(run, "t", edit.text);
}

export function applyTextRunPropertiesEdit(
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
  for (const property of edit.clear ?? []) cleared = clearRunProperty(rPr, property) || cleared;
  if (set.bold !== undefined) rPr["@_b"] = booleanOoxmlValue(set.bold);
  if (set.italic !== undefined) rPr["@_i"] = booleanOoxmlValue(set.italic);
  if (set.underline !== undefined) rPr["@_u"] = set.underline ? "sng" : "none";
  if (set.fontSize !== undefined) rPr["@_sz"] = String(Math.round(set.fontSize * 100));
  if (set.typeface !== undefined) ensureChild(rPr, "latin")["@_typeface"] = set.typeface;
  if (set.color !== undefined) {
    replaceChild(rPr, "solidFill", {
      "a:srgbClr": { "@_val": set.color.hex.toUpperCase() },
    });
  }
  if (!hasSet && cleared && xmlNodeIsEmpty(rPr)) deleteChild(run, "rPr");
}

export function applyParagraphTextEdit(
  root: XmlNode,
  edit: PptxSourceModelParagraphTextEdit,
): void {
  const locator = parseParagraphLocator(edit.handle.nodeId);
  const shape = locateTextShape(root, locator);
  const paragraphs = getChildArray(getChild(shape, "txBody"), "p");
  const paragraph = locatePhysicalParagraphForTextEdit(paragraphs, locator, edit.handle.nodeId);
  if (paragraph === undefined) {
    throw new Error(
      `writePptx: paragraph handle '${edit.handle.nodeId}' no longer matches source XML`,
    );
  }
  replaceParagraphRunsWithSingleTextRun(paragraph, edit.text);
}

export function applyParagraphPropertiesEdit(
  root: XmlNode,
  edit: PptxSourceModelParagraphPropertiesEdit,
): void {
  assertParagraphPropertiesEdit(edit);
  const locator = parseParagraphLocator(edit.handle.nodeId);
  const shape = locateTextShape(root, locator);
  const paragraphs = getChildArray(getChild(shape, "txBody"), "p");
  const target = locateParagraphPropertiesForEdit(paragraphs, locator);
  if (target === undefined) {
    throw new Error(
      `writePptx: paragraph properties handle '${edit.handle.nodeId}' no longer matches source XML`,
    );
  }

  const set = edit.set ?? {};
  const hasSet = hasParagraphPropertiesSetValues(set);
  const existingParagraphProperties = target.properties;
  if (existingParagraphProperties === undefined && !hasSet) return;

  const pPr = existingParagraphProperties ?? ensureParagraphProperties(target.paragraph);
  let cleared = false;
  for (const property of edit.clear ?? []) {
    cleared = clearParagraphProperty(pPr, property) || cleared;
  }
  if (set.align !== undefined) pPr["@_algn"] = paragraphAlignOoxmlValue(set.align);
  if (set.level !== undefined) pPr["@_lvl"] = String(set.level);
  if (set.bullet !== undefined) setParagraphBullet(pPr, set.bullet);
  if (!hasSet && cleared && xmlNodeIsEmpty(pPr)) deleteParagraphProperties(target.paragraph, pPr);
}

function locateTextShape(
  root: XmlNode,
  locator: ReturnType<typeof parseTextRunLocator> | ReturnType<typeof parseParagraphLocator>,
): XmlNode | undefined {
  const slide = getChild(root, "sld");
  return locateShape(getChild(getChild(slide, "cSld"), "spTree"), locator);
}

function locateTextRun(
  root: XmlNode,
  nodeId: PptxSourceModelTextRunEdit["handle"]["nodeId"],
): XmlNode | undefined {
  const locator = parseTextRunLocator(nodeId);
  const shape = locateTextShape(root, locator);
  const paragraph = getChildArray(getChild(shape, "txBody"), "p")[locator.paragraphIndex];
  return getChildArray(paragraph, "r")[locator.runIndex];
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

function hasParagraphPropertiesSetValues(properties: EditableParagraphProperties): boolean {
  return (
    properties.align !== undefined ||
    properties.level !== undefined ||
    properties.bullet !== undefined
  );
}

function assertTextRunPropertiesEdit(edit: PptxSourceModelTextRunPropertiesEdit): void {
  if (!hasTextRunPropertiesSetValues(edit.set ?? {}) && (edit.clear ?? []).length === 0) {
    throw new Error("writePptx: text run properties edit must set or clear at least one property");
  }
}

function assertParagraphPropertiesEdit(edit: PptxSourceModelParagraphPropertiesEdit): void {
  if (!hasParagraphPropertiesSetValues(edit.set ?? {}) && (edit.clear ?? []).length === 0) {
    throw new Error("writePptx: paragraph properties edit must set or clear at least one property");
  }
}

interface ParagraphPropertiesTarget {
  readonly paragraph: XmlNode;
  readonly properties?: XmlNode;
}

function locateParagraphPropertiesForEdit(
  paragraphs: readonly XmlNode[],
  locator: ParagraphTextLocator,
): ParagraphPropertiesTarget | undefined {
  let logicalParagraphIndex = 0;
  for (const paragraph of paragraphs) {
    const logicalCount = getLogicalParagraphCount(paragraph);
    if (
      locator.paragraphIndex >= logicalParagraphIndex &&
      locator.paragraphIndex < logicalParagraphIndex + logicalCount
    ) {
      if (logicalCount === 1) return { paragraph, properties: getChild(paragraph, "pPr") };
      const relativeIndex = locator.paragraphIndex - logicalParagraphIndex;
      return { paragraph, properties: getBulletParagraphProperties(paragraph)[relativeIndex] };
    }
    logicalParagraphIndex += logicalCount;
  }
  return undefined;
}

function getBulletParagraphProperties(paragraph: XmlNode): readonly XmlNode[] {
  return getChildArray(paragraph, "pPr").filter(
    (properties) =>
      getChild(properties, "buChar") !== undefined ||
      getChild(properties, "buAutoNum") !== undefined,
  );
}

function ensureParagraphProperties(paragraph: XmlNode): XmlNode {
  const existing = getChild(paragraph, "pPr");
  if (existing !== undefined) return existing;
  const entries: [string, unknown][] = [];
  let inserted = false;
  for (const [key, value] of Object.entries(paragraph)) {
    if (!inserted && !key.startsWith("@_")) {
      entries.push(["a:pPr", {}]);
      inserted = true;
    }
    entries.push([key, value]);
  }
  if (!inserted) entries.push(["a:pPr", {}]);
  replaceNodeEntries(paragraph, entries);
  return getChild(paragraph, "pPr") ?? {};
}

function deleteParagraphProperties(paragraph: XmlNode, pPr: XmlNode): void {
  if (getChild(paragraph, "pPr") === pPr) deleteChild(paragraph, "pPr");
}

function clearParagraphProperty(
  pPr: XmlNode,
  property: NonNullable<PptxSourceModelParagraphPropertiesEdit["clear"]>[number],
): boolean {
  switch (property) {
    case "align":
      if (pPr["@_algn"] === undefined) return false;
      delete pPr["@_algn"];
      return true;
    case "level":
      if (pPr["@_lvl"] === undefined) return false;
      delete pPr["@_lvl"];
      return true;
    case "bullet":
      return deleteParagraphBullet(pPr);
  }
}

function deleteParagraphBullet(pPr: XmlNode): boolean {
  const deletedNone = deleteChild(pPr, "buNone");
  const deletedChar = deleteChild(pPr, "buChar");
  const deletedAutoNum = deleteChild(pPr, "buAutoNum");
  return deletedNone || deletedChar || deletedAutoNum;
}

function setParagraphBullet(
  pPr: XmlNode,
  bullet: NonNullable<EditableParagraphProperties["bullet"]>,
): void {
  deleteParagraphBullet(pPr);
  if (bullet.type === "none") return replaceChild(pPr, "buNone", {});
  if (bullet.type === "char") return replaceChild(pPr, "buChar", { "@_char": bullet.char });
  replaceChild(pPr, "buAutoNum", {
    "@_type": bullet.scheme,
    "@_startAt": String(bullet.startAt),
  });
}

function paragraphAlignOoxmlValue(
  align: NonNullable<EditableParagraphProperties["align"]>,
): string {
  if (align === "center") return "ctr";
  if (align === "right") return "r";
  if (align === "justify") return "just";
  return "l";
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
    if (key.startsWith("@_") || !isRunLikeLocalName(localName(key))) continue;
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
  return Math.max(1, getBulletParagraphProperties(paragraph).length);
}
