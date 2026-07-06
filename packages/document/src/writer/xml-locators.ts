import { getAttr, getChild, getChildArray, localName, type XmlNode } from "../reader/xml.js";
import type {
  PptxSourceModelParagraphTextEdit,
  PptxSourceModelShapeTransformEdit,
  PptxSourceModelTextRunEdit,
} from "../source/index.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";

export interface TextRunLocator {
  readonly shapeNodeId?: string;
  readonly shapeOrderingSlot?: number;
  readonly paragraphIndex: number;
  readonly runIndex: number;
}

export type ParagraphTextLocator = Omit<TextRunLocator, "runIndex">;

interface ShapeLocator {
  readonly nodeId: string;
}

export function parseShapeLocator(
  handle: PptxSourceModelShapeTransformEdit["handle"],
): ShapeLocator {
  if (handle.nodeId !== undefined) return { nodeId: String(handle.nodeId) };
  throw new Error("writePptx: shape transform edit requires nodeId in handle");
}

export function parseTextRunLocator(
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

export function parseParagraphLocator(
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

export function locateShape(
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

export function locateShapeTreeNode(
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

export function deleteShapeXml(spTree: XmlNode | undefined, nodeId: string): boolean {
  if (spTree === undefined) return false;
  const entry = Object.entries(spTree).find(
    ([key, value]) =>
      !key.startsWith("@_") &&
      (localName(key) === "sp" || localName(key) === "cxnSp") &&
      getShapeTreeNodes(value).some((shape) => getShapeTreeNodeId(shape) === nodeId),
  );
  if (entry === undefined) return false;

  const [key, value] = entry;
  const shapes = getShapeTreeNodes(value);
  const nextShapes = shapes.filter((shape) => getShapeTreeNodeId(shape) !== nodeId);
  if (nextShapes.length === shapes.length) return false;
  if (nextShapes.length === 0) delete spTree[key];
  else spTree[key] = Array.isArray(value) ? nextShapes : nextShapes[0];
  return true;
}

function getShapeTreeNodes(value: unknown): XmlNode[] {
  const items = Array.isArray(value) ? unsafeOoxmlBoundaryAssertion<unknown[]>(value) : [value];
  return items.map((item) => unsafeOoxmlBoundaryAssertion<XmlNode>(item));
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

export function getShapeTransformNode(shape: XmlNode | undefined): XmlNode | undefined {
  if (shape === undefined) return undefined;
  return (
    getChild(getChild(shape, "spPr"), "xfrm") ??
    getChild(getChild(shape, "grpSpPr"), "xfrm") ??
    getChild(shape, "xfrm")
  );
}
