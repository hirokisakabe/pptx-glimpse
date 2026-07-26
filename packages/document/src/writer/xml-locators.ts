import { getAttr, getChild, getChildArray, localName, type XmlNode } from "../reader/xml.js";
import type {
  PptxSourceModelParagraphTextEdit,
  PptxSourceModelTextRunEdit,
  SourceHandle,
} from "../source/index.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";

export interface TextRunLocator {
  readonly ownerKind: "shape" | "table";
  readonly shapeNodeId?: string;
  readonly shapeOrderingSlot?: number;
  readonly tableNodeId?: string;
  readonly rowIndex?: number;
  readonly cellIndex?: number;
  readonly paragraphIndex: number;
  readonly runIndex: number;
}

export type ParagraphTextLocator = Omit<TextRunLocator, "runIndex">;

interface ShapeLocator {
  readonly nodeId: string;
}

type ShapeTreeNodeKind = "sp" | "pic" | "cxnSp" | "graphicFrame" | "grpSp";

interface ShapeTreeNodeLocation {
  readonly node: XmlNode;
  readonly nodeKind: ShapeTreeNodeKind;
  readonly parentContainer: XmlNode;
  readonly nested: boolean;
}

export function parseShapeLocator(handle: SourceHandle, editName = "shape edit"): ShapeLocator {
  if (handle.nodeId !== undefined) return { nodeId: String(handle.nodeId) };
  throw new Error(`writePptx: ${editName} requires nodeId in handle`);
}

export function parseTextRunLocator(
  nodeId: PptxSourceModelTextRunEdit["handle"]["nodeId"],
): TextRunLocator {
  const value = String(nodeId ?? "");
  const byShapeId = /^text:shape:(.+):p:(\d+):r:(\d+)$/.exec(value);
  if (byShapeId !== null) {
    return {
      ownerKind: "shape",
      shapeNodeId: byShapeId[1],
      paragraphIndex: Number(byShapeId[2]),
      runIndex: Number(byShapeId[3]),
    };
  }

  const byShapeSlot = /^text:shapeSlot:(\d+):p:(\d+):r:(\d+)$/.exec(value);
  if (byShapeSlot !== null) {
    return {
      ownerKind: "shape",
      shapeOrderingSlot: Number(byShapeSlot[1]),
      paragraphIndex: Number(byShapeSlot[2]),
      runIndex: Number(byShapeSlot[3]),
    };
  }

  const byTableId = /^text:table:(.+):row:(\d+):cell:(\d+):p:(\d+):r:(\d+)$/.exec(value);
  if (byTableId !== null) {
    return {
      ownerKind: "table",
      tableNodeId: byTableId[1],
      rowIndex: Number(byTableId[2]),
      cellIndex: Number(byTableId[3]),
      paragraphIndex: Number(byTableId[4]),
      runIndex: Number(byTableId[5]),
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
      ownerKind: "shape",
      shapeNodeId: byShapeId[1],
      paragraphIndex: Number(byShapeId[2]),
    };
  }

  const byShapeSlot = /^text:shapeSlot:(\d+):p:(\d+)$/.exec(value);
  if (byShapeSlot !== null) {
    return {
      ownerKind: "shape",
      shapeOrderingSlot: Number(byShapeSlot[1]),
      paragraphIndex: Number(byShapeSlot[2]),
    };
  }

  const byTableId = /^text:table:(.+):row:(\d+):cell:(\d+):p:(\d+)$/.exec(value);
  if (byTableId !== null) {
    return {
      ownerKind: "table",
      tableNodeId: byTableId[1],
      rowIndex: Number(byTableId[2]),
      cellIndex: Number(byTableId[3]),
      paragraphIndex: Number(byTableId[4]),
    };
  }

  throw new Error(`writePptx: unsupported paragraph handle '${value}'`);
}

export function locateShape(
  spTree: XmlNode | undefined,
  locator: TextRunLocator | ParagraphTextLocator,
): XmlNode | undefined {
  if (locator.ownerKind !== "shape") return undefined;
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

export function locateTable(
  spTree: XmlNode | undefined,
  locator: TextRunLocator | ParagraphTextLocator,
): XmlNode | undefined {
  if (locator.ownerKind !== "table") return undefined;
  const frames = getChildArray(spTree, "graphicFrame");
  if (locator.tableNodeId !== undefined) {
    const frame = frames.find(
      (frame) =>
        getAttr(getChild(getChild(frame, "nvGraphicFramePr"), "cNvPr"), "id") ===
        locator.tableNodeId,
    );
    if (frame !== undefined) return frame;
    for (const group of getChildArray(spTree, "grpSp")) {
      const nested = locateTable(group, locator);
      if (nested !== undefined) return nested;
    }
    for (const alternateContent of getChildArray(spTree, "AlternateContent")) {
      for (const branchName of ["Choice", "Fallback"]) {
        for (const branch of getChildArray(alternateContent, branchName)) {
          const nested = locateTable(branch, locator);
          if (nested !== undefined) return nested;
        }
      }
    }
  }
  return undefined;
}

export function locateShapeTreeNode(
  spTree: XmlNode | undefined,
  locator: ShapeLocator,
): XmlNode | undefined {
  return locateShapeTreeNodeLocation(spTree, locator)?.node;
}

/**
 * Resolves a drawing id across the complete part-local shape tree.
 *
 * The returned parent is the immediate `p:spTree` or `p:grpSp` container so future topology
 * edits can patch the correct subtree. Ambiguous/non-conforming ids and AlternateContent
 * targets are rejected instead of selecting an arbitrary branch.
 */
export function locateShapeTreeNodeLocation(
  spTree: XmlNode | undefined,
  locator: ShapeLocator,
): ShapeTreeNodeLocation | undefined {
  if (spTree === undefined) return undefined;
  const matches: LocatedShapeTreeNode[] = [];
  collectShapeTreeNodeLocations(spTree, locator.nodeId, 0, false, matches);
  if (matches.some((match) => match.insideAlternateContent)) {
    throw new Error(
      `writePptx: shape id '${locator.nodeId}' inside mc:AlternateContent is not supported`,
    );
  }
  if (matches.length > 1) {
    throw new Error(
      `writePptx: duplicate shape id '${locator.nodeId}' in the same drawing part is not supported`,
    );
  }
  const match = matches[0];
  if (match === undefined) return undefined;
  return {
    node: match.node,
    nodeKind: match.nodeKind,
    parentContainer: match.parentContainer,
    nested: match.nested,
  };
}

export function deleteShapeXml(spTree: XmlNode | undefined, nodeId: string): boolean {
  if (spTree === undefined) return false;
  const location = locateShapeTreeNodeLocation(spTree, { nodeId });
  if (
    location === undefined ||
    location.nested ||
    (location.nodeKind !== "sp" && location.nodeKind !== "cxnSp")
  ) {
    return false;
  }
  const entry = Object.entries(location.parentContainer).find(
    ([key, value]) =>
      !key.startsWith("@_") &&
      localName(key) === location.nodeKind &&
      getShapeTreeNodes(value).some((shape) => shape === location.node),
  );
  if (entry === undefined) return false;

  const [key, value] = entry;
  const shapes = getShapeTreeNodes(value);
  const nextShapes = shapes.filter((shape) => shape !== location.node);
  if (nextShapes.length === shapes.length) return false;
  if (nextShapes.length === 0) delete location.parentContainer[key];
  else location.parentContainer[key] = Array.isArray(value) ? nextShapes : nextShapes[0];
  return true;
}

interface LocatedShapeTreeNode extends ShapeTreeNodeLocation {
  readonly insideAlternateContent: boolean;
}

function collectShapeTreeNodeLocations(
  container: XmlNode,
  nodeId: string,
  depth: number,
  insideAlternateContent: boolean,
  matches: LocatedShapeTreeNode[],
): void {
  for (const [key, value] of Object.entries(container)) {
    if (key.startsWith("@_")) continue;
    const keyLocalName = localName(key);
    const items = getShapeTreeNodes(value);
    if (isShapeTreeNodeKind(keyLocalName)) {
      for (const node of items) {
        if (getShapeTreeNodeId(node) === nodeId) {
          matches.push({
            node,
            nodeKind: keyLocalName,
            parentContainer: container,
            nested: depth > 0,
            insideAlternateContent,
          });
        }
        if (keyLocalName === "grpSp") {
          collectShapeTreeNodeLocations(node, nodeId, depth + 1, insideAlternateContent, matches);
        }
      }
      continue;
    }
    if (keyLocalName === "AlternateContent" || insideAlternateContent) {
      for (const node of items) {
        collectShapeTreeNodeLocations(node, nodeId, depth, true, matches);
      }
    }
  }
}

function isShapeTreeNodeKind(value: string): value is ShapeTreeNodeKind {
  return (
    value === "sp" ||
    value === "pic" ||
    value === "cxnSp" ||
    value === "graphicFrame" ||
    value === "grpSp"
  );
}

function getShapeTreeNodes(value: unknown): XmlNode[] {
  const items = Array.isArray(value) ? unsafeOoxmlBoundaryAssertion<unknown[]>(value) : [value];
  return items
    .filter((item) => typeof item === "object" && item !== null && !Array.isArray(item))
    .map((item) => unsafeOoxmlBoundaryAssertion<XmlNode>(item));
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
