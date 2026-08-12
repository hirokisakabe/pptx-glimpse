import { localName, type XmlNode } from "./xml.js";

const SHAPE_TREE_CONTAINER_PROPERTY_LOCAL_NAMES = ["nvGrpSpPr", "grpSpPr", "extLst"] as const;

/** Resolve the authored child keys that belong to the shape-tree container rather than z-order. */
export function getShapeTreeContainerPropertyKeys(node: XmlNode): ReadonlySet<string> {
  const containerPrefix = shapeTreeContainerPrefix(node);
  if (containerPrefix === undefined) return new Set();
  return new Set(
    SHAPE_TREE_CONTAINER_PROPERTY_LOCAL_NAMES.flatMap((name) => {
      const key = qualifiedChildKey(node, containerPrefix, name);
      return key === undefined ? [] : [key];
    }),
  );
}

/** Classify element children that participate in shape-tree z-order and contiguity. */
export function isShapeTreeZOrderChildKey(
  key: string,
  containerPropertyKeys: ReadonlySet<string>,
): boolean {
  return !key.startsWith("#") && !containerPropertyKeys.has(key);
}

function shapeTreeContainerPrefix(node: XmlNode): string | undefined {
  const candidates = childKeys(node, "grpSpPr");
  const preferred =
    candidates.find((key) => key === "p:grpSpPr") ??
    candidates.find((key) => key === "grpSpPr") ??
    (candidates.length === 1 ? candidates[0] : undefined);
  if (preferred === undefined) return undefined;
  const colon = preferred.indexOf(":");
  return colon === -1 ? "" : preferred.slice(0, colon);
}

function qualifiedChildKey(
  node: XmlNode,
  prefix: string,
  childLocalName: string,
): string | undefined {
  const qualifiedName = prefix === "" ? childLocalName : `${prefix}:${childLocalName}`;
  return childKeys(node, childLocalName).find((key) => key === qualifiedName);
}

function childKeys(node: XmlNode, childLocalName: string): string[] {
  return Object.keys(node).filter(
    (key) => !key.startsWith("@_") && localName(key) === childLocalName,
  );
}
