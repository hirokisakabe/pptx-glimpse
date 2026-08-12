import { localName, type XmlNode } from "./xml.js";

const SHAPE_TREE_CONTAINER_PROPERTY_LOCAL_NAMES = ["nvGrpSpPr", "grpSpPr", "extLst"] as const;

/** Resolve the authored child keys that belong to the shape-tree container rather than z-order. */
export function getShapeTreeContainerPropertyKeys(node: XmlNode): ReadonlySet<string> {
  return new Set(
    SHAPE_TREE_CONTAINER_PROPERTY_LOCAL_NAMES.flatMap((name) => {
      const key = preferredQualifiedChildKey(node, name);
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

function preferredQualifiedChildKey(node: XmlNode, childLocalName: string): string | undefined {
  const candidates = Object.keys(node).filter(
    (key) => !key.startsWith("@_") && localName(key) === childLocalName,
  );
  return (
    candidates.find((key) => key === `p:${childLocalName}`) ??
    candidates.find((key) => key === childLocalName) ??
    (candidates.length === 1 ? candidates[0] : undefined)
  );
}
