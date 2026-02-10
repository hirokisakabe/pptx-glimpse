import { parseXml } from "./xml-parser.js";

const WARN_PREFIX = "[pptx-glimpse]";

export interface Relationship {
  id: string;
  type: string;
  target: string;
}

export function parseRelationships(xml: string): Map<string, Relationship> {
  const parsed = parseXml(xml);
  const rels = new Map<string, Relationship>();

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const root = parsed as any;

  if (!root?.Relationships) {
    console.warn(`${WARN_PREFIX} Relationship: missing root element "Relationships" in XML`);
    return rels;
  }

  const relationships = root.Relationships.Relationship;
  if (!relationships) return rels;

  for (const rel of relationships) {
    const id = rel["@_Id"] as string | undefined;
    const type = rel["@_Type"] as string | undefined;
    const target = rel["@_Target"] as string | undefined;

    if (!id || !type || !target) {
      console.warn(`${WARN_PREFIX} Relationship: entry missing required attribute, skipping`);
      continue;
    }

    rels.set(id, { id, type, target });
  }

  return rels;
}

export function buildRelsPath(xmlPath: string): string {
  const lastSlash = xmlPath.lastIndexOf("/");
  const dir = xmlPath.substring(0, lastSlash);
  const filename = xmlPath.substring(lastSlash + 1);
  return `${dir}/_rels/${filename}.rels`;
}

export function resolveRelationshipTarget(basePath: string, relTarget: string): string {
  if (relTarget.startsWith("/")) {
    return relTarget.slice(1);
  }
  const baseDir = basePath.substring(0, basePath.lastIndexOf("/"));
  const parts = `${baseDir}/${relTarget}`.split("/");
  const resolved: string[] = [];
  for (const part of parts) {
    if (part === "..") {
      resolved.pop();
    } else if (part !== ".") {
      resolved.push(part);
    }
  }
  return resolved.join("/");
}
