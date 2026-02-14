import { parseXml, type XmlNode } from "./xml-parser.js";
import { debug } from "../warning-logger.js";

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

export function parseRelationships(xml: string): Map<string, Relationship> {
  const parsed = parseXml(xml);
  const rels = new Map<string, Relationship>();

  const root = parsed.Relationships as XmlNode | undefined;

  if (!root) {
    debug("relationship.missing", `missing root element "Relationships" in XML`);
    return rels;
  }

  const relationships = root.Relationship as XmlNode[] | undefined;
  if (!relationships) return rels;

  for (const rel of relationships) {
    const id = rel["@_Id"] as string | undefined;
    const type = rel["@_Type"] as string | undefined;
    const target = rel["@_Target"] as string | undefined;
    const targetMode = rel["@_TargetMode"] as string | undefined;

    if (!id || !type || !target) {
      debug("relationship.attribute", "entry missing required attribute, skipping");
      continue;
    }

    rels.set(id, { id, type, target, ...(targetMode && { targetMode }) });
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
