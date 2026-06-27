import { debug } from "@pptx-glimpse/renderer";

import { unsafeTypeAssertion } from "../unsafe-type-assertion.js";
import { parseXml, type XmlNode } from "./xml-parser.js";

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

export function parseRelationships(xml: string): Map<string, Relationship> {
  const parsed = parseXml(xml);
  const rels = new Map<string, Relationship>();

  const root = unsafeTypeAssertion<XmlNode | undefined>(parsed.Relationships);

  if (!root) {
    debug("relationship.missing", `missing root element "Relationships" in XML`);
    return rels;
  }

  const relationships = unsafeTypeAssertion<XmlNode[] | undefined>(root.Relationship);
  if (!relationships) return rels;

  for (const rel of relationships) {
    const id = unsafeTypeAssertion<string | undefined>(rel["@_Id"]);
    const type = unsafeTypeAssertion<string | undefined>(rel["@_Type"]);
    const target = unsafeTypeAssertion<string | undefined>(rel["@_Target"]);
    const targetMode = unsafeTypeAssertion<string | undefined>(rel["@_TargetMode"]);

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
