import { asPartPath, type PartPath } from "./handles.js";
import type { Relationship, RelationshipTargetMode } from "./package-graph.js";

const RELS_SUFFIX = ".rels";
const RELS_MARKER = "_rels/";
const ABSOLUTE_URI_PATTERN = /^[a-zA-Z][a-zA-Z0-9+.-]*:/;

export function isRelationshipPart(path: string): boolean {
  return path.endsWith(RELS_SUFFIX) && path.includes(RELS_MARKER);
}

export function relationshipsPartPath(sourcePartPath: PartPath): string {
  if (sourcePartPath === "") return "_rels/.rels";
  const slash = sourcePartPath.lastIndexOf("/");
  const dir = slash === -1 ? "" : sourcePartPath.slice(0, slash + 1);
  const file = slash === -1 ? sourcePartPath : sourcePartPath.slice(slash + 1);
  return `${dir}_rels/${file}.rels`;
}

export function relationshipsSourcePartPath(relsPath: string): PartPath {
  const idx = relsPath.lastIndexOf(RELS_MARKER);
  const dir = relsPath.slice(0, idx);
  const file = relsPath.slice(idx + RELS_MARKER.length);
  const base = file.endsWith(RELS_SUFFIX) ? file.slice(0, -RELS_SUFFIX.length) : file;
  return asPartPath(dir + base);
}

export function resolveRelationshipTarget(sourcePartPath: string, target: string): string {
  if (ABSOLUTE_URI_PATTERN.test(target)) return target;

  const combined = target.startsWith("/")
    ? target.slice(1)
    : joinPackageRelativeTarget(sourcePartPath, target);

  return normalizePackagePath(combined);
}

export function resolveInternalRelationshipTarget(
  sourcePartPath: PartPath,
  relationship: Relationship,
): PartPath | undefined {
  if (relationship.targetMode === "External") return undefined;
  return asPartPath(resolveRelationshipTarget(sourcePartPath, relationship.target));
}

export function parseRelationshipTargetMode(
  value: string | undefined,
): RelationshipTargetMode | undefined {
  if (value === "Internal" || value === "External") return value;
  return undefined;
}

function joinPackageRelativeTarget(sourcePartPath: string, target: string): string {
  const slash = sourcePartPath.lastIndexOf("/");
  const baseDir = slash === -1 ? "" : sourcePartPath.slice(0, slash);
  return baseDir === "" ? target : `${baseDir}/${target}`;
}

function normalizePackagePath(path: string): string {
  const segments: string[] = [];
  for (const segment of path.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") {
      segments.pop();
      continue;
    }
    segments.push(segment);
  }
  return segments.join("/");
}
