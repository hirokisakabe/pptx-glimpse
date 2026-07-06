import { getAttr, getChild, getChildArray, parseXml } from "../reader/xml.js";
import { editDirtyPartPath, editInvalidatingPartPaths } from "./edit-descriptors.js";
import type {
  PartPath,
  PartRelationships,
  PptxSourceModel,
  PptxSourceModelEdit,
  RawPackagePart,
  Relationship,
  RelationshipId,
  SourceShapeNode,
} from "./index.js";
import { resolveInternalRelationshipTarget } from "./package-paths.js";

export const textDecoder = new TextDecoder();

export const SLIDE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
export const NOTES_SLIDE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
export const SLIDE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
export const SLIDE_LAYOUT_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
export const NOTES_SLIDE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
export const IMAGE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
export const EMPTY_SLIDE_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
  `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
  `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
  `<p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
  `<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
  `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

export function requirePartRelationships(
  source: PptxSourceModel,
  partPath: PartPath,
  operationName: string,
): PartRelationships {
  const relationships = source.packageGraph.relationships.find(
    (candidate) => candidate.sourcePartPath === partPath,
  );
  if (relationships === undefined) {
    throw new Error(`${operationName}: presentation relationships were not found`);
  }
  return relationships;
}

export function requireSlideRelationship(
  source: PptxSourceModel,
  relationships: PartRelationships,
  slidePartPath: PartPath,
  operationName: string,
): Relationship {
  const relationship = relationships.relationships.find(
    (candidate) =>
      candidate.type === SLIDE_REL_TYPE &&
      candidate.targetMode !== "External" &&
      resolveInternalRelationshipTarget(source.presentation.partPath, candidate) === slidePartPath,
  );
  if (relationship === undefined) {
    throw new Error(`${operationName}: slide relationship was not found in presentation.xml.rels`);
  }
  return relationship;
}

export function requireRawBinaryPart(
  source: PptxSourceModel,
  partPath: PartPath,
  operationName: string,
): Extract<RawPackagePart, { readonly kind: "binary" }> {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === partPath);
  if (rawPart === undefined) {
    throw new Error(`${operationName}: part '${partPath}' has no preserved raw package material`);
  }
  if (rawPart.kind !== "binary") {
    throw new Error(
      `${operationName}: part '${partPath}' is not backed by binary package material`,
    );
  }
  return rawPart;
}

/**
 * Assigns the numeric `p:sldId@id` for a new slide at edit time by reading the
 * preserved presentation XML and the ids already claimed by pending slide edits.
 * Ids freed by pending deletes are intentionally never reused.
 */
export function nextSlideNumericId(source: PptxSourceModel, operationName: string): number {
  const rawPart = requireRawBinaryPart(source, source.presentation.partPath, operationName);
  const root = parseXml(textDecoder.decode(rawPart.bytes));
  const presentation = getChild(root, "presentation");
  if (presentation === undefined) {
    throw new Error(`${operationName}: presentation part does not contain p:presentation root`);
  }
  const used = new Set<number>();
  for (const item of getChildArray(getChild(presentation, "sldIdLst"), "sldId")) {
    const id = Number(getAttr(item, "id"));
    if (Number.isFinite(id)) used.add(id);
  }
  for (const edit of source.edits ?? []) {
    if (edit.kind === "addEmptySlideFromLayout" || edit.kind === "duplicateSlide") {
      used.add(edit.newSlideNumericId);
    }
  }
  const max = Math.max(255, ...used);
  for (let candidate = max + 1; ; candidate += 1) {
    if (!used.has(candidate)) return candidate;
  }
}

export function presentationSlideRelationship(
  source: PptxSourceModel,
  relationshipId: RelationshipId,
  slidePartPath: PartPath,
): Relationship {
  return {
    id: relationshipId,
    type: SLIDE_REL_TYPE,
    target: relativeTarget(source.presentation.partPath, slidePartPath),
  };
}

export function relativeTarget(sourcePartPath: PartPath, targetPartPath: PartPath): string {
  const sourceDir = sourcePartPath.split("/").slice(0, -1);
  const targetSegments = targetPartPath.split("/");
  while (sourceDir.length > 0 && targetSegments.length > 0 && sourceDir[0] === targetSegments[0]) {
    sourceDir.shift();
    targetSegments.shift();
  }
  return [...sourceDir.map(() => ".."), ...targetSegments].join("/");
}

export function cloneJson<T>(value: T): T {
  return structuredClone(value);
}

export function copyBytes(bytes: Uint8Array): Uint8Array {
  return new Uint8Array(bytes);
}

export function insertAtReadonly<T>(items: readonly T[], index: number, item: T): readonly T[] {
  return [...items.slice(0, index), item, ...items.slice(index)];
}

export function hasDirtyEditForPart(
  edits: readonly PptxSourceModelEdit[],
  partPath: PartPath,
): boolean {
  return edits.some((edit) => editDirtyPartPath(edit) === partPath);
}

export function editIsInvalidatedByDeletedParts(
  edit: PptxSourceModelEdit,
  partPaths: ReadonlySet<string>,
): boolean {
  return editInvalidatingPartPaths(edit).some((partPath) => partPaths.has(partPath));
}

export function assertNeverShapeNode(shape: never): never {
  throw new Error(`editing: unhandled source shape node kind '${String(shape)}'`);
}
