/**
 * Part-local numeric allocation shared by drawing authoring operations.
 *
 * Shape ids are finalized when an edit is created, so allocation must account for the
 * complete typed shape tree and every id still reserved by the pending edit journal.
 * The root `p:spTree` non-visual group properties reserve their own id (commonly 1)
 * even though that node is not represented in `SourceShapeNode[]`. Deleted ids
 * intentionally remain reserved.
 */

import { getAttr, getChild, parseXml } from "../reader/xml.js";
import { editReservedShapeId } from "./edit-descriptors.js";
import { asSourceNodeId, type PartPath, type SourceNodeId } from "./handles.js";
import { nextNumberedName } from "./package-graph-mutations.js";
import type { PptxSourceModel } from "./pptx-source-model.js";
import type { SourceShapeNode } from "./shapes.js";

const decoder = new TextDecoder();

export function nextDrawingShapeId(
  source: PptxSourceModel,
  shapes: readonly SourceShapeNode[],
  partPath: PartPath,
): SourceNodeId {
  const used = new Set<number>();
  const rootShapeId = shapeTreeRootId(source, partPath);
  if (rootShapeId !== undefined) used.add(rootShapeId);
  collectNumericShapeIds(shapes, used);
  for (const edit of source.edits ?? []) {
    const numericId = Number(editReservedShapeId(edit, partPath));
    if (Number.isInteger(numericId) && numericId > 0) used.add(numericId);
  }
  return asSourceNodeId(nextNumberedName(new Set([...used].map(String)), /^(\d+)$/, String));
}

function shapeTreeRootId(source: PptxSourceModel, partPath: PartPath): number | undefined {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === partPath);
  if (rawPart?.kind !== "binary") return undefined;
  const root = parseXml(decoder.decode(rawPart.bytes));
  const partRoot =
    getChild(root, "sld") ?? getChild(root, "sldLayout") ?? getChild(root, "sldMaster");
  const shapeTree = getChild(getChild(partRoot, "cSld"), "spTree");
  const nonVisualProperties = getChild(getChild(shapeTree, "nvGrpSpPr"), "cNvPr");
  const value = Number(getAttr(nonVisualProperties, "id"));
  return Number.isInteger(value) && value >= 0 ? value : undefined;
}

export function nextDrawingOrderingSlot(shapes: readonly SourceShapeNode[]): number {
  return (
    shapes.reduce((current, shape) => Math.max(current, shape.handle?.orderingSlot ?? -1), -1) + 1
  );
}

function collectNumericShapeIds(shapes: readonly SourceShapeNode[], used: Set<number>): void {
  for (const shape of shapes) {
    const numericId = Number(shape.nodeId);
    if (Number.isInteger(numericId) && numericId > 0) used.add(numericId);
    if (shape.kind === "group") collectNumericShapeIds(shape.children, used);
  }
}
