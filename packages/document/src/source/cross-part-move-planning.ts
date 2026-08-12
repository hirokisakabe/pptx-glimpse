import { sourceHandlesEqual } from "./edit-descriptors.js";
import { hasDirtyEditForPart, relativeTarget } from "./editing-shared.js";
import {
  asRelationshipId,
  asSourceNodeId,
  type PartPath,
  type RelationshipId,
  type SourceHandle,
  type SourceNodeId,
} from "./handles.js";
import type { Relationship } from "./package-graph.js";
import { resolveInternalRelationshipTarget } from "./package-paths.js";
import type { PptxSourceModel } from "./pptx-source-model.js";
import type { RawOoxmlNode, RawSidecar } from "./raw.js";
import type {
  SourceCellBorders,
  SourceFill,
  SourceOutline,
  SourceShapeNode,
  SourceTableCell,
  SourceTextBody,
} from "./shapes.js";

const TRANSITIONAL_RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const RELATIONSHIP_ATTRIBUTE_LOCAL_NAMES = new Set(["id", "embed", "link", "dm", "lo", "qs", "cs"]);
const MAX_DRAWING_NODE_ID = 4_294_967_295n;
const STANDARD_RAW_NAMESPACE_CONTEXT = new Map([["r", TRANSITIONAL_RELATIONSHIPS_NAMESPACE]]);

type DrawingPartKind = "slide" | "layout" | "master";

const DRAWING_RELATIONSHIP_CAPABILITIES: Readonly<Record<DrawingPartKind, ReadonlySet<string>>> = {
  slide: drawingRelationshipTypes(),
  layout: drawingRelationshipTypes(),
  master: drawingRelationshipTypes(),
};

interface DrawingPartRoot {
  readonly kind: DrawingPartKind;
  readonly partPath: PartPath;
  readonly handle: SourceHandle;
  readonly shapes: readonly SourceShapeNode[];
}

interface NodeRecord {
  readonly node: SourceShapeNode;
  readonly beforeHandle: SourceHandle;
  readonly oldNodeId: SourceNodeId;
  readonly newNodeId: SourceNodeId;
}

/** One relationship-valued attribute found in preserved OOXML. */
interface RawRelationshipReference {
  readonly elementPath: readonly string[];
  readonly attributeName: string;
  readonly relationshipId: RelationshipId;
}

interface RawRelationshipInventoryOptions {
  readonly inheritedNamespaces?: ReadonlyMap<string, string>;
  readonly rejectUnboundRelationshipAttributes?: boolean;
}

/** One typed or preserved relationship reference owned by a moved drawing node. */
interface CrossPartRelationshipReference {
  readonly ownerHandle: SourceHandle;
  readonly relationshipId: RelationshipId;
  readonly source: "typed" | "raw";
  readonly location: string;
}

/** A relationship to reuse or add in the destination drawing part. */
interface CrossPartRelationshipRemap {
  readonly before: Relationship;
  readonly after: Relationship;
  readonly resolvedTarget: PartPath;
  readonly reusedDestinationRelationship: boolean;
}

/** Stable identity transition for a moved root or descendant. */
interface CrossPartHandleMapping {
  readonly before: SourceHandle;
  readonly after: SourceHandle;
}

/** One connector endpoint that must be rewritten with the node ID map. */
interface CrossPartNodeReferenceRemap {
  readonly ownerHandle: SourceHandle;
  readonly location: "start" | "end";
  readonly before: SourceNodeId;
  readonly after: SourceNodeId;
}

/** Immutable preflight result consumed by a future cross-part writer operation. */
interface CrossPartDrawingMovePlan {
  readonly sourcePartPath: PartPath;
  readonly destinationPartPath: PartPath;
  readonly movedRootNodeIds: readonly SourceNodeId[];
  readonly nodeIdMappings: readonly {
    readonly before: SourceNodeId;
    readonly after: SourceNodeId;
  }[];
  readonly nodeReferenceRemaps: readonly CrossPartNodeReferenceRemap[];
  readonly handleMappings: readonly CrossPartHandleMapping[];
  readonly relationshipReferences: readonly CrossPartRelationshipReference[];
  readonly relationshipRemaps: readonly CrossPartRelationshipRemap[];
  readonly affectedDrawingParts: readonly PartPath[];
  readonly affectedSlidePartPaths: readonly PartPath[];
}

interface PlanCrossPartDrawingMoveOptions {
  /** IDs already claimed by another plan in the same future atomic batch. */
  readonly reservedDestinationNodeIds?: readonly SourceNodeId[];
  /** Relationship IDs already claimed by another plan in the same future atomic batch. */
  readonly reservedDestinationRelationshipIds?: readonly RelationshipId[];
}

/**
 * Builds, but does not apply, a safe root-to-root drawing move plan.
 *
 * This is deliberately package-internal foundation. It does not append an edit or expose the
 * public slide-to-slide operation: the future writer slice will consume the finalized plan.
 */
export function planCrossPartDrawingMove(
  source: PptxSourceModel,
  shapeHandles: readonly SourceHandle[],
  destinationPartHandle: SourceHandle,
  options: PlanCrossPartDrawingMoveOptions = {},
): CrossPartDrawingMovePlan {
  if (shapeHandles.length === 0) {
    throw new Error("planCrossPartDrawingMove: at least one shape handle is required");
  }

  const sourcePartPath = shapeHandles[0]?.partPath;
  if (
    sourcePartPath === undefined ||
    shapeHandles.some((item) => item.partPath !== sourcePartPath)
  ) {
    throw new Error("planCrossPartDrawingMove: moved shapes must belong to one drawing part");
  }
  const sourceRoot = requireDrawingPartRoot(source, sourcePartPath);
  const destinationRoot = drawingPartRoots(source).find((root) =>
    sourceHandlesEqual(root.handle, destinationPartHandle),
  );
  if (destinationRoot === undefined) {
    throw new Error("planCrossPartDrawingMove: destination must be a drawing-part root handle");
  }
  if (sourceRoot.partPath === destinationRoot.partPath) {
    throw new Error("planCrossPartDrawingMove: source and destination parts must differ");
  }
  if (
    hasDirtyEditForPart(source.edits ?? [], sourceRoot.partPath) ||
    hasDirtyEditForPart(source.edits ?? [], destinationRoot.partPath)
  ) {
    throw new Error(
      "planCrossPartDrawingMove: source and destination parts must not have pending edits",
    );
  }

  const movedRoots = requireConsecutiveRootBlock(sourceRoot, shapeHandles);
  for (const root of movedRoots) validateSupportedSubtree(root);

  collectUniqueNodeIds(
    sourceRoot.shapes,
    "planCrossPartDrawingMove: source contains duplicate node ids",
  );

  const movedIds = new Set<string>();
  for (const root of movedRoots) {
    visitShapeTree(root, (node) => {
      if (node.nodeId === undefined) {
        throw new Error(
          "planCrossPartDrawingMove: every moved root and descendant needs a node id",
        );
      }
      if (movedIds.has(node.nodeId)) {
        throw new Error(`planCrossPartDrawingMove: duplicate moved node id '${node.nodeId}'`);
      }
      movedIds.add(node.nodeId);
    });
  }
  validateRawReferenceClosure(sourceRoot.shapes, movedIds);
  validateConnectorClosure(sourceRoot.shapes, movedIds);

  const destinationIds = collectUniqueNodeIds(
    destinationRoot.shapes,
    "planCrossPartDrawingMove: destination contains duplicate node ids",
  );
  for (const reserved of options.reservedDestinationNodeIds ?? []) destinationIds.add(reserved);
  const allocateNodeId = createNumericIdAllocator(
    destinationIds,
    asSourceNodeId,
    "",
    MAX_DRAWING_NODE_ID,
  );
  const nodeRecords: NodeRecord[] = [];
  for (const root of movedRoots) {
    visitShapeTree(root, (node) => {
      const oldNodeId = requireNodeId(node);
      const beforeHandle = requireNodeHandle(node);
      if (beforeHandle.partPath !== sourceRoot.partPath || beforeHandle.nodeId !== oldNodeId) {
        throw new Error(
          "planCrossPartDrawingMove: every moved handle must match its owning part and node id",
        );
      }
      validatePrimaryRelationshipHandle(node, beforeHandle);
      nodeRecords.push({ node, oldNodeId, beforeHandle, newNodeId: allocateNodeId() });
    });
  }

  const relationshipReferences = collectRelationshipReferences(nodeRecords);
  const relationshipRemaps = planRelationshipRemaps(
    source,
    sourceRoot,
    destinationRoot,
    relationshipReferences,
    options.reservedDestinationRelationshipIds ?? [],
  );
  const relationshipIdMap = new Map<string, RelationshipId>(
    relationshipRemaps.map((item) => [item.before.id, item.after.id]),
  );
  const nodeIdMap = new Map<string, SourceNodeId>(
    nodeRecords.map((record) => [record.oldNodeId, record.newNodeId]),
  );
  const nodeReferenceRemaps = nodeRecords.flatMap((record) => {
    if (record.node.kind !== "connector") return [];
    const connector = record.node;
    return (["start", "end"] as const).flatMap((location) => {
      const endpoint = connector.connection?.[location];
      if (endpoint === undefined) return [];
      const after = nodeIdMap.get(endpoint.shapeId);
      if (after === undefined) {
        throw new Error("planCrossPartDrawingMove: connector endpoint mapping was not finalized");
      }
      return [
        freezeObject({
          ownerHandle: freezeHandle(record.beforeHandle),
          location,
          before: endpoint.shapeId,
          after,
        }),
      ];
    });
  });

  const handleMappings = nodeRecords.map((record) => {
    let relationshipId: RelationshipId | undefined;
    if (record.beforeHandle.relationshipId !== undefined) {
      relationshipId = relationshipIdMap.get(record.beforeHandle.relationshipId);
      if (relationshipId === undefined) {
        throw new Error(
          "planCrossPartDrawingMove: a handle relationship id has no finalized remap",
        );
      }
    }
    const after: SourceHandle = {
      partPath: destinationRoot.partPath,
      nodeId: record.newNodeId,
      ...(relationshipId !== undefined ? { relationshipId } : {}),
    };
    return freezeObject({
      before: freezeHandle(record.beforeHandle),
      after: freezeHandle(after),
    });
  });

  const affectedDrawingParts = freezeArray([sourceRoot.partPath, destinationRoot.partPath]);
  const affectedSlidePartPaths = freezeArray(resolveAffectedSlides(source, affectedDrawingParts));

  return freezeObject({
    sourcePartPath: sourceRoot.partPath,
    destinationPartPath: destinationRoot.partPath,
    movedRootNodeIds: freezeArray(movedRoots.map(requireNodeId)),
    nodeIdMappings: freezeArray(
      nodeRecords.map((record) =>
        freezeObject({ before: record.oldNodeId, after: record.newNodeId }),
      ),
    ),
    nodeReferenceRemaps: freezeArray(nodeReferenceRemaps),
    handleMappings: freezeArray(handleMappings),
    relationshipReferences: freezeArray(relationshipReferences.map(freezeRelationshipReference)),
    relationshipRemaps: freezeArray(relationshipRemaps.map(freezeRelationshipRemap)),
    affectedDrawingParts,
    affectedSlidePartPaths,
  });
}

/** Inventories relationship attributes using namespace URI bindings, not prefix spelling. */
export function inventoryRawRelationshipReferences(
  root: RawOoxmlNode,
  options: RawRelationshipInventoryOptions = {},
): readonly RawRelationshipReference[] {
  const references: RawRelationshipReference[] = [];
  visitRawNode(
    root,
    options.inheritedNamespaces ?? new Map(),
    [],
    references,
    options.rejectUnboundRelationshipAttributes ?? false,
  );
  return freezeArray(
    references.map((item) => freezeObject({ ...item, elementPath: freezeArray(item.elementPath) })),
  );
}

function visitRawNode(
  node: RawOoxmlNode,
  inheritedNamespaces: ReadonlyMap<string, string>,
  parentPath: readonly string[],
  references: RawRelationshipReference[],
  rejectUnboundRelationshipAttributes: boolean,
): void {
  const namespaces = new Map(inheritedNamespaces);
  for (const [name, value] of Object.entries(node.attributes ?? {})) {
    if (name === "xmlns") namespaces.set("", value);
    else if (name.startsWith("xmlns:")) namespaces.set(name.slice("xmlns:".length), value);
  }
  const elementPath = [...parentPath, node.name];
  for (const [name, value] of Object.entries(node.attributes ?? {})) {
    const separator = name.indexOf(":");
    if (separator === -1 || name.startsWith("xmlns:")) continue;
    const namespace = namespaces.get(name.slice(0, separator));
    const localAttribute = name.slice(separator + 1);
    if (
      namespace === undefined &&
      rejectUnboundRelationshipAttributes &&
      RELATIONSHIP_ATTRIBUTE_LOCAL_NAMES.has(localAttribute)
    ) {
      throw new Error(
        `planCrossPartDrawingMove: namespace binding for relationship-like attribute '${name}' is unavailable`,
      );
    }
    if (!isRelationshipsNamespace(namespace)) continue;
    references.push({
      elementPath,
      attributeName: name,
      relationshipId: asRelationshipId(value),
    });
  }
  for (const child of node.children ?? []) {
    visitRawNode(child, namespaces, elementPath, references, rejectUnboundRelationshipAttributes);
  }
}

function isRelationshipsNamespace(value: string | undefined): boolean {
  return value === TRANSITIONAL_RELATIONSHIPS_NAMESPACE || value === STRICT_RELATIONSHIPS_NAMESPACE;
}

function requireConsecutiveRootBlock(
  sourceRoot: DrawingPartRoot,
  handles: readonly SourceHandle[],
): readonly SourceShapeNode[] {
  const indexes = handles.map((handle) =>
    sourceRoot.shapes.findIndex(
      (shape) => shape.handle !== undefined && sourceHandlesEqual(shape.handle, handle),
    ),
  );
  if (indexes.some((index) => index < 0)) {
    throw new Error("planCrossPartDrawingMove: every moved shape must be a root direct child");
  }
  if (new Set(indexes).size !== indexes.length) {
    throw new Error("planCrossPartDrawingMove: moved shape handles must be unique");
  }
  const orderedIndexes = [...indexes].sort((left, right) => left - right);
  if (
    orderedIndexes.some(
      (index, position) => position > 0 && index !== orderedIndexes[position - 1] + 1,
    )
  ) {
    throw new Error("planCrossPartDrawingMove: moved root block must be consecutive");
  }
  return orderedIndexes.map((index) => sourceRoot.shapes[index]);
}

function validateSupportedSubtree(node: SourceShapeNode): void {
  if (node.kind === "raw" || node.kind === "smartArt") {
    throw new Error("planCrossPartDrawingMove: raw and SmartArt drawing nodes are not supported");
  }
  if ("placeholder" in node && node.placeholder !== undefined) {
    throw new Error("planCrossPartDrawingMove: placeholder drawings are not supported");
  }
  for (const sidecar of nodeRawSidecars(node)) {
    if (rawContainsElement(sidecar.node, "AlternateContent")) {
      throw new Error("planCrossPartDrawingMove: mc:AlternateContent is not supported");
    }
  }
  validateNodeFills(node);
  if (node.kind === "group") {
    for (const child of node.children) validateSupportedSubtree(child);
  }
}

function validateNodeFills(
  node: Exclude<SourceShapeNode, { readonly kind: "raw" | "smartArt" }>,
): void {
  const fills: (SourceFill | undefined)[] = [];
  if (node.kind === "shape" || node.kind === "group") fills.push(node.fill);
  if (node.kind === "shape" || node.kind === "connector") fills.push(node.outline?.fill);
  if (node.kind === "table") {
    for (const row of node.table.rows) {
      for (const cell of row.cells) {
        fills.push(cell.fill, ...borderFills(cell.borders));
      }
    }
  }
  if (fills.some((fill) => fill?.kind === "raw")) {
    throw new Error("planCrossPartDrawingMove: raw fill content is not supported");
  }
}

function borderFills(borders: SourceCellBorders | undefined): readonly (SourceFill | undefined)[] {
  return [borders?.top?.fill, borders?.bottom?.fill, borders?.left?.fill, borders?.right?.fill];
}

function rawContainsElement(node: RawOoxmlNode, local: string): boolean {
  return (
    node.name.split(":").at(-1) === local ||
    (node.children ?? []).some((child) => rawContainsElement(child, local))
  );
}

function validateConnectorClosure(
  partShapes: readonly SourceShapeNode[],
  movedIds: ReadonlySet<string>,
): void {
  visitShapeTrees(partShapes, (node) => {
    if (node.kind !== "connector") return;
    const connectorMoved = node.nodeId !== undefined && movedIds.has(node.nodeId);
    for (const endpoint of [node.connection?.start, node.connection?.end]) {
      if (endpoint === undefined) continue;
      const endpointMoved = movedIds.has(endpoint.shapeId);
      if (connectorMoved !== endpointMoved) {
        throw new Error(
          "planCrossPartDrawingMove: connector references must stay inside the moved closure",
        );
      }
    }
  });
}

function validateRawReferenceClosure(
  partShapes: readonly SourceShapeNode[],
  movedIds: ReadonlySet<string>,
): void {
  visitShapeTrees(partShapes, (node) => {
    if (node.kind === "raw") {
      throw new Error(
        "planCrossPartDrawingMove: source-part raw drawing nodes make reference closure unprovable",
      );
    }
    for (const sidecar of nodeRawSidecars(node)) {
      if (rawContainsElement(sidecar.node, "AlternateContent")) {
        throw new Error("planCrossPartDrawingMove: mc:AlternateContent is not supported");
      }
      if (rawContainsUnprovenNodeReference(sidecar.node, movedIds)) {
        throw new Error(
          "planCrossPartDrawingMove: a preserved raw sidecar may reference a moved node id",
        );
      }
    }
  });
}

function rawContainsUnprovenNodeReference(
  node: RawOoxmlNode,
  values: ReadonlySet<string>,
): boolean {
  const localElement = node.name.split(":").at(-1);
  for (const [attributeName, value] of Object.entries(node.attributes ?? {})) {
    if (!values.has(value)) continue;
    const localAttribute = attributeName.split(":").at(-1);
    if (localAttribute !== "id" && localAttribute !== "spid") continue;
    const knownNodeIdAttribute =
      localAttribute === "id" &&
      (localElement === "cNvPr" || localElement === "stCxn" || localElement === "endCxn");
    if (!knownNodeIdAttribute) return true;
  }
  return (node.children ?? []).some((child) => rawContainsUnprovenNodeReference(child, values));
}

function collectRelationshipReferences(
  records: readonly NodeRecord[],
): CrossPartRelationshipReference[] {
  const references: CrossPartRelationshipReference[] = [];
  for (const record of records) {
    const typed = typedRelationshipIds(record.node);
    for (const [location, relationshipId] of typed) {
      references.push({
        ownerHandle: record.beforeHandle,
        relationshipId,
        source: "typed",
        location,
      });
    }
    for (const sidecar of nodeRawSidecars(record.node)) {
      for (const raw of inventoryRawRelationshipReferences(sidecar.node, {
        inheritedNamespaces: STANDARD_RAW_NAMESPACE_CONTEXT,
        rejectUnboundRelationshipAttributes: true,
      })) {
        references.push({
          ownerHandle: record.beforeHandle,
          relationshipId: raw.relationshipId,
          source: "raw",
          location: `${sidecar.node.name}/${raw.elementPath.join("/")}@${raw.attributeName}`,
        });
      }
    }
  }
  return references;
}

function typedRelationshipIds(node: SourceShapeNode): readonly [string, RelationshipId][] {
  const result: [string, RelationshipId][] = [];
  if (node.kind === "image" && node.blipRelationshipId !== undefined) {
    result.push(["blipRelationshipId", node.blipRelationshipId]);
  }
  if (node.kind === "chart" && node.chartRelationshipId !== undefined) {
    result.push(["chartRelationshipId", node.chartRelationshipId]);
  }
  collectFillRelationship(
    result,
    "fill",
    node.kind === "shape" || node.kind === "group" ? node.fill : undefined,
  );
  collectOutlineRelationship(
    result,
    "outline",
    node.kind === "shape" || node.kind === "connector" ? node.outline : undefined,
  );
  if (node.kind === "table") {
    node.table.rows.forEach((row, rowIndex) =>
      row.cells.forEach((cell, cellIndex) =>
        collectCellRelationships(result, cell, rowIndex, cellIndex),
      ),
    );
  }
  return result;
}

function collectCellRelationships(
  result: [string, RelationshipId][],
  cell: SourceTableCell,
  rowIndex: number,
  cellIndex: number,
): void {
  const prefix = `table.rows[${rowIndex}].cells[${cellIndex}]`;
  collectFillRelationship(result, `${prefix}.fill`, cell.fill);
  collectOutlineRelationship(result, `${prefix}.borders.top`, cell.borders?.top);
  collectOutlineRelationship(result, `${prefix}.borders.bottom`, cell.borders?.bottom);
  collectOutlineRelationship(result, `${prefix}.borders.left`, cell.borders?.left);
  collectOutlineRelationship(result, `${prefix}.borders.right`, cell.borders?.right);
}

function collectOutlineRelationship(
  result: [string, RelationshipId][],
  location: string,
  outline: SourceOutline | undefined,
): void {
  collectFillRelationship(result, `${location}.fill`, outline?.fill);
}

function collectFillRelationship(
  result: [string, RelationshipId][],
  location: string,
  fill: SourceFill | undefined,
): void {
  if (fill?.kind === "image" && fill.blipRelationshipId !== undefined) {
    result.push([`${location}.blipRelationshipId`, fill.blipRelationshipId]);
  }
}

function nodeRawSidecars(node: SourceShapeNode): readonly RawSidecar[] {
  if (node.kind === "raw") return [node.raw];
  const sidecars = [...(node.rawSidecars ?? [])];
  if (node.kind === "shape") sidecars.push(...textBodyRawSidecars(node.textBody));
  if (node.kind === "table") {
    for (const row of node.table.rows) {
      for (const cell of row.cells) {
        sidecars.push(...(cell.rawSidecars ?? []), ...textBodyRawSidecars(cell.textBody));
        sidecars.push(...rawFillSidecars(cell.fill));
        for (const outline of [
          cell.borders?.top,
          cell.borders?.bottom,
          cell.borders?.left,
          cell.borders?.right,
        ]) {
          sidecars.push(...rawFillSidecars(outline?.fill));
        }
      }
    }
  }
  if (node.kind === "shape" || node.kind === "group") {
    sidecars.push(...rawFillSidecars(node.fill));
  }
  if (node.kind === "shape" || node.kind === "connector") {
    sidecars.push(...rawFillSidecars(node.outline?.fill));
  }
  return sidecars;
}

function textBodyRawSidecars(textBody: SourceTextBody | undefined): readonly RawSidecar[] {
  if (textBody === undefined) return [];
  return [
    ...(textBody.rawSidecars ?? []),
    ...textBody.paragraphs.flatMap((paragraph) => [
      ...(paragraph.rawSidecars ?? []),
      ...paragraph.runs.flatMap((run) => run.rawSidecars ?? []),
    ]),
  ];
}

function rawFillSidecars(fill: SourceFill | undefined): readonly RawSidecar[] {
  return fill?.kind === "raw" ? [fill.raw] : [];
}

function planRelationshipRemaps(
  source: PptxSourceModel,
  sourceRoot: DrawingPartRoot,
  destinationRoot: DrawingPartRoot,
  references: readonly CrossPartRelationshipReference[],
  reservedIds: readonly RelationshipId[],
): CrossPartRelationshipRemap[] {
  const sourceRelationships = requireRelationshipGroup(source, sourceRoot.partPath);
  const destinationRelationships = requireRelationshipGroup(source, destinationRoot.partPath);
  assertUniqueRelationshipIds(sourceRelationships, sourceRoot.partPath);
  assertUniqueRelationshipIds(destinationRelationships, destinationRoot.partPath);
  const usedDestinationIds = new Set<string>([
    ...destinationRelationships.map((item) => item.id),
    ...reservedIds,
  ]);
  const allocateRelationshipId = createNumericIdAllocator(
    usedDestinationIds,
    asRelationshipId,
    "rId",
  );
  const remaps: CrossPartRelationshipRemap[] = [];
  const seen = new Set<string>();

  for (const reference of references) {
    if (seen.has(reference.relationshipId)) continue;
    seen.add(reference.relationshipId);
    const before = sourceRelationships.find((item) => item.id === reference.relationshipId);
    if (before === undefined) {
      throw new Error(
        `planCrossPartDrawingMove: relationship '${reference.relationshipId}' was not found`,
      );
    }
    if (before.targetMode === "External") {
      throw new Error("planCrossPartDrawingMove: external relationships are not supported");
    }
    if (!relationshipAllowed(destinationRoot.kind, before.type)) {
      throw new Error(
        `planCrossPartDrawingMove: relationship type '${before.type}' is not allowed by the destination part`,
      );
    }
    const resolvedTarget = resolveInternalRelationshipTarget(sourceRoot.partPath, before);
    if (resolvedTarget === undefined || !packagePartExists(source, resolvedTarget)) {
      throw new Error(`planCrossPartDrawingMove: internal target for '${before.id}' was not found`);
    }
    const reusable = destinationRelationships.find(
      (candidate) =>
        candidate.type === before.type &&
        candidate.targetMode !== "External" &&
        resolveInternalRelationshipTarget(destinationRoot.partPath, candidate) === resolvedTarget,
    );
    const after =
      reusable ??
      ({
        id: allocateRelationshipId(),
        type: before.type,
        target: relativeTarget(destinationRoot.partPath, resolvedTarget),
        ...(before.targetMode !== undefined ? { targetMode: before.targetMode } : {}),
      } satisfies Relationship);
    remaps.push({
      before: { ...before },
      after: { ...after },
      resolvedTarget,
      reusedDestinationRelationship: reusable !== undefined,
    });
  }
  return remaps;
}

function relationshipAllowed(kind: DrawingPartKind, type: string): boolean {
  return DRAWING_RELATIONSHIP_CAPABILITIES[kind].has(type);
}

function drawingRelationshipTypes(): ReadonlySet<string> {
  const suffixes = [
    "chart",
    "image",
    "hyperlink",
    "audio",
    "video",
    "media",
    "oleObject",
    "package",
  ];
  return new Set(
    [
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "http://purl.oclc.org/ooxml/officeDocument/relationships",
    ].flatMap((base) => suffixes.map((suffix) => `${base}/${suffix}`)),
  );
}

function packagePartExists(source: PptxSourceModel, partPath: PartPath): boolean {
  return (
    source.packageGraph.parts.some((part) => part.partPath === partPath) ||
    source.packageGraph.media.some((part) => part.partPath === partPath) ||
    (source.packageGraph.rawParts ?? []).some((part) => part.partPath === partPath)
  );
}

function resolveAffectedSlides(
  source: PptxSourceModel,
  drawingParts: readonly PartPath[],
): readonly PartPath[] {
  const affected = new Set<string>();
  for (const partPath of drawingParts) {
    const slide = source.slides.find((item) => item.partPath === partPath);
    if (slide !== undefined) {
      affected.add(slide.partPath);
      continue;
    }
    const layout = source.slideLayouts.find((item) => item.partPath === partPath);
    if (layout !== undefined) {
      for (const item of source.slides)
        if (item.layoutPartPath === layout.partPath) affected.add(item.partPath);
      continue;
    }
    const master = source.slideMasters.find((item) => item.partPath === partPath);
    if (master !== undefined) {
      const layouts = new Set(
        source.slideLayouts
          .filter((item) => item.masterPartPath === master.partPath)
          .map((item) => item.partPath),
      );
      for (const item of source.slides)
        if (layouts.has(item.layoutPartPath)) affected.add(item.partPath);
    }
  }
  return source.presentation.slidePartPaths.filter((partPath) => affected.has(partPath));
}

function drawingPartRoots(source: PptxSourceModel): readonly DrawingPartRoot[] {
  return [
    ...source.slides.flatMap((item) =>
      item.handle === undefined
        ? []
        : [
            {
              kind: "slide" as const,
              partPath: item.partPath,
              handle: item.handle,
              shapes: item.shapes,
            },
          ],
    ),
    ...source.slideLayouts.flatMap((item) =>
      item.handle === undefined
        ? []
        : [
            {
              kind: "layout" as const,
              partPath: item.partPath,
              handle: item.handle,
              shapes: item.shapes,
            },
          ],
    ),
    ...source.slideMasters.flatMap((item) =>
      item.handle === undefined
        ? []
        : [
            {
              kind: "master" as const,
              partPath: item.partPath,
              handle: item.handle,
              shapes: item.shapes,
            },
          ],
    ),
  ];
}

function requireDrawingPartRoot(source: PptxSourceModel, partPath: PartPath): DrawingPartRoot {
  const root = drawingPartRoots(source).find((item) => item.partPath === partPath);
  if (root === undefined)
    throw new Error(`planCrossPartDrawingMove: drawing part '${partPath}' was not found`);
  return root;
}

function requireRelationshipGroup(
  source: PptxSourceModel,
  partPath: PartPath,
): readonly Relationship[] {
  return (
    source.packageGraph.relationships.find((item) => item.sourcePartPath === partPath)
      ?.relationships ?? []
  );
}

function assertUniqueRelationshipIds(
  relationships: readonly Relationship[],
  partPath: PartPath,
): void {
  const ids = new Set<string>();
  for (const relationship of relationships) {
    if (ids.has(relationship.id))
      throw new Error(
        `planCrossPartDrawingMove: duplicate relationship id '${relationship.id}' in '${partPath}'`,
      );
    ids.add(relationship.id);
  }
}

function collectUniqueNodeIds(shapes: readonly SourceShapeNode[], message: string): Set<string> {
  const ids = new Set<string>();
  visitShapeTrees(shapes, (node) => {
    if (node.nodeId === undefined) return;
    if (ids.has(node.nodeId)) throw new Error(`${message}: '${node.nodeId}'`);
    ids.add(node.nodeId);
  });
  return ids;
}

function createNumericIdAllocator<T extends string>(
  used: Set<string>,
  brand: (value: string) => T,
  prefix = "",
  maximum?: bigint,
): () => T {
  let largest = 0n;
  for (const value of used) {
    const numeric =
      prefix === "" ? value : value.startsWith(prefix) ? value.slice(prefix.length) : "";
    if (!/^\d+$/.test(numeric)) continue;
    const parsed = BigInt(numeric);
    if (maximum !== undefined && parsed > maximum) {
      throw new Error(
        `planCrossPartDrawingMove: numeric id '${value}' exceeds the supported OOXML range`,
      );
    }
    if (parsed > largest) largest = parsed;
  }
  let next = largest + 1n;
  return () => {
    if (maximum !== undefined && next > maximum) {
      throw new Error("planCrossPartDrawingMove: no drawing node id remains in the OOXML range");
    }
    while (used.has(`${prefix}${next}`)) next += 1n;
    const value = `${prefix}${next}`;
    used.add(value);
    next += 1n;
    return brand(value);
  };
}

function visitShapeTrees(
  shapes: readonly SourceShapeNode[],
  visit: (node: SourceShapeNode) => void,
): void {
  for (const shape of shapes) visitShapeTree(shape, visit);
}

function visitShapeTree(node: SourceShapeNode, visit: (node: SourceShapeNode) => void): void {
  visit(node);
  if (node.kind === "group") for (const child of node.children) visitShapeTree(child, visit);
}

function requireNodeId(node: SourceShapeNode): SourceNodeId {
  if (node.nodeId === undefined) throw new Error("planCrossPartDrawingMove: node id is required");
  return node.nodeId;
}

function requireNodeHandle(node: SourceShapeNode): SourceHandle {
  if (node.handle === undefined)
    throw new Error("planCrossPartDrawingMove: source handle is required for every moved node");
  return node.handle;
}

function validatePrimaryRelationshipHandle(node: SourceShapeNode, handle: SourceHandle): void {
  const typedRelationshipId =
    node.kind === "image"
      ? node.blipRelationshipId
      : node.kind === "chart"
        ? node.chartRelationshipId
        : undefined;
  if (
    typedRelationshipId !== undefined &&
    handle.relationshipId !== undefined &&
    typedRelationshipId !== handle.relationshipId
  ) {
    throw new Error(
      "planCrossPartDrawingMove: typed relationship id and source handle relationship id differ",
    );
  }
}

function freezeHandle(handle: SourceHandle): SourceHandle {
  return freezeObject({
    ...handle,
    ...(handle.rawSidecarIds !== undefined
      ? { rawSidecarIds: freezeArray(handle.rawSidecarIds) }
      : {}),
  });
}

function freezeRelationshipReference(
  item: CrossPartRelationshipReference,
): CrossPartRelationshipReference {
  return freezeObject({ ...item, ownerHandle: freezeHandle(item.ownerHandle) });
}

function freezeRelationshipRemap(item: CrossPartRelationshipRemap): CrossPartRelationshipRemap {
  return freezeObject({
    ...item,
    before: freezeObject({ ...item.before }),
    after: freezeObject({ ...item.after }),
  });
}

function freezeArray<T>(items: readonly T[]): readonly T[] {
  return Object.freeze([...items]);
}

function freezeObject<T extends object>(value: T): T {
  return Object.freeze(value);
}
