import { getChild, parseXml, type XmlNode } from "../reader/xml.js";
import {
  type CrossPartHandleMapping,
  planCrossPartDrawingMove,
  referencedDrawingRelationshipIds,
} from "./cross-part-move-planning.js";
import { sourceHandlesEqual } from "./edit-descriptors.js";
import type { PartPath, RelationshipId, SourceHandle, SourceNodeId } from "./handles.js";
import type { PptxSourceModel } from "./pptx-source-model.js";
import type { RawOoxmlNode, RawSidecar } from "./raw.js";
import type {
  SourceChart,
  SourceConnector,
  SourceFill,
  SourceImage,
  SourceOutline,
  SourceShape,
  SourceShapeNode,
  SourceTextBody,
} from "./shapes.js";

const RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const RELATIONSHIP_ATTRIBUTE_LOCAL_NAMES = new Set(["id", "embed", "link", "dm", "lo", "qs", "cs"]);
const textDecoder = new TextDecoder();

export interface MoveShapesAcrossSlidesOptions {
  /** Insert immediately before this destination-slide root drawing. Omit to append. */
  readonly beforeShapeHandle?: SourceHandle;
}

export interface MoveShapesAcrossSlidesResult {
  readonly document: PptxSourceModel;
  /** Root and nested text identity transitions. Old handles are invalid after the move. */
  readonly moved: readonly CrossPartHandleMapping[];
  readonly affectedSlidePartPaths: readonly PartPath[];
}

type SupportedCrossSlideDrawing = SourceShape | SourceImage | SourceConnector | SourceChart;

/**
 * Moves consecutive non-placeholder shape, picture, connector, and chart roots between two slides.
 * Authored local OOXML is retained, so effective appearance may change under the destination
 * slide's layout, master, theme, or color-map cascade.
 */
export function moveShapesAcrossSlides(
  source: PptxSourceModel,
  shapeHandles: readonly SourceHandle[],
  destinationSlideHandle: SourceHandle,
  options: MoveShapesAcrossSlidesOptions = {},
): MoveShapesAcrossSlidesResult {
  const destinationIndex = source.slides.findIndex((slide) =>
    sourceHandlesEqual(slide.handle, destinationSlideHandle),
  );
  const destinationSlide = source.slides[destinationIndex];
  if (destinationSlide === undefined) {
    throw new Error("moveShapesAcrossSlides: destination must be a slide root handle");
  }
  const sourcePartPath = shapeHandles[0]?.partPath;
  const sourceIndex = source.slides.findIndex((slide) => slide.partPath === sourcePartPath);
  const sourceSlide = source.slides[sourceIndex];
  if (sourceSlide === undefined) {
    throw new Error("moveShapesAcrossSlides: moved drawings must belong to a slide root");
  }

  const selectedIds = new Set(shapeHandles.map((handle) => String(handle.nodeId ?? "")));
  const movedRoots = sourceSlide.shapes.filter(
    (shape) => shape.nodeId !== undefined && selectedIds.has(String(shape.nodeId)),
  );
  if (movedRoots.length !== shapeHandles.length) {
    throw new Error(
      "moveShapesAcrossSlides: every moved handle must identify a slide-root drawing",
    );
  }
  for (const shape of movedRoots) assertSupportedRoot(shape);
  const supportedMovedRoots = movedRoots.filter(isSupportedRoot);
  if (supportedMovedRoots.length !== movedRoots.length) {
    throw new Error("moveShapesAcrossSlides: moved drawing kind is not supported");
  }
  assertCompatibleSlideRootNamespaces(source, sourceSlide.partPath, destinationSlide.partPath);

  let beforeShapeId: SourceNodeId | undefined;
  let insertionIndex = destinationSlide.shapes.length;
  if (options.beforeShapeHandle !== undefined) {
    const anchorIndex = destinationSlide.shapes.findIndex((shape) =>
      sourceHandlesEqual(shape.handle, options.beforeShapeHandle!),
    );
    const anchor = destinationSlide.shapes[anchorIndex];
    if (anchorIndex < 0 || anchor?.nodeId === undefined) {
      throw new Error(
        "moveShapesAcrossSlides: anchor must be a direct child of the destination slide",
      );
    }
    beforeShapeId = anchor.nodeId;
    insertionIndex = anchorIndex;
  }

  let plan;
  try {
    plan = planCrossPartDrawingMove(source, shapeHandles, destinationSlideHandle, {
      // Reissue identity deterministically even when the destination happens to have the same
      // numeric ids available.
      reservedDestinationNodeIds: movedRoots.flatMap((shape) =>
        shape.nodeId === undefined ? [] : [shape.nodeId],
      ),
      reservedDestinationRelationshipIds:
        source.packageGraph.relationships
          .find((group) => group.sourcePartPath === sourceSlide.partPath)
          ?.relationships.map((relationship) => relationship.id) ?? [],
    });
  } catch (cause) {
    if (cause instanceof Error && cause.message.startsWith("planCrossPartDrawingMove:")) {
      throw new Error(`moveShapesAcrossSlides: ${cause.message}`, { cause });
    }
    throw cause;
  }
  const rootSlots = new Map(
    plan.movedRootNodeIds.map((nodeId, index) => [String(nodeId), insertionIndex + index]),
  );
  const moved = plan.handleMappings.map((mapping) => {
    const orderingSlot =
      mapping.before.nodeId === undefined
        ? undefined
        : rootSlots.get(String(mapping.before.nodeId));
    return {
      before: mapping.before,
      after: {
        ...mapping.after,
        ...(orderingSlot !== undefined
          ? { orderingSlot }
          : mapping.after.orderingSlot !== undefined
            ? { orderingSlot: mapping.after.orderingSlot }
            : {}),
        ...(mapping.before.rawSidecarIds !== undefined
          ? { rawSidecarIds: mapping.before.rawSidecarIds }
          : {}),
      },
    };
  });
  const handleMap = new Map(
    moved.map((mapping) => [handleIdentityKey(mapping.before), mapping.after]),
  );
  const nodeIdMap = new Map(
    plan.nodeIdMappings.map((mapping) => [String(mapping.before), mapping.after]),
  );
  const relationshipIdMap = new Map(
    plan.relationshipRemaps.map((mapping) => [String(mapping.before.id), mapping.after.id]),
  );
  const remappedRoots = supportedMovedRoots.map((shape) =>
    remapSupportedDrawing(shape, handleMap, nodeIdMap, relationshipIdMap),
  );
  const remainingSourceShapes = sourceSlide.shapes.filter(
    (shape) => shape.nodeId === undefined || !selectedIds.has(String(shape.nodeId)),
  );
  const nextDestinationShapes = [
    ...destinationSlide.shapes.slice(0, insertionIndex),
    ...remappedRoots,
    ...destinationSlide.shapes.slice(insertionIndex),
  ];

  const stillReferencedSourceRelationships = new Set(
    referencedDrawingRelationshipIds(remainingSourceShapes, [
      ...(sourceSlide.rawSidecars ?? []),
      ...(sourceSlide.background?.kind === "raw" ? [sourceSlide.background.raw] : []),
      ...(sourceSlide.background?.kind === "fill" && sourceSlide.background.fill.kind === "raw"
        ? [sourceSlide.background.fill.raw]
        : []),
    ]),
  );
  if (
    sourceSlide.background?.kind === "fill" &&
    sourceSlide.background.fill.kind === "image" &&
    sourceSlide.background.fill.blipRelationshipId !== undefined
  ) {
    stillReferencedSourceRelationships.add(sourceSlide.background.fill.blipRelationshipId);
  }
  const movedSourceRelationshipIds = new Set(
    plan.relationshipRemaps.map((mapping) => mapping.before.id),
  );
  const nextRelationships = source.packageGraph.relationships.map((group) => {
    if (group.sourcePartPath === sourceSlide.partPath) {
      return {
        ...group,
        relationships: group.relationships.filter(
          (relationship) =>
            !movedSourceRelationshipIds.has(relationship.id) ||
            stillReferencedSourceRelationships.has(relationship.id),
        ),
      };
    }
    if (group.sourcePartPath === destinationSlide.partPath) {
      const additions = plan.relationshipRemaps
        .filter((mapping) => !mapping.reusedDestinationRelationship)
        .map((mapping) => mapping.after);
      return { ...group, relationships: [...group.relationships, ...additions] };
    }
    return group;
  });

  const nextSlides = source.slides.map((slide, index) => {
    if (index === sourceIndex) return { ...slide, shapes: remainingSourceShapes };
    if (index === destinationIndex) return { ...slide, shapes: nextDestinationShapes };
    return slide;
  });
  const document: PptxSourceModel = {
    ...source,
    packageGraph: { ...source.packageGraph, relationships: nextRelationships },
    slides: nextSlides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "moveShapesAcrossSlides",
        sourcePartPath: sourceSlide.partPath,
        destinationPartPath: destinationSlide.partPath,
        sourceShapeIds: plan.movedRootNodeIds,
        destinationShapeIds: plan.nodeIdMappings
          .filter((mapping) => plan.movedRootNodeIds.includes(mapping.before))
          .map((mapping) => mapping.after),
        nodeIdMappings: plan.nodeIdMappings,
        relationshipIdMappings: plan.relationshipRemaps.map((mapping) => ({
          before: mapping.before.id,
          after: mapping.after.id,
        })),
        ...(beforeShapeId !== undefined ? { beforeShapeId } : {}),
      },
    ],
  };
  return {
    document,
    moved,
    affectedSlidePartPaths: plan.affectedSlidePartPaths,
  };
}

function assertSupportedRoot(shape: SourceShapeNode): asserts shape is SupportedCrossSlideDrawing {
  if (
    shape.kind !== "shape" &&
    shape.kind !== "image" &&
    shape.kind !== "connector" &&
    shape.kind !== "chart"
  ) {
    throw new Error(
      `moveShapesAcrossSlides: drawing kind '${shape.kind}' is outside the slide-to-slide typed slice`,
    );
  }
  if (shape.kind !== "connector" && shape.placeholder !== undefined) {
    throw new Error("moveShapesAcrossSlides: placeholder drawings are not supported");
  }
}

function isSupportedRoot(shape: SourceShapeNode): shape is SupportedCrossSlideDrawing {
  return (
    shape.kind === "shape" ||
    shape.kind === "image" ||
    shape.kind === "connector" ||
    shape.kind === "chart"
  );
}

function remapSupportedDrawing(
  shape: SupportedCrossSlideDrawing,
  handleMap: ReadonlyMap<string, SourceHandle>,
  nodeIdMap: ReadonlyMap<string, SourceNodeId>,
  relationshipIdMap: ReadonlyMap<string, RelationshipId>,
): SupportedCrossSlideDrawing {
  const nodeId = shape.nodeId === undefined ? undefined : nodeIdMap.get(String(shape.nodeId));
  const handle =
    shape.handle === undefined ? undefined : handleMap.get(handleIdentityKey(shape.handle));
  const rawSidecars = remapRawSidecars(shape.rawSidecars, nodeIdMap, relationshipIdMap);
  if (shape.kind === "image") {
    return {
      ...shape,
      ...(nodeId !== undefined ? { nodeId } : {}),
      ...(handle !== undefined ? { handle } : {}),
      ...(shape.blipRelationshipId !== undefined
        ? {
            blipRelationshipId:
              relationshipIdMap.get(String(shape.blipRelationshipId)) ?? shape.blipRelationshipId,
          }
        : {}),
      ...(rawSidecars !== undefined ? { rawSidecars } : {}),
    };
  }
  if (shape.kind === "connector") {
    return {
      ...shape,
      ...(nodeId !== undefined ? { nodeId } : {}),
      ...(handle !== undefined ? { handle } : {}),
      ...(shape.connection !== undefined
        ? {
            connection: {
              ...(shape.connection.start !== undefined
                ? {
                    start: {
                      ...shape.connection.start,
                      shapeId:
                        nodeIdMap.get(String(shape.connection.start.shapeId)) ??
                        shape.connection.start.shapeId,
                    },
                  }
                : {}),
              ...(shape.connection.end !== undefined
                ? {
                    end: {
                      ...shape.connection.end,
                      shapeId:
                        nodeIdMap.get(String(shape.connection.end.shapeId)) ??
                        shape.connection.end.shapeId,
                    },
                  }
                : {}),
            },
          }
        : {}),
      ...(shape.outline !== undefined
        ? { outline: remapOutline(shape.outline, relationshipIdMap) }
        : {}),
      ...(rawSidecars !== undefined ? { rawSidecars } : {}),
    };
  }
  if (shape.kind === "chart") {
    return {
      ...shape,
      ...(nodeId !== undefined ? { nodeId } : {}),
      ...(handle !== undefined ? { handle } : {}),
      ...(shape.chartRelationshipId !== undefined
        ? {
            chartRelationshipId:
              relationshipIdMap.get(String(shape.chartRelationshipId)) ?? shape.chartRelationshipId,
          }
        : {}),
      ...(rawSidecars !== undefined ? { rawSidecars } : {}),
    };
  }
  return {
    ...shape,
    ...(nodeId !== undefined ? { nodeId } : {}),
    ...(handle !== undefined ? { handle } : {}),
    ...(shape.fill !== undefined ? { fill: remapFill(shape.fill, relationshipIdMap) } : {}),
    ...(shape.outline !== undefined
      ? { outline: remapOutline(shape.outline, relationshipIdMap) }
      : {}),
    ...(shape.textBody !== undefined
      ? {
          textBody: remapTextBody(
            shape.textBody,
            handle?.partPath ?? shape.handle?.partPath,
            handleMap,
            nodeIdMap,
            relationshipIdMap,
          ),
        }
      : {}),
    ...(rawSidecars !== undefined ? { rawSidecars } : {}),
  };
}

function remapTextBody(
  textBody: SourceTextBody,
  destinationPartPath: PartPath | undefined,
  handleMap: ReadonlyMap<string, SourceHandle>,
  nodeIdMap: ReadonlyMap<string, SourceNodeId>,
  relationshipIdMap: ReadonlyMap<string, RelationshipId>,
): SourceTextBody {
  return {
    ...textBody,
    ...(textBody.handle !== undefined
      ? {
          handle:
            handleMap.get(handleIdentityKey(textBody.handle)) ??
            (destinationPartPath === undefined
              ? textBody.handle
              : { ...textBody.handle, partPath: destinationPartPath }),
        }
      : {}),
    paragraphs: textBody.paragraphs.map((paragraph) => ({
      ...paragraph,
      ...(paragraph.handle !== undefined
        ? { handle: handleMap.get(handleIdentityKey(paragraph.handle)) ?? paragraph.handle }
        : {}),
      runs: paragraph.runs.map((run) => ({
        ...run,
        ...(run.handle !== undefined
          ? { handle: handleMap.get(handleIdentityKey(run.handle)) ?? run.handle }
          : {}),
        ...(run.rawSidecars !== undefined
          ? { rawSidecars: remapRawSidecars(run.rawSidecars, nodeIdMap, relationshipIdMap) }
          : {}),
      })),
      ...(paragraph.rawSidecars !== undefined
        ? {
            rawSidecars: remapRawSidecars(paragraph.rawSidecars, nodeIdMap, relationshipIdMap),
          }
        : {}),
    })),
    ...(textBody.rawSidecars !== undefined
      ? { rawSidecars: remapRawSidecars(textBody.rawSidecars, nodeIdMap, relationshipIdMap) }
      : {}),
  };
}

function remapOutline(
  outline: SourceOutline,
  relationshipIdMap: ReadonlyMap<string, RelationshipId>,
): SourceOutline {
  return {
    ...outline,
    ...(outline.fill !== undefined ? { fill: remapFill(outline.fill, relationshipIdMap) } : {}),
  };
}

function remapFill(
  fill: SourceFill,
  relationshipIdMap: ReadonlyMap<string, RelationshipId>,
): SourceFill {
  if (fill.kind === "image" && fill.blipRelationshipId !== undefined) {
    return {
      ...fill,
      blipRelationshipId:
        relationshipIdMap.get(String(fill.blipRelationshipId)) ?? fill.blipRelationshipId,
    };
  }
  if (fill.kind === "raw") {
    return {
      ...fill,
      raw: {
        ...fill.raw,
        node: remapRawNode(fill.raw.node, new Map(), relationshipIdMap),
      },
    };
  }
  return fill;
}

function remapRawSidecars(
  sidecars: readonly RawSidecar[] | undefined,
  nodeIdMap: ReadonlyMap<string, SourceNodeId>,
  relationshipIdMap: ReadonlyMap<string, RelationshipId>,
): readonly RawSidecar[] | undefined {
  return sidecars?.map((sidecar) => ({
    ...sidecar,
    node: remapRawNode(sidecar.node, nodeIdMap, relationshipIdMap),
  }));
}

function remapRawNode(
  node: RawOoxmlNode,
  nodeIdMap: ReadonlyMap<string, SourceNodeId>,
  relationshipIdMap: ReadonlyMap<string, RelationshipId>,
  inheritedNamespaces: ReadonlyMap<string, string> = new Map([["r", RELATIONSHIPS_NAMESPACE]]),
): RawOoxmlNode {
  const namespaces = new Map(inheritedNamespaces);
  for (const [name, value] of Object.entries(node.attributes ?? {})) {
    if (name === "xmlns") namespaces.set("", value);
    else if (name.startsWith("xmlns:")) namespaces.set(name.slice(6), value);
  }
  const local = node.name.split(":").at(-1);
  const attributes = Object.fromEntries(
    Object.entries(node.attributes ?? {}).map(([name, value]) => {
      const colon = name.indexOf(":");
      const prefix = colon < 0 ? "" : name.slice(0, colon);
      const attributeLocal = colon < 0 ? name : name.slice(colon + 1);
      const namespace = namespaces.get(prefix);
      const isRelationshipAttribute =
        colon >= 0 &&
        (namespace === RELATIONSHIPS_NAMESPACE || namespace === STRICT_RELATIONSHIPS_NAMESPACE) &&
        RELATIONSHIP_ATTRIBUTE_LOCAL_NAMES.has(attributeLocal);
      if (isRelationshipAttribute) {
        return [name, relationshipIdMap.get(value) ?? value];
      }
      if ((local === "stCxn" || local === "endCxn") && attributeLocal === "id") {
        return [name, nodeIdMap.get(value) ?? value];
      }
      return [name, value];
    }),
  );
  return {
    ...node,
    ...(Object.keys(attributes).length > 0 ? { attributes } : {}),
    ...(node.children !== undefined
      ? {
          children: node.children.map((child) =>
            remapRawNode(child, nodeIdMap, relationshipIdMap, namespaces),
          ),
        }
      : {}),
  };
}

function assertCompatibleSlideRootNamespaces(
  source: PptxSourceModel,
  sourcePartPath: PartPath,
  destinationPartPath: PartPath,
): void {
  const sourceRoot = slideRoot(source, sourcePartPath);
  const destinationRoot = slideRoot(source, destinationPartPath);
  for (const [key, sourceValue] of Object.entries(sourceRoot)) {
    if (key !== "@_xmlns" && !key.startsWith("@_xmlns:")) continue;
    const destinationValue = destinationRoot[key];
    if (destinationValue !== undefined && destinationValue !== sourceValue) {
      const prefix = key === "@_xmlns" ? "(default)" : key.slice("@_xmlns:".length);
      throw new Error(
        `moveShapesAcrossSlides: source and destination bind namespace prefix '${prefix}' differently`,
      );
    }
  }
}

function slideRoot(source: PptxSourceModel, partPath: PartPath): XmlNode {
  const rawPart = source.packageGraph.rawParts?.find((part) => part.partPath === partPath);
  if (rawPart?.kind !== "binary") {
    throw new Error(`moveShapesAcrossSlides: slide '${partPath}' has no raw XML material`);
  }
  const root = getChild(parseXml(textDecoder.decode(rawPart.bytes)), "sld");
  if (root === undefined) {
    throw new Error(`moveShapesAcrossSlides: slide '${partPath}' has no slide root`);
  }
  return root;
}

function handleIdentityKey(handle: SourceHandle): string {
  return `${handle.partPath}\u0000${handle.nodeId ?? ""}`;
}
