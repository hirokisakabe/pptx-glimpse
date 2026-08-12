import { describe, expect, it } from "vitest";

import { createPptx } from "../builder/create-pptx.js";
import { readPptx } from "../reader/index.js";
import { writePptx } from "../writer/index.js";
import { BLUE_PNG } from "../writer/write-pptx.test-helpers.js";
import { createPptxAuthoringSession } from "./authoring-session.js";
import {
  inventoryRawRelationshipReferences,
  planCrossPartDrawingMove,
} from "./cross-part-move-planning.js";
import { asRawSidecarId, asRelationshipId, asSourceNodeId, type SourceHandle } from "./handles.js";
import { resolveInternalRelationshipTarget } from "./package-paths.js";
import type { PptxSourceModel } from "./pptx-source-model.js";
import type { SourceShapeNode } from "./shapes.js";
import { asEmu } from "./units.js";

describe("cross-part move planning", () => {
  it("plans collision-free picture/media and chart/workbook relationship sharing", () => {
    const source = buildCleanTwoSlideSource((session, first, second) => {
      const sourceTarget = session.target(first);
      sourceTarget.addPicture({
        bytes: BLUE_PNG,
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
      });
      sourceTarget.addChart({
        chartType: "bar",
        offsetX: asEmu(1200),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        series: [{ categories: ["A"], values: [1] }],
      });
      session.target(second).addShape(rectInput(0));
    });
    const sourceSlide = requireValue(source.slides[0]);
    const destinationSlide = requireValue(source.slides[1]);
    const movedHandles = sourceSlide.shapes.map((node) => requireValue(node.handle));
    const picture = requireValue(sourceSlide.shapes.find((node) => node.kind === "image"));
    const pictureRelationshipId = requireValue(picture.handle?.relationshipId);
    const sourceWithRawReference = replaceFirstShape(source, {
      ...picture,
      rawSidecars: [
        ...(picture.rawSidecars ?? []),
        {
          id: asRawSidecarId("raw-relationship"),
          node: {
            name: "p:ext",
            attributes: {
              "xmlns:link": "http://purl.oclc.org/ooxml/officeDocument/relationships",
              "link:id": pictureRelationshipId,
            },
          },
        },
      ],
    });

    const plan = planCrossPartDrawingMove(
      sourceWithRawReference,
      [...movedHandles].reverse(),
      requireValue(destinationSlide.handle),
    );

    expect(plan.movedRootNodeIds).toEqual(sourceSlide.shapes.map((node) => node.nodeId));
    const destinationIds = new Set(destinationSlide.shapes.map((node) => node.nodeId));
    expect(plan.nodeIdMappings.map((item) => Number(item.after))).toEqual([1, 3]);
    expect(plan.nodeIdMappings.every((item) => !destinationIds.has(item.after))).toBe(true);
    expect(plan.relationshipRemaps).toHaveLength(2);
    expect(plan.relationshipReferences.filter((item) => item.source === "typed")).toHaveLength(2);
    expect(plan.relationshipReferences.filter((item) => item.source === "raw")).toHaveLength(2);
    expect(new Set(plan.relationshipReferences.map((item) => item.relationshipId))).toEqual(
      new Set(plan.relationshipRemaps.map((item) => item.before.id)),
    );
    expect(plan.relationshipRemaps.map((item) => item.before.type.split("/").at(-1))).toEqual([
      "image",
      "chart",
    ]);
    expect(
      plan.relationshipRemaps.every((item) => item.before.targetMode === item.after.targetMode),
    ).toBe(true);
    expect(
      plan.relationshipRemaps.every((item) => packagePartExists(source, item.resolvedTarget)),
    ).toBe(true);
    expect(plan.relationshipRemaps.every((item) => !item.reusedDestinationRelationship)).toBe(true);
    expect(new Set(plan.relationshipRemaps.map((item) => item.after.id)).size).toBe(2);
    expect(
      plan.relationshipRemaps.map((item) =>
        resolveInternalRelationshipTarget(destinationSlide.partPath, item.after),
      ),
    ).toEqual(plan.relationshipRemaps.map((item) => item.resolvedTarget));

    const chartPath = requireValue(
      plan.relationshipRemaps.find((item) => item.before.type.endsWith("/chart")),
    ).resolvedTarget;
    const workbookRelationship = source.packageGraph.relationships
      .find((group) => group.sourcePartPath === chartPath)
      ?.relationships.find((relationship) => relationship.type.endsWith("/package"));
    expect(workbookRelationship).toBeDefined();
    expect(
      resolveInternalRelationshipTarget(chartPath, requireValue(workbookRelationship)),
    ).toSatisfy((partPath: string | undefined) =>
      partPath === undefined ? false : packagePartExists(source, partPath),
    );
    expect(plan.handleMappings.map((item) => item.after.partPath)).toEqual([
      destinationSlide.partPath,
      destinationSlide.partPath,
    ]);
    expect(plan.affectedDrawingParts).toEqual([sourceSlide.partPath, destinationSlide.partPath]);
    expect(plan.affectedSlidePartPaths).toEqual([sourceSlide.partPath, destinationSlide.partPath]);
    expect(Object.isFrozen(plan)).toBe(true);
    expect(Object.isFrozen(plan.relationshipRemaps)).toBe(true);

    const rawSidecarIds = [asRawSidecarId("mutable-sidecar")];
    const mutableHandleSource = replaceFirstShape(source, {
      ...picture,
      handle: { ...requireValue(picture.handle), rawSidecarIds },
    });
    const immutablePlan = planCrossPartDrawingMove(
      mutableHandleSource,
      [{ ...requireValue(picture.handle), rawSidecarIds }],
      requireValue(destinationSlide.handle),
    );
    rawSidecarIds.push(asRawSidecarId("later-sidecar"));
    expect(immutablePlan.handleMappings[0]?.before.rawSidecarIds).toEqual(["mutable-sidecar"]);
    expect(Object.isFrozen(immutablePlan.handleMappings[0]?.before.rawSidecarIds)).toBe(true);

    const reservedPlan = planCrossPartDrawingMove(
      source,
      movedHandles,
      requireValue(destinationSlide.handle),
      {
        reservedDestinationNodeIds: [asSourceNodeId("1")],
        reservedDestinationRelationshipIds: [asRelationshipId("rId100")],
      },
    );
    expect(reservedPlan.nodeIdMappings.map((item) => Number(item.after))).toEqual([3, 4]);
    expect(reservedPlan.relationshipRemaps.map((item) => item.after.id)).toEqual([
      "rId101",
      "rId102",
    ]);

    const pictureRelationship = requireValue(
      source.packageGraph.relationships
        .find((group) => group.sourcePartPath === sourceSlide.partPath)
        ?.relationships.find((relationship) => relationship.id === pictureRelationshipId),
    );
    const withReusableDestinationRelationship: PptxSourceModel = {
      ...source,
      packageGraph: {
        ...source.packageGraph,
        relationships: source.packageGraph.relationships.map((group) =>
          group.sourcePartPath !== destinationSlide.partPath
            ? group
            : {
                ...group,
                relationships: [
                  ...group.relationships,
                  { ...pictureRelationship, id: asRelationshipId("rId77") },
                ],
              },
        ),
      },
    };
    const reusedPlan = planCrossPartDrawingMove(
      withReusableDestinationRelationship,
      [requireValue(picture.handle)],
      requireValue(destinationSlide.handle),
    );
    expect(reusedPlan.relationshipRemaps).toHaveLength(1);
    expect(reusedPlan.relationshipRemaps[0]?.after.id).toBe("rId77");
    expect(reusedPlan.relationshipRemaps[0]?.reusedDestinationRelationship).toBe(true);

    const explicitInternalReuse: PptxSourceModel = {
      ...withReusableDestinationRelationship,
      packageGraph: {
        ...withReusableDestinationRelationship.packageGraph,
        relationships: withReusableDestinationRelationship.packageGraph.relationships.map(
          (group) => ({
            ...group,
            relationships: group.relationships.map((relationship) =>
              relationship.id === "rId77"
                ? { ...relationship, targetMode: "Internal" }
                : relationship,
            ),
          }),
        ),
      },
    };
    expect(
      planCrossPartDrawingMove(
        explicitInternalReuse,
        [requireValue(picture.handle)],
        requireValue(destinationSlide.handle),
      ).relationshipRemaps[0]?.reusedDestinationRelationship,
    ).toBe(true);

    const external: PptxSourceModel = {
      ...sourceWithRawReference,
      packageGraph: {
        ...sourceWithRawReference.packageGraph,
        relationships: sourceWithRawReference.packageGraph.relationships.map((group) =>
          group.sourcePartPath !== sourceSlide.partPath
            ? group
            : {
                ...group,
                relationships: group.relationships.map((relationship) =>
                  relationship.id !== picture.handle?.relationshipId
                    ? relationship
                    : {
                        ...relationship,
                        target: "https://example.com/image.png",
                        targetMode: "External",
                      },
                ),
              },
        ),
      },
    };
    expect(() =>
      planCrossPartDrawingMove(external, movedHandles, requireValue(destinationSlide.handle)),
    ).toThrow("external relationships");

    const unsupportedType: PptxSourceModel = {
      ...sourceWithRawReference,
      packageGraph: {
        ...sourceWithRawReference.packageGraph,
        relationships: sourceWithRawReference.packageGraph.relationships.map((group) =>
          group.sourcePartPath !== sourceSlide.partPath
            ? group
            : {
                ...group,
                relationships: group.relationships.map((relationship) =>
                  relationship.id !== pictureRelationshipId
                    ? relationship
                    : { ...relationship, type: "urn:vendor/relationships/image" },
                ),
              },
        ),
      },
    };
    expect(() =>
      planCrossPartDrawingMove(
        unsupportedType,
        movedHandles,
        requireValue(destinationSlide.handle),
      ),
    ).toThrow("not allowed");
  });

  it("builds preorder handle and internal connector reference mappings for a group closure", () => {
    const source = buildCleanTwoSlideSource((session, first) => {
      const target = session.target(first);
      const start = target.addShape(rectInput(0));
      const end = target.addShape(rectInput(2000));
      const connector = target.addConnector({
        preset: "straightConnector1",
        offsetX: asEmu(1000),
        offsetY: asEmu(500),
        width: asEmu(1000),
        height: asEmu(1),
        start: { shapeHandle: start, connectionSiteIndex: 1 },
        end: { shapeHandle: end, connectionSiteIndex: 3 },
      });
      target.groupShapes([start, end, connector]);
    });
    const sourceSlide = requireValue(source.slides[0]);
    const destinationSlide = requireValue(source.slides[1]);
    const group = requireValue(sourceSlide.shapes[0]);
    expect(group.kind).toBe("group");

    const plan = planCrossPartDrawingMove(
      source,
      [requireValue(group.handle)],
      requireValue(destinationSlide.handle),
    );

    expect(plan.nodeIdMappings).toHaveLength(4);
    expect(plan.handleMappings).toHaveLength(4);
    expect(plan.nodeReferenceRemaps).toHaveLength(2);
    const mapping = new Map(plan.nodeIdMappings.map((item) => [item.before, item.after]));
    expect(plan.nodeReferenceRemaps.map((item) => item.location)).toEqual(["start", "end"]);
    expect(plan.nodeReferenceRemaps.every((item) => item.rawSidecarIds.length === 1)).toBe(true);
    expect(plan.nodeReferenceRemaps.every((item) => item.after === mapping.get(item.before))).toBe(
      true,
    );
    expect(plan.relationshipReferences).toEqual([]);
  });

  it("inventories text-run and table-cell hyperlink sidecars in the relationship closure", () => {
    const source = buildCleanTwoSlideSource((session, first) => {
      session.target(first).addShape({
        ...rectInput(0),
        paragraphs: [{ runs: [{ text: "shape" }] }],
      });
      session.target(first).addTable({
        offsetX: asEmu(1200),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        columnWidths: [asEmu(1000)],
        rows: [{ height: asEmu(1000), cells: [{ text: "cell" }] }],
      });
    });
    const sourceSlide = requireValue(source.slides[0]);
    const destination = requireValue(source.slides[1]?.handle);
    const textHandlePlan = planCrossPartDrawingMove(
      source,
      sourceSlide.shapes.map((node) => requireValue(node.handle)),
      destination,
    );
    expect(textHandlePlan.handleMappings).toHaveLength(6);
    expect(
      textHandlePlan.handleMappings
        .filter((item) => String(item.before.nodeId).startsWith("text:"))
        .map((item) => String(item.after.nodeId)),
    ).toEqual([
      `text:shape:${textHandlePlan.nodeIdMappings[0]?.after}:p:0`,
      `text:shape:${textHandlePlan.nodeIdMappings[0]?.after}:p:0:r:0`,
      `text:table:${textHandlePlan.nodeIdMappings[1]?.after}:row:0:cell:0:p:0`,
      `text:table:${textHandlePlan.nodeIdMappings[1]?.after}:row:0:cell:0:p:0:r:0`,
    ]);
    const hyperlinkId = asRelationshipId("rId99");
    const hyperlink = {
      id: hyperlinkId,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      target: "https://example.com/",
      targetMode: "External" as const,
    };
    const hyperlinkSidecar = {
      id: asRawSidecarId("hyperlink"),
      node: {
        name: "a:hlinkClick",
        attributes: {
          "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "r:id": hyperlinkId,
        },
      },
    };
    const [shape, table] = sourceSlide.shapes;
    if (shape?.kind !== "shape" || table?.kind !== "table") {
      throw new Error("text relationship fixtures are missing");
    }
    const shapeRun = requireValue(shape.textBody?.paragraphs[0]?.runs[0]);
    const tableRun = requireValue(table.table.rows[0]?.cells[0]?.textBody?.paragraphs[0]?.runs[0]);
    const withSidecars: PptxSourceModel = {
      ...source,
      slides: [
        {
          ...sourceSlide,
          shapes: [
            {
              ...shape,
              textBody: {
                ...requireValue(shape.textBody),
                paragraphs: [
                  {
                    ...requireValue(shape.textBody?.paragraphs[0]),
                    runs: [{ ...shapeRun, rawSidecars: [hyperlinkSidecar] }],
                  },
                ],
              },
            },
            {
              ...table,
              table: {
                ...table.table,
                rows: [
                  {
                    ...requireValue(table.table.rows[0]),
                    cells: [
                      {
                        ...requireValue(table.table.rows[0]?.cells[0]),
                        textBody: {
                          ...requireValue(table.table.rows[0]?.cells[0]?.textBody),
                          paragraphs: [
                            {
                              ...requireValue(
                                table.table.rows[0]?.cells[0]?.textBody?.paragraphs[0],
                              ),
                              runs: [{ ...tableRun, rawSidecars: [hyperlinkSidecar] }],
                            },
                          ],
                        },
                      },
                    ],
                  },
                ],
              },
            },
          ],
        },
        ...source.slides.slice(1),
      ],
      packageGraph: {
        ...source.packageGraph,
        relationships: source.packageGraph.relationships.map((group) =>
          group.sourcePartPath === sourceSlide.partPath
            ? { ...group, relationships: [...group.relationships, hyperlink] }
            : group,
        ),
      },
    };

    expect(() =>
      planCrossPartDrawingMove(
        withSidecars,
        [requireValue(shape.handle), requireValue(table.handle)],
        destination,
      ),
    ).toThrow("external relationships");
  });

  it("uses declared namespace bindings when inventorying preserved relationship attributes", () => {
    const references = inventoryRawRelationshipReferences({
      name: "p:ext",
      attributes: {
        "xmlns:link": "http://purl.oclc.org/ooxml/officeDocument/relationships",
        "link:id": "rId7",
        "other:id": "not-a-relationship",
      },
      children: [{ name: "a:blip", attributes: { "link:embed": "rId8" } }],
    });

    expect(references).toEqual([
      { elementPath: ["p:ext"], attributeName: "link:id", relationshipId: "rId7" },
      {
        elementPath: ["p:ext", "a:blip"],
        attributeName: "link:embed",
        relationshipId: "rId8",
      },
    ]);

    const inherited = inventoryRawRelationshipReferences(
      { name: "a:hlinkClick", attributes: { "rel:id": "rId9" } },
      {
        inheritedNamespaces: new Map([
          ["rel", "http://purl.oclc.org/ooxml/officeDocument/relationships"],
        ]),
        rejectUnboundRelationshipAttributes: true,
      },
    );
    expect(inherited[0]?.relationshipId).toBe("rId9");
    expect(() =>
      inventoryRawRelationshipReferences(
        { name: "a:hlinkClick", attributes: { "rel:id": "rId9" } },
        { rejectUnboundRelationshipAttributes: true },
      ),
    ).toThrow("namespace binding");
  });

  it("resolves template fan-out in presentation slide order", () => {
    const source = buildCleanTwoSlideSource((session, _first, _second, layout) => {
      session.target(layout).addShape(rectInput(0));
    });
    const layout = requireValue(source.slideLayouts[0]);
    const master = requireValue(source.slideMasters[0]);
    const shape = requireValue(layout.shapes[0]);

    const plan = planCrossPartDrawingMove(
      source,
      [requireValue(shape.handle)],
      requireValue(master.handle),
    );

    expect(plan.affectedSlidePartPaths).toEqual(source.presentation.slidePartPaths);
  });

  it("atomically rejects incomplete or unsupported closures and pending drawing-part edits", () => {
    const source = buildCleanTwoSlideSource((session, first) => {
      const target = session.target(first);
      const start = target.addShape(rectInput(0));
      const end = target.addShape(rectInput(2000));
      target.addConnector({
        preset: "straightConnector1",
        offsetX: asEmu(1000),
        offsetY: asEmu(500),
        width: asEmu(1000),
        height: asEmu(1),
        start: { shapeHandle: start, connectionSiteIndex: 1 },
        end: { shapeHandle: end, connectionSiteIndex: 3 },
      });
    });
    const before = structuredClone(source);
    const sourceSlide = requireValue(source.slides[0]);
    const destination = requireValue(source.slides[1]?.handle);
    const handles = sourceSlide.shapes.map((node) => requireValue(node.handle));

    expect(() => planCrossPartDrawingMove(source, [handles[0]], destination)).toThrow(
      "connector references",
    );
    expect(() => planCrossPartDrawingMove(source, [handles[0], handles[2]], destination)).toThrow(
      "consecutive",
    );

    const placeholder = replaceFirstShape(source, {
      ...requireValue(sourceSlide.shapes[0]),
      placeholder: { type: "body", index: 1 },
    });
    expect(() => planCrossPartDrawingMove(placeholder, [handles[0]], destination)).toThrow(
      "placeholder",
    );

    const alternate = replaceFirstShape(source, {
      ...requireValue(sourceSlide.shapes[0]),
      rawSidecars: [
        {
          id: asRawSidecarId("alternate"),
          node: { name: "mc:AlternateContent" },
        },
      ],
    });
    expect(() => planCrossPartDrawingMove(alternate, [handles[0]], destination)).toThrow(
      "AlternateContent",
    );

    const timingReference: PptxSourceModel = {
      ...source,
      slides: source.slides.map((slide, index) =>
        index !== 0
          ? slide
          : {
              ...slide,
              rawSidecars: [
                {
                  id: asRawSidecarId("timing"),
                  node: {
                    name: "p:timing",
                    children: [
                      {
                        name: "p:spTgt",
                        attributes: { spid: String(handles[0].nodeId) },
                      },
                    ],
                  },
                },
              ],
            },
      ),
    };
    expect(() => planCrossPartDrawingMove(timingReference, handles, destination)).toThrow(
      "may reference a moved node id",
    );

    const rawConnectorReference = replaceFirstShape(source, {
      ...requireValue(sourceSlide.shapes[0]),
      rawSidecars: [
        {
          id: asRawSidecarId("raw-connector"),
          node: {
            name: "p:stCxn",
            attributes: { id: String(handles[1].nodeId), idx: "0" },
          },
        },
      ],
    });
    expect(() => planCrossPartDrawingMove(rawConnectorReference, handles, destination)).toThrow(
      "may reference a moved node id",
    );

    const vendorReference = replaceFirstShape(source, {
      ...requireValue(sourceSlide.shapes[0]),
      rawSidecars: [
        {
          id: asRawSidecarId("vendor-reference"),
          node: {
            name: "vendor:target",
            attributes: { shapeId: String(handles[1].nodeId) },
          },
        },
      ],
    });
    expect(() => planCrossPartDrawingMove(vendorReference, handles, destination)).toThrow(
      "may reference a moved node id",
    );

    const rawNode: SourceShapeNode = {
      kind: "raw",
      nodeId: requireValue(sourceSlide.shapes[0]?.nodeId),
      handle: handles[0],
      raw: { id: asRawSidecarId("raw"), node: { name: "p:contentPart" } },
    };
    const raw = replaceFirstShape(source, rawNode);
    expect(() => planCrossPartDrawingMove(raw, [handles[0]], destination)).toThrow("raw");

    const pending: PptxSourceModel = {
      ...source,
      edits: [
        {
          kind: "moveShapes",
          targetPartPath: sourceSlide.partPath,
          shapeIds: [String(handles[0].nodeId)],
        },
      ],
    };
    expect(() => planCrossPartDrawingMove(pending, handles, destination)).toThrow("pending edits");

    const maxDestinationId: PptxSourceModel = {
      ...source,
      slides: source.slides.map((slide, index) =>
        index !== 1
          ? slide
          : {
              ...slide,
              shapes: [
                {
                  ...requireValue(sourceSlide.shapes[0]),
                  nodeId: asSourceNodeId("4294967295"),
                  handle: {
                    ...requireValue(sourceSlide.shapes[0]?.handle),
                    partPath: slide.partPath,
                    nodeId: asSourceNodeId("4294967295"),
                  },
                },
              ],
            },
      ),
    };
    expect(
      planCrossPartDrawingMove(maxDestinationId, handles, destination).nodeIdMappings[0]?.after,
    ).toBe("1");
    expect(source).toEqual(before);
    expect(source.edits).toBeUndefined();
  });
});

type AuthoringSession = ReturnType<typeof createPptxAuthoringSession>;

function buildCleanTwoSlideSource(
  author: (
    session: AuthoringSession,
    firstSlide: SourceHandle,
    secondSlide: SourceHandle,
    layout: SourceHandle,
  ) => void,
): PptxSourceModel {
  const initial = createPptx();
  const session = createPptxAuthoringSession(initial);
  const firstSlide = requireValue(initial.slides[0]?.handle);
  const layout = requireValue(initial.slideLayouts[0]?.handle);
  const secondSlide = session.addEmptySlideFromLayout({ layoutPartPath: layout.partPath });
  author(session, firstSlide, secondSlide, layout);
  return readPptx(writePptx(session.source));
}

function rectInput(offsetX: number) {
  return {
    geometry: { kind: "preset" as const, preset: "rect" },
    offsetX: asEmu(offsetX),
    offsetY: asEmu(0),
    width: asEmu(1000),
    height: asEmu(1000),
  };
}

function replaceFirstShape(source: PptxSourceModel, replacement: SourceShapeNode): PptxSourceModel {
  const slide = requireValue(source.slides[0]);
  return {
    ...source,
    slides: [
      { ...slide, shapes: [replacement, ...slide.shapes.slice(1)] },
      ...source.slides.slice(1),
    ],
  };
}

function packagePartExists(source: PptxSourceModel, partPath: string): boolean {
  return (
    source.packageGraph.parts.some((part) => part.partPath === partPath) ||
    source.packageGraph.media.some((part) => part.partPath === partPath) ||
    (source.packageGraph.rawParts ?? []).some((part) => part.partPath === partPath)
  );
}

function requireValue<T>(value: T | undefined): T {
  if (value === undefined) throw new Error("fixture value is missing");
  return value;
}
