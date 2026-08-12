import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { createPptx } from "../builder/create-pptx.js";
import {
  addChart,
  addConnector,
  addPicture,
  addShape,
  asEmu,
  asPartPath,
  asRawSidecarId,
  createPptxAuthoringSession,
  groupShapes,
  moveShapes,
  type PptxSourceModel,
  readPptx,
  type SourceGroup,
  type SourceHandle,
  type SourceShapeNode,
  writePptx,
} from "../index.js";
import { BLUE_PNG } from "../writer/write-pptx.test-helpers.js";
import {
  buildSourceGroupMappingMatrix,
  buildSourceTransformMatrix,
  multiplySourceAffineMatrices,
  type SourceAffineMatrix,
  type SourceTransformMatrixResult,
} from "./group-transform-matrix.js";
import { resolveInternalRelationshipTarget } from "./package-paths.js";

describe("moveShapes", () => {
  it("moves single and multi-node blocks at slide, layout, and master roots in source order", () => {
    const source = createPptx();
    const session = createPptxAuthoringSession(source);
    const targetHandles = [
      requireValue(source.slides[0]?.handle),
      requireValue(source.slideLayouts[0]?.handle),
      requireValue(source.slideMasters[0]?.handle),
    ];

    for (const targetHandle of targetHandles) {
      const target = session.target(targetHandle);
      const first = addRect(target, 0);
      const second = addRect(target, 1000);
      const third = addRect(target, 2000);
      const fourth = addRect(target, 3000);

      target.moveShapes([third, second], { beforeShapeHandle: first });
      expect(directChildren(session.source, targetHandle).map((shape) => shape.nodeId)).toEqual([
        second.nodeId,
        third.nodeId,
        first.nodeId,
        fourth.nodeId,
      ]);
      target.moveShapes([first]);
      expect(directChildren(session.source, targetHandle).map((shape) => shape.nodeId)).toEqual([
        second.nodeId,
        third.nodeId,
        fourth.nodeId,
        first.nodeId,
      ]);
    }

    const reread = readPptx(writePptx(session.source));
    for (const collection of [reread.slides, reread.slideLayouts, reread.slideMasters]) {
      expect(collection[0]?.shapes.map((shape) => Number(shape.nodeId))).toEqual([2, 3, 4, 1]);
    }
  });

  it("immutably patches only one nested native group and preserves transforms and handles", () => {
    const source = createPptx();
    const session = createPptxAuthoringSession(source);
    const root = session.target(requireValue(source.slides[0]?.handle));
    const first = addRect(root, 0);
    const second = addRect(root, 1000);
    const third = addRect(root, 2000);
    const sibling = addRect(root, 3000);
    const inner = root.groupShapes([first, second, third]);
    root.groupShapes([inner, sibling]);
    const beforeSlide = requireValue(session.source.slides[0]);
    const beforeOuter = requireGroup(beforeSlide.shapes[0]);
    const beforeInner = requireGroup(beforeOuter.children[0]);
    const beforeSibling = requireValue(beforeOuter.children[1]);

    session.target(inner).moveShapes([third], { beforeShapeHandle: first });

    const afterSlide = requireValue(session.source.slides[0]);
    const afterOuter = requireGroup(afterSlide.shapes[0]);
    const afterInner = requireGroup(afterOuter.children[0]);
    expect(afterSlide).not.toBe(beforeSlide);
    expect(afterOuter).not.toBe(beforeOuter);
    expect(afterInner).not.toBe(beforeInner);
    expect(afterOuter.children[1]).toBe(beforeSibling);
    expect(afterInner.children.map((shape) => shape.nodeId)).toEqual([
      third.nodeId,
      first.nodeId,
      second.nodeId,
    ]);
    expect(afterInner.children.map((shape) => shape.handle)).toEqual([third, first, second]);
    expect(afterInner.transform).toBe(beforeInner.transform);
    expect(afterInner.childTransform).toBe(beforeInner.childTransform);

    const rereadInner = requireGroup(
      requireGroup(readPptx(writePptx(session.source)).slides[0]?.shapes[0]).children[0],
    );
    expect(rereadInner.children.map((shape) => shape.nodeId)).toEqual([
      third.nodeId,
      first.nodeId,
      second.nodeId,
    ]);
    expect(rereadInner.transform).toEqual(beforeInner.transform);
    expect(rereadInner.childTransform).toEqual(beforeInner.childTransform);
  });

  it("moves blocks between roots and identity groups on slide, layout, and master parts", () => {
    const initial = createPptx();
    const session = createPptxAuthoringSession(initial);
    for (const targetHandle of [
      requireValue(initial.slides[0]?.handle),
      requireValue(initial.slideLayouts[0]?.handle),
      requireValue(initial.slideMasters[0]?.handle),
    ]) {
      const target = session.target(targetHandle);
      const first = addRect(target, 0);
      const second = addRect(target, 1000);
      const third = addRect(target, 2000);
      const group = target.groupShapes([first, second]);
      target.moveShapes([third], { beforeShapeHandle: group });
      session.target(group).moveShapes([third]);
      expect(directChildren(session.source, group).map(shapeIdentity)).toEqual([
        String(first.nodeId),
        String(second.nodeId),
        String(third.nodeId),
      ]);
      session.target(targetHandle).moveShapes([third]);
      expect(directChildren(session.source, targetHandle).at(-1)?.nodeId).toBe(third.nodeId);
    }

    const reread = readPptx(writePptx(session.source));
    for (const root of [reread.slides[0], reread.slideLayouts[0], reread.slideMasters[0]]) {
      expect(root?.shapes.at(-1)?.kind).toBe("shape");
      expect(requireGroup(root?.shapes[0]).children).toHaveLength(2);
    }
  });

  it("preserves connector endpoints and picture/chart relationships and package parts", () => {
    let source = createPptx();
    const slideHandle = requireValue(source.slides[0]?.handle);
    source = addShape(source, slideHandle, shapeInput(0));
    source = addShape(source, slideHandle, shapeInput(1000));
    const [first, second] = requireValue(source.slides[0]).shapes.map((shape) =>
      requireValue(shape.handle),
    );
    source = addConnector(source, slideHandle, {
      preset: "straightConnector1",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1),
      start: { shapeHandle: first, connectionSiteIndex: 1 },
      end: { shapeHandle: second, connectionSiteIndex: 3 },
    });
    source = addPicture(source, slideHandle, {
      bytes: BLUE_PNG,
      offsetX: asEmu(2000),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
    });
    source = addChart(source, slideHandle, {
      chartType: "bar",
      series: [{ name: "Series", categories: ["A"], values: [1] }],
      offsetX: asEmu(3000),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
    });
    source = addShape(source, slideHandle, shapeInput(4000));
    source = addShape(source, slideHandle, shapeInput(5000));
    const destinationMembers = requireValue(source.slides[0])
      .shapes.slice(-2)
      .map((shape) => requireValue(shape.handle));
    source = groupShapes(source, destinationMembers);
    const beforeGraph = source.packageGraph;
    const slide = requireValue(source.slides[0]);
    const connector = requireValue(slide.shapes.find((shape) => shape.kind === "connector"));
    const picture = requireValue(slide.shapes.find((shape) => shape.kind === "image"));
    const chart = requireValue(slide.shapes.find((shape) => shape.kind === "chart"));
    const destinationGroup = requireValue(slide.shapes.find((shape) => shape.kind === "group"));
    const slideRelationships = requireValue(
      source.packageGraph.relationships.find((group) => group.sourcePartPath === slide.partPath),
    );
    const pictureRelationship = requireValue(
      slideRelationships.relationships.find(
        (relationship) => relationship.id === picture.handle?.relationshipId,
      ),
    );
    const chartRelationship = requireValue(
      slideRelationships.relationships.find(
        (relationship) => relationship.id === chart.handle?.relationshipId,
      ),
    );
    const picturePath = resolveInternalRelationshipTarget(slide.partPath, pictureRelationship);
    const chartPath = resolveInternalRelationshipTarget(slide.partPath, chartRelationship);
    const workbookRelationship = requireValue(
      source.packageGraph.relationships
        .find((group) => group.sourcePartPath === chartPath)
        ?.relationships.find((relationship) => relationship.type.endsWith("/package")),
    );
    const workbookPath = resolveInternalRelationshipTarget(chartPath, workbookRelationship);
    const beforeArchive = unzipSync(writePptx(source));
    const edited = moveShapes(
      source,
      [requireValue(chart.handle), requireValue(picture.handle)],
      requireValue(destinationGroup.handle),
    );

    expect(edited.packageGraph).toBe(beforeGraph);
    const reread = readPptx(writePptx(edited));
    const rereadConnector = requireValue(
      reread.slides[0]?.shapes.find((shape) => shape.kind === "connector"),
    );
    const rereadDestinationGroup = requireGroup(
      reread.slides[0]?.shapes.find((shape) => shape.kind === "group"),
    );
    const rereadPicture = requireValue(
      rereadDestinationGroup.children.find((shape) => shape.kind === "image"),
    );
    const rereadChart = requireValue(
      rereadDestinationGroup.children.find((shape) => shape.kind === "chart"),
    );
    expect(rereadConnector.kind).toBe("connector");
    if (rereadConnector.kind !== "connector") throw new Error("connector fixture was lost");
    expect(rereadConnector.connection).toEqual(
      connector.kind === "connector" ? connector.connection : undefined,
    );
    expect(rereadPicture.kind === "image" ? rereadPicture.blipRelationshipId : undefined).toBe(
      picture.kind === "image" ? picture.blipRelationshipId : undefined,
    );
    expect(rereadChart.kind === "chart" ? rereadChart.chartRelationshipId : undefined).toBe(
      chart.kind === "chart" ? chart.chartRelationshipId : undefined,
    );
    expect(reread.packageGraph.parts.map((part) => part.partPath).sort()).toEqual(
      source.packageGraph.parts.map((part) => part.partPath).sort(),
    );
    const rereadSlideRelationships = requireValue(
      reread.packageGraph.relationships.find((group) => group.sourcePartPath === slide.partPath),
    );
    expect(
      relationshipIdentity(
        requireValue(
          rereadSlideRelationships.relationships.find(
            (relationship) => relationship.id === pictureRelationship.id,
          ),
        ),
      ),
    ).toEqual(relationshipIdentity(pictureRelationship));
    expect(
      relationshipIdentity(
        requireValue(
          rereadSlideRelationships.relationships.find(
            (relationship) => relationship.id === chartRelationship.id,
          ),
        ),
      ),
    ).toEqual(relationshipIdentity(chartRelationship));
    expect(
      relationshipIdentity(
        requireValue(
          reread.packageGraph.relationships
            .find((group) => group.sourcePartPath === chartPath)
            ?.relationships.find((relationship) => relationship.id === workbookRelationship.id),
        ),
      ),
    ).toEqual(relationshipIdentity(workbookRelationship));

    const afterArchive = unzipSync(writePptx(edited));
    for (const partPath of [picturePath, chartPath, workbookPath]) {
      expect(afterArchive[partPath], `missing preserved part ${partPath}`).toBeDefined();
      expect(afterArchive[partPath]).toEqual(beforeArchive[partPath]);
    }
    expect(reread.packageGraph.parts).toEqual(
      expect.arrayContaining(
        source.packageGraph.parts.filter((part) =>
          [picturePath, chartPath, workbookPath].includes(part.partPath),
        ),
      ),
    );
  });

  it("allows a connector closure to move together across an identity-mapped parent", () => {
    let source = createShapes(4);
    const slide = requireValue(source.slides[0]);
    const slideHandle = requireValue(slide.handle);
    const handles = slide.shapes.map((shape) => requireValue(shape.handle));
    source = addConnector(source, slideHandle, {
      preset: "straightConnector1",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1),
      start: { shapeHandle: handles[0], connectionSiteIndex: 1 },
      end: { shapeHandle: handles[1], connectionSiteIndex: 3 },
    });
    const connector = requireValue(source.slides[0]?.shapes.at(-1));
    const connectorHandle = requireValue(connector.handle);
    source = moveShapes(source, [connectorHandle], slideHandle, { beforeShapeHandle: handles[2] });
    source = groupShapes(source, handles.slice(2));
    const group = requireGroup(source.slides[0]?.shapes.at(-1));
    const edited = moveShapes(
      source,
      [connectorHandle, handles[1], handles[0]],
      requireValue(group.handle),
      {
        beforeShapeHandle: handles[2],
      },
    );
    const reread = readPptx(writePptx(edited));
    const rereadGroup = requireGroup(reread.slides[0]?.shapes[0]);
    expect(rereadGroup.children.map(shapeIdentity)).toEqual([
      String(handles[0].nodeId),
      String(handles[1].nodeId),
      String(connectorHandle.nodeId),
      String(handles[2].nodeId),
      String(handles[3].nodeId),
    ]);
    const rereadConnector = rereadGroup.children[2];
    expect(rereadConnector?.kind).toBe("connector");
    if (rereadConnector?.kind !== "connector") throw new Error("moved connector was lost");
    expect(rereadConnector.connection?.start?.shapeId).toBe(handles[0].nodeId);
    expect(rereadConnector.connection?.end?.shapeId).toBe(handles[1].nodeId);
  });

  it("moves typed shapes when preserved raw part XML is unavailable", () => {
    let source = createShapes(3);
    const slide = requireValue(source.slides[0]);
    const [first, second, third] = slide.shapes.map((shape) => requireValue(shape.handle));
    source = groupShapes(source, [second, third]);
    const group = requireGroup(source.slides[0]?.shapes.at(-1));
    source = {
      ...source,
      packageGraph: { ...source.packageGraph, rawParts: undefined },
    };

    const edited = moveShapes(source, [first], requireValue(group.handle), {
      beforeShapeHandle: second,
    });

    expect(requireGroup(edited.slides[0]?.shapes[0]).children.map(shapeIdentity)).toEqual([
      String(first.nodeId),
      String(second.nodeId),
      String(third.nodeId),
    ]);
  });

  it("keeps non-drawing XML in its authored slot while moving a block", () => {
    const authored = createShapes(3);
    const slidePath = "ppt/slides/slide1.xml";
    const archive = unzipSync(writePptx(authored));
    const inputXml = new TextDecoder().decode(requireValue(archive[slidePath]));
    archive[slidePath] = new TextEncoder().encode(
      inputXml.replace(
        /(<p:sp>[\s\S]*?<\/p:sp>)(<p:sp>)/,
        '$1<p:extLst><p:ext uri="slot"/></p:extLst>$2',
      ),
    );
    const source = readPptx(zipSync(archive));
    const slide = requireValue(source.slides[0]);
    const [first, , third] = slide.shapes.map((shape) => requireValue(shape.handle));
    const edited = moveShapes(source, [third], requireValue(slide.handle), {
      beforeShapeHandle: first,
    });
    const outputXml = new TextDecoder().decode(
      requireValue(unzipSync(writePptx(edited))[slidePath]),
    );

    expectXmlOrder(outputXml, [`id="${third.nodeId}"`, 'uri="slot"', `id="${first.nodeId}"`]);
  });

  it("keeps raw drawing siblings in z-order during a forward move and rejects forged gaps", () => {
    const input = insertAfterFirstShape(
      createShapes(3),
      '<p:contentPart data-marker="raw-drawing"><p:nvContentPartPr><p:cNvPr id="99" name="Raw Drawing"/></p:nvContentPartPr></p:contentPart>',
    );
    const source = readPptx(input);
    const slide = requireValue(source.slides[0]);
    const [first, raw, second, third] = slide.shapes;
    expect(raw?.kind).toBe("raw");
    const edited = moveShapes(source, [requireValue(first?.handle)], requireValue(slide.handle), {
      beforeShapeHandle: requireValue(third?.handle),
    });
    expect(edited.slides[0]?.shapes.map(shapeIdentity)).toEqual([
      "raw:contentPart",
      String(second?.nodeId),
      String(first?.nodeId),
      String(third?.nodeId),
    ]);

    const output = writePptx(edited);
    const outputXml = slideXml(output);
    expectXmlOrder(outputXml, [
      'data-marker="raw-drawing"',
      `id="${second?.nodeId}"`,
      `id="${first?.nodeId}"`,
      `id="${third?.nodeId}"`,
    ]);
    expect(readPptx(output).slides[0]?.shapes.map(shapeIdentity)).toEqual(
      edited.slides[0]?.shapes.map(shapeIdentity),
    );
    expect(extractElementContaining(slideXml(output), "p:contentPart", "raw-drawing")).toBe(
      extractElementContaining(slideXml(input), "p:contentPart", "raw-drawing"),
    );

    const forged: PptxSourceModel = {
      ...source,
      edits: [
        {
          kind: "moveShapes",
          targetPartPath: slide.partPath,
          shapeIds: [String(first?.nodeId), String(second?.nodeId)],
        },
      ],
    };
    expect(() => writePptx(forged)).toThrow("not consecutive");
  });

  it("treats unknown future drawing children as z-order siblings and rejects forged gaps", () => {
    const input = insertAfterFirstShape(
      createShapes(3),
      '<p15:extLst xmlns:p15="urn:future-drawing" data-marker="future-drawing" p15:flag="keep"><p15:payload value="42">raw payload</p15:payload></p15:extLst>',
    );
    const source = readPptx(input);
    const slide = requireValue(source.slides[0]);
    const [first, raw, second, third] = slide.shapes;
    expect(raw?.kind).toBe("raw");
    expect(shapeIdentity(requireValue(raw))).toBe("raw:extLst");

    const edited = moveShapes(source, [requireValue(first?.handle)], requireValue(slide.handle), {
      beforeShapeHandle: requireValue(third?.handle),
    });
    const expectedOrder = [
      "raw:extLst",
      String(second?.nodeId),
      String(first?.nodeId),
      String(third?.nodeId),
    ];
    expect(edited.slides[0]?.shapes.map(shapeIdentity)).toEqual(expectedOrder);

    const output = writePptx(edited);
    const outputXml = slideXml(output);
    expectXmlOrder(outputXml, [
      'data-marker="future-drawing"',
      `id="${second?.nodeId}"`,
      `id="${first?.nodeId}"`,
      `id="${third?.nodeId}"`,
    ]);
    const reread = readPptx(output);
    expect(reread.slides[0]?.shapes.map(shapeIdentity)).toEqual(expectedOrder);
    const rereadRaw = requireValue(reread.slides[0]?.shapes[0]);
    expect(rereadRaw.kind).toBe("raw");
    if (raw?.kind !== "raw" || rereadRaw.kind !== "raw") {
      throw new Error("future drawing fixture was not preserved as raw XML");
    }
    expect(rereadRaw.raw.node).toEqual(raw.raw.node);
    expect(extractElementContaining(outputXml, "p15:extLst", "future-drawing")).toBe(
      extractElementContaining(slideXml(input), "p15:extLst", "future-drawing"),
    );

    const forged: PptxSourceModel = {
      ...source,
      edits: [
        {
          kind: "moveShapes",
          targetPartPath: slide.partPath,
          shapeIds: [String(first?.nodeId), String(second?.nodeId)],
        },
      ],
    };
    expect(() => writePptx(forged)).toThrow("not consecutive");
  });

  it("keeps idless drawing siblings in z-order during an end move and rejects forged gaps", () => {
    const input = insertAfterFirstShape(createShapes(3), idlessShapeXml(createShapes(1)));
    const source = readPptx(input);
    const slide = requireValue(source.slides[0]);
    const [first, idless, second, third] = slide.shapes;
    expect(idless?.kind).toBe("shape");
    expect(idless?.nodeId).toBeUndefined();
    const edited = moveShapes(source, [requireValue(first?.handle)], requireValue(slide.handle));
    expect(edited.slides[0]?.shapes.map(shapeIdentity)).toEqual([
      "idless:shape",
      String(second?.nodeId),
      String(third?.nodeId),
      String(first?.nodeId),
    ]);

    const output = writePptx(edited);
    const outputXml = slideXml(output);
    expectXmlOrder(outputXml, [
      'name="Idless Drawing"',
      `id="${second?.nodeId}"`,
      `id="${third?.nodeId}"`,
      `id="${first?.nodeId}"`,
    ]);
    expect(readPptx(output).slides[0]?.shapes.map(shapeIdentity)).toEqual(
      edited.slides[0]?.shapes.map(shapeIdentity),
    );
    expect(extractElementContaining(outputXml, "p:sp", "Idless Drawing")).toBe(
      extractElementContaining(slideXml(input), "p:sp", "Idless Drawing"),
    );

    const forged: PptxSourceModel = {
      ...source,
      edits: [
        {
          kind: "moveShapes",
          targetPartPath: slide.partPath,
          shapeIds: [String(first?.nodeId), String(second?.nodeId)],
        },
      ],
    };
    expect(() => writePptx(forged)).toThrow("not consecutive");
  });

  it("atomically rejects invalid blocks, anchors, parts, ids, and unsupported targets", () => {
    const source = createShapes(4);
    const slide = requireValue(source.slides[0]);
    const slideHandle = requireValue(slide.handle);
    const handles = slide.shapes.map((shape) => requireValue(shape.handle));
    const before = structuredClone(source);
    const beforeEdits = source.edits;

    const cases: (() => unknown)[] = [
      () => moveShapes(source, [], slideHandle),
      () => moveShapes(source, [handles[0], handles[2]], slideHandle),
      () => moveShapes(source, [handles[0], handles[0]], slideHandle),
      () => moveShapes(source, [{ ...handles[0], nodeId: undefined }], slideHandle),
      () =>
        moveShapes(
          source,
          [{ ...handles[0], partPath: asPartPath("ppt/slides/other.xml") }],
          slideHandle,
        ),
      () => moveShapes(source, [handles[0]], slideHandle, { beforeShapeHandle: handles[0] }),
      () =>
        moveShapes(source, [handles[0]], slideHandle, {
          beforeShapeHandle: { ...handles[1], nodeId: undefined },
        }),
      () =>
        moveShapes(source, [handles[0]], slideHandle, {
          beforeShapeHandle: { ...handles[1], partPath: asPartPath("ppt/slides/other.xml") },
        }),
    ];
    for (const reject of cases) expect(reject).toThrow();
    expect(source).toEqual(before);
    expect(source.edits).toBe(beforeEdits);

    const unsupported = ["smartArt", "raw"] as const;
    for (const kind of unsupported) {
      const target = kind === "smartArt" ? asSmartArt(slide.shapes[0]) : asRaw(slide.shapes[0]);
      const invalid = replaceFirstShape(source, target);
      expect(() => moveShapes(invalid, [requireValue(target.handle)], slideHandle)).toThrow(
        "SmartArt and raw",
      );
      expect(invalid.edits).toBe(source.edits);
    }

    const alternate = replaceFirstShape(source, {
      ...requireValue(slide.shapes[0]),
      rawSidecars: [{ id: asRawSidecarId("alternate"), node: { name: "mc:AlternateContent" } }],
    });
    expect(() => moveShapes(alternate, [handles[0]], slideHandle)).toThrow("AlternateContent");
    expect(alternate.edits).toBe(source.edits);
  });

  it("rejects non-direct children and forged writer journals atomically", () => {
    const source = createShapes(4);
    const slide = requireValue(source.slides[0]);
    const handles = slide.shapes.map((shape) => requireValue(shape.handle));
    const grouped = groupShapes(source, handles.slice(0, 2));
    const group = requireGroup(grouped.slides[0]?.shapes[0]);
    const groupHandle = requireValue(group.handle);

    const movedToRoot = moveShapes(grouped, [handles[0]], requireValue(slide.handle));
    expect(movedToRoot.slides[0]?.shapes.map(shapeIdentity)).toEqual([
      String(groupHandle.nodeId),
      String(handles[2].nodeId),
      String(handles[3].nodeId),
      String(handles[0].nodeId),
    ]);
    expect(() =>
      moveShapes(grouped, [requireValue(group.children[0]?.handle)], groupHandle, {
        beforeShapeHandle: handles[2],
      }),
    ).toThrow("direct child");

    const valid = moveShapes(grouped, [handles[2]], requireValue(slide.handle));
    const edit = requireValue(valid.edits?.at(-1));
    if (edit.kind !== "moveShapes") throw new Error("move edit fixture is missing");
    const forgedEdits = [
      { ...edit, shapeIds: [] },
      { ...edit, shapeIds: [String(handles[2].nodeId), String(handles[2].nodeId)] },
      { ...edit, shapeIds: [String(groupHandle.nodeId), String(handles[3].nodeId)] },
      { ...edit, beforeShapeId: String(handles[2].nodeId) },
      { ...edit, beforeShapeId: "999" },
    ];
    for (const forged of forgedEdits) {
      const forgedSource: PptxSourceModel = {
        ...grouped,
        edits: [...(grouped.edits ?? []), forged],
      };
      expect(() => writePptx(forgedSource)).toThrow();
    }
  });

  it("moves root/group and group/group blocks while preserving identity and empty groups", () => {
    const source = createShapes(4);
    const handles = requireValue(source.slides[0]).shapes.map((shape) =>
      requireValue(shape.handle),
    );
    const withGroupA = groupShapes(source, handles.slice(0, 2));
    const withBothGroups = groupShapes(withGroupA, handles.slice(2));
    const [groupA, groupB] = requireValue(withBothGroups.slides[0]).shapes.map(requireGroup);
    const firstHandle = requireValue(groupA.children[0]?.handle);
    const secondHandle = requireValue(groupA.children[1]?.handle);
    const groupBHandle = requireValue(groupB.handle);
    const groupToGroup = moveShapes(withBothGroups, [secondHandle, firstHandle], groupBHandle, {
      beforeShapeHandle: requireValue(groupB.children[0]?.handle),
    });
    const [emptyA, filledB] = requireValue(groupToGroup.slides[0]).shapes.map(requireGroup);
    expect(emptyA.children).toEqual([]);
    expect(filledB.children.map(shapeIdentity)).toEqual([
      String(firstHandle.nodeId),
      String(secondHandle.nodeId),
      String(handles[2].nodeId),
      String(handles[3].nodeId),
    ]);
    expect(filledB.children.slice(0, 2).map((shape) => shape.handle)).toEqual([
      firstHandle,
      secondHandle,
    ]);

    const slideHandle = requireValue(groupToGroup.slides[0]?.handle);
    const groupToRoot = moveShapes(groupToGroup, [firstHandle, secondHandle], slideHandle);
    expect(requireGroup(groupToRoot.slides[0]?.shapes[0]).children).toEqual([]);
    expect(groupToRoot.slides[0]?.shapes.slice(-2).map(shapeIdentity)).toEqual([
      String(firstHandle.nodeId),
      String(secondHandle.nodeId),
    ]);

    const rootToGroup = moveShapes(groupToRoot, [firstHandle, secondHandle], groupBHandle);
    const reread = readPptx(writePptx(rootToGroup));
    const rereadGroups = requireValue(reread.slides[0]).shapes.map(requireGroup);
    expect(rereadGroups[0]?.children).toEqual([]);
    expect(rereadGroups[1]?.children.map(shapeIdentity)).toEqual([
      String(handles[2].nodeId),
      String(handles[3].nodeId),
      String(firstHandle.nodeId),
      String(secondHandle.nodeId),
    ]);
    expect(rereadGroups[1]?.children.at(-2)?.transform).toEqual(groupA.children[0]?.transform);
    expect(rereadGroups[1]?.children.at(-1)?.transform).toEqual(groupA.children[1]?.transform);
  });

  it("re-expresses root/group, group/root, and group/group transforms across representable affine parents", () => {
    let source = createPptx();
    const slideHandle = requireValue(source.slides[0]?.handle);
    for (const offsetX of [2000, 6000, 10_000, 14_000, 18_000]) {
      source = addShape(source, slideHandle, shapeInput(offsetX));
    }
    const handles = requireValue(source.slides[0]).shapes.map((shape) =>
      requireValue(shape.handle),
    );
    source = groupShapes(source, handles.slice(1, 3));
    source = groupShapes(source, handles.slice(3, 5));
    const [rootShape, sourceGroup, destinationGroup] = requireValue(source.slides[0]).shapes;
    const sourceGroupId = String(requireValue(sourceGroup?.nodeId));
    const destinationGroupId = String(requireValue(destinationGroup?.nodeId));
    source = materializeGroupTransforms(
      source,
      new Map([
        [
          sourceGroupId,
          `<a:xfrm rot="5400000" flipV="1"><a:off x="2000" y="3000"/><a:ext cx="4000" cy="3000"/><a:chOff x="4000" y="0"/><a:chExt cx="2000" cy="1000"/></a:xfrm>`,
        ],
        [
          destinationGroupId,
          `<a:xfrm flipH="1"><a:off x="10000" y="2000"/><a:ext cx="6000" cy="2000"/><a:chOff x="12000" y="0"/><a:chExt cx="3000" cy="1000"/></a:xfrm>`,
        ],
      ]),
    );

    const materializedSlide = requireValue(source.slides[0]);
    const materializedRoot = requireValue(
      materializedSlide.shapes.find((shape) => shape.nodeId === rootShape?.nodeId),
    );
    const materializedSourceGroup = requireGroup(
      materializedSlide.shapes.find((shape) => String(shape.nodeId) === sourceGroupId),
    );
    const materializedDestinationGroup = requireGroup(
      materializedSlide.shapes.find((shape) => String(shape.nodeId) === destinationGroupId),
    );
    const sourceChild = requireValue(materializedSourceGroup.children[0]);
    const destinationHandle = requireValue(materializedDestinationGroup.handle);

    const rootBefore = valueOf(buildSourceTransformMatrix(requireTransform(materializedRoot)));
    const rootToGroup = moveShapes(
      source,
      [requireValue(materializedRoot.handle)],
      destinationHandle,
    );
    const rootAfter = requireValue(
      requireGroup(
        rootToGroup.slides[0]?.shapes.find((shape) => String(shape.nodeId) === destinationGroupId),
      ).children.find((shape) => shape.nodeId === materializedRoot.nodeId),
    );
    expectMatrixClose(
      multiplySourceAffineMatrices(
        valueOf(
          buildSourceGroupMappingMatrix(
            materializedDestinationGroup.transform,
            materializedDestinationGroup.childTransform,
          ),
        ),
        valueOf(buildSourceTransformMatrix(requireTransform(rootAfter))),
      ),
      rootBefore,
    );

    const sourceChildBefore = multiplySourceAffineMatrices(
      valueOf(
        buildSourceGroupMappingMatrix(
          materializedSourceGroup.transform,
          materializedSourceGroup.childTransform,
        ),
      ),
      valueOf(buildSourceTransformMatrix(requireTransform(sourceChild))),
    );
    const groupToRoot = moveShapes(
      rootToGroup,
      [requireValue(sourceChild.handle)],
      requireValue(materializedSlide.handle),
    );
    const sourceChildAtRoot = requireValue(
      groupToRoot.slides[0]?.shapes.find((shape) => shape.nodeId === sourceChild.nodeId),
    );
    expectMatrixClose(
      valueOf(buildSourceTransformMatrix(requireTransform(sourceChildAtRoot))),
      sourceChildBefore,
    );

    const secondSourceChild = requireValue(materializedSourceGroup.children[1]);
    const secondBefore = multiplySourceAffineMatrices(
      valueOf(
        buildSourceGroupMappingMatrix(
          materializedSourceGroup.transform,
          materializedSourceGroup.childTransform,
        ),
      ),
      valueOf(buildSourceTransformMatrix(requireTransform(secondSourceChild))),
    );
    const groupToGroup = moveShapes(
      groupToRoot,
      [requireValue(secondSourceChild.handle)],
      destinationHandle,
    );
    const generalEdit = requireValue(groupToGroup.edits?.at(-1));
    if (generalEdit.kind !== "moveShapes" || generalEdit.transformedRoots === undefined) {
      throw new Error("general affine move edit fixture is missing");
    }
    const finalized = requireValue(generalEdit.transformedRoots[0]);
    const forged: PptxSourceModel = {
      ...groupToRoot,
      edits: [
        ...(groupToRoot.edits ?? []),
        {
          ...generalEdit,
          transformedRoots: [
            {
              ...finalized,
              transform: {
                ...finalized.transform,
                offsetX: asEmu(Number(finalized.transform.offsetX) + 1),
              },
            },
          ],
        },
      ],
    };
    expect(() => writePptx(forged)).toThrow("finalized transform is stale");

    const reread = readPptx(writePptx(groupToGroup));
    const rereadDestination = requireGroup(
      reread.slides[0]?.shapes.find((shape) => String(shape.nodeId) === destinationGroupId),
    );
    const rereadMoved = requireValue(
      rereadDestination.children.find((shape) => shape.nodeId === secondSourceChild.nodeId),
    );
    expectMatrixClose(
      multiplySourceAffineMatrices(
        valueOf(
          buildSourceGroupMappingMatrix(
            rereadDestination.transform,
            rereadDestination.childTransform,
          ),
        ),
        valueOf(buildSourceTransformMatrix(requireTransform(rereadMoved))),
      ),
      secondBefore,
    );
    expect(rereadMoved.handle?.partPath).toBe(secondSourceChild.handle?.partPath);
    expect(rereadMoved.handle?.nodeId).toBe(secondSourceChild.handle?.nodeId);
  });

  it("re-expresses only a moved group root and preserves its complete subtree", () => {
    let source = createShapes(6);
    const initialSlide = requireValue(source.slides[0]);
    const handles = initialSlide.shapes.map((shape) => requireValue(shape.handle));
    source = groupShapes(source, handles.slice(0, 2));
    const innerHandle = requireValue(requireGroup(source.slides[0]?.shapes[0]).handle);
    source = groupShapes(source, [innerHandle, handles[2]]);
    source = groupShapes(source, handles.slice(3, 5));

    const groupedSlide = requireValue(source.slides[0]);
    const sourceParent = requireGroup(groupedSlide.shapes[0]);
    const destinationParent = requireGroup(groupedSlide.shapes[1]);
    const rootSibling = requireValue(groupedSlide.shapes[2]);
    source = materializeGroupTransforms(
      source,
      new Map([
        [
          String(sourceParent.nodeId),
          `<a:xfrm><a:off x="2000" y="3000"/><a:ext cx="4000" cy="3000"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="1000"/></a:xfrm>`,
        ],
        [
          String(destinationParent.nodeId),
          `<a:xfrm><a:off x="10000" y="2000"/><a:ext cx="3000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="3000" cy="1000"/></a:xfrm>`,
        ],
      ]),
    );
    source = insertIntoGroupProperties(
      source,
      String(innerHandle.nodeId),
      `<a:extLst><a:ext uri="{AFFINE-GROUP-PRESERVE}"/></a:extLst>`,
    );

    const beforeSlide = requireValue(source.slides[0]);
    const beforeSourceParent = requireGroup(beforeSlide.shapes[0]);
    const beforeDestinationParent = requireGroup(beforeSlide.shapes[1]);
    const beforeInner = requireGroup(beforeSourceParent.children[0]);
    const beforeChildren = beforeInner.children.map((child) => ({
      nodeId: child.nodeId,
      handle: requireValue(child.handle),
      transform: requireTransform(child),
    }));
    const beforeChildTransform = beforeInner.childTransform;
    const beforeXml = slideXml(writePptx(source));
    const beforeSiblingXml = extractElementContaining(
      beforeXml,
      "p:sp",
      `<p:cNvPr id="${String(rootSibling.nodeId)}"`,
    );
    const absoluteBefore = multiplySourceAffineMatrices(
      valueOf(
        buildSourceGroupMappingMatrix(
          beforeSourceParent.transform,
          beforeSourceParent.childTransform,
        ),
      ),
      valueOf(buildSourceTransformMatrix(requireTransform(beforeInner))),
    );

    const moved = moveShapes(source, [innerHandle], requireValue(beforeDestinationParent.handle));
    const reread = readPptx(writePptx(moved));
    const rereadDestination = requireGroup(reread.slides[0]?.shapes[1]);
    const rereadInner = requireGroup(
      rereadDestination.children.find((child) => child.nodeId === innerHandle.nodeId),
    );
    expectMatrixClose(
      multiplySourceAffineMatrices(
        valueOf(
          buildSourceGroupMappingMatrix(
            rereadDestination.transform,
            rereadDestination.childTransform,
          ),
        ),
        valueOf(buildSourceTransformMatrix(requireTransform(rereadInner))),
      ),
      absoluteBefore,
    );
    expect(rereadInner.childTransform).toEqual(beforeChildTransform);
    expect(
      rereadInner.children.map((child) => ({
        nodeId: child.nodeId,
        handle: requireValue(child.handle),
        transform: requireTransform(child),
      })),
    ).toEqual(beforeChildren);

    const afterXml = slideXml(writePptx(moved));
    expect(afterXml).toContain(`uri="{AFFINE-GROUP-PRESERVE}"`);
    expect(
      extractElementContaining(afterXml, "p:sp", `<p:cNvPr id="${String(rootSibling.nodeId)}"`),
    ).toBe(beforeSiblingXml);
  });

  it.each([
    [
      "shear",
      `<a:xfrm rot="2700000"><a:off x="0" y="0"/><a:ext cx="4000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="1000"/></a:xfrm>`,
      "shear",
    ],
    [
      "singular destination",
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="1000"/></a:xfrm>`,
      "singular-matrix",
    ],
    [
      "missing child extent",
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="2000" cy="1000"/><a:chOff x="0" y="0"/></a:xfrm>`,
      "missing-child-transform",
    ],
    [
      "zero child extent",
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="2000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="1000"/></a:xfrm>`,
      "zero-child-extent",
    ],
    [
      "invalid extent",
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="-1" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="1000"/></a:xfrm>`,
      "invalid-extent",
    ],
    [
      "EMU quantization mismatch",
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="3000" cy="3000"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm>`,
      "quantization-mismatch",
    ],
    [
      "OOXML coordinate overflow",
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="1000000000000" cy="1000"/></a:xfrm>`,
      "out-of-range-transform",
    ],
  ] as const)("rejects %s atomically at the move boundary", (_, transformXml, reason) => {
    let source = createShapes(3);
    const slide = requireValue(source.slides[0]);
    const handles = slide.shapes.map((shape) => requireValue(shape.handle));
    source = groupShapes(source, handles.slice(1));
    const group = requireGroup(source.slides[0]?.shapes[1]);
    source = materializeGroupTransforms(source, new Map([[String(group.nodeId), transformXml]]));
    const before = structuredClone(source);
    const materializedSlide = requireValue(source.slides[0]);
    const root = requireValue(materializedSlide.shapes[0]);
    const destination = requireGroup(materializedSlide.shapes[1]);

    expect(() =>
      moveShapes(source, [requireValue(root.handle)], requireValue(destination.handle)),
    ).toThrow(reason);
    expect(source).toEqual(before);
  });

  it("rejects cycles, non-identity ancestors, and connector boundary crossings atomically", () => {
    const source = createShapes(4);
    const slide = requireValue(source.slides[0]);
    const handles = slide.shapes.map((shape) => requireValue(shape.handle));
    const grouped = groupShapes(source, handles.slice(0, 3));
    const group = requireGroup(grouped.slides[0]?.shapes[0]);
    const nested = groupShapes(
      grouped,
      group.children.slice(0, 2).map((shape) => requireValue(shape.handle)),
    );
    const outer = requireGroup(nested.slides[0]?.shapes[0]);
    const inner = requireGroup(outer.children[0]);
    expect(() =>
      moveShapes(nested, [requireValue(outer.handle)], requireValue(inner.handle)),
    ).toThrow("inside the moved block");

    const invalidAncestors: SourceGroup[] = [
      {
        ...outer,
        transform: {
          ...requireValue(outer.transform),
          width: asEmu(Number(requireValue(outer.transform).width) + 1),
        },
      },
      { ...outer, transform: undefined },
      { ...outer, childTransform: undefined },
      {
        ...outer,
        childTransform: { ...requireValue(outer.childTransform), width: asEmu(0) },
      },
    ];
    for (const invalidAncestor of invalidAncestors) {
      const invalid: PptxSourceModel = {
        ...nested,
        slides: [
          {
            ...requireValue(nested.slides[0]),
            shapes: [invalidAncestor, ...requireValue(nested.slides[0]).shapes.slice(1)],
          },
        ],
      };
      expect(() =>
        moveShapes(invalid, [requireValue(inner.children[0]?.handle)], requireValue(slide.handle)),
      ).toThrow("affine transform is not exactly representable");
    }

    let connected = createShapes(4);
    const connectedSlide = requireValue(connected.slides[0]);
    const connectedHandles = connectedSlide.shapes.map((shape) => requireValue(shape.handle));
    connected = addConnector(connected, requireValue(connectedSlide.handle), {
      preset: "straightConnector1",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1),
      start: { shapeHandle: connectedHandles[0], connectionSiteIndex: 0 },
      end: { shapeHandle: connectedHandles[1], connectionSiteIndex: 0 },
    });
    const groupedConnected = groupShapes(connected, connectedHandles.slice(2));
    const destination = requireGroup(groupedConnected.slides[0]?.shapes[2]);
    const before = structuredClone(groupedConnected);
    expect(() =>
      moveShapes(groupedConnected, [connectedHandles[0]], requireValue(destination.handle)),
    ).toThrow("connector endpoint crosses");
    expect(groupedConnected).toEqual(before);

    const materialized = readPptx(writePptx(connected));
    const materializedSlide = requireValue(materialized.slides[0]);
    const materializedHandles = materializedSlide.shapes
      .filter((shape) => shape.kind !== "connector")
      .map((shape) => requireValue(shape.handle));
    const materializedGrouped = groupShapes(materialized, materializedHandles.slice(2));
    const rawOnly: PptxSourceModel = {
      ...materializedGrouped,
      slides: [
        {
          ...requireValue(materializedGrouped.slides[0]),
          shapes: requireValue(materializedGrouped.slides[0]).shapes.map((shape) =>
            shape.kind === "connector" ? { ...shape, connection: undefined } : shape,
          ),
        },
      ],
    };
    const rawDestination = requireGroup(rawOnly.slides[0]?.shapes[2]);
    expect(() =>
      moveShapes(rawOnly, [materializedHandles[0]], requireValue(rawDestination.handle)),
    ).toThrow("raw XML connector endpoint crosses");
    const forgedRawBoundary: PptxSourceModel = {
      ...materializedGrouped,
      edits: [
        ...(materializedGrouped.edits ?? []),
        {
          kind: "moveShapes",
          targetPartPath: requireValue(materializedGrouped.slides[0]).partPath,
          crossParent: true,
          destinationParentGroupId: String(rawDestination.nodeId),
          shapeIds: [String(materializedHandles[0].nodeId)],
        },
      ],
    };
    expect(() => writePptx(forgedRawBoundary)).toThrow("connector endpoint crosses");
  });

  it("rejects duplicate node ids without adding an edit", () => {
    const source = createShapes(2);
    const slide = requireValue(source.slides[0]);
    const first = requireValue(slide.shapes[0]);
    const second = requireValue(slide.shapes[1]);
    const duplicate = {
      ...source,
      slides: [
        {
          ...slide,
          shapes: [
            first,
            {
              ...second,
              nodeId: first.nodeId,
              handle: { ...requireValue(second.handle), nodeId: first.nodeId },
            },
          ],
        },
      ],
    } satisfies PptxSourceModel;

    expect(() =>
      moveShapes(duplicate, [requireValue(first.handle)], requireValue(slide.handle)),
    ).toThrow("duplicate node id");
    expect(duplicate.edits).toBe(source.edits);
  });
});

type Target = ReturnType<ReturnType<typeof createPptxAuthoringSession>["target"]>;

function addRect(target: Target, offsetX: number): SourceHandle {
  return target.addShape(shapeInput(offsetX));
}

function shapeInput(offsetX: number) {
  return {
    geometry: { kind: "preset" as const, preset: "rect" },
    offsetX: asEmu(offsetX),
    offsetY: asEmu(0),
    width: asEmu(1000),
    height: asEmu(1000),
  };
}

function createShapes(count: number): PptxSourceModel {
  let source = createPptx();
  const target = requireValue(source.slides[0]?.handle);
  for (let index = 0; index < count; index += 1)
    source = addShape(source, target, shapeInput(index));
  return source;
}

function insertAfterFirstShape(source: PptxSourceModel, fragment: string): Uint8Array {
  const archive = unzipSync(writePptx(source));
  const partPath = "ppt/slides/slide1.xml";
  const xml = new TextDecoder().decode(requireValue(archive[partPath]));
  archive[partPath] = new TextEncoder().encode(
    xml.replace(/(<p:sp>[\s\S]*?<\/p:sp>)/, `$1${fragment}`),
  );
  return zipSync(archive);
}

function idlessShapeXml(source: PptxSourceModel): string {
  const xml = slideXml(writePptx(source));
  const shape = requireValue(xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)?.[0]);
  return shape.replace(/ id="[^"]+"/, "").replace(/ name="[^"]*"/, ' name="Idless Drawing"');
}

function slideXml(bytes: Uint8Array): string {
  return new TextDecoder().decode(requireValue(unzipSync(bytes)["ppt/slides/slide1.xml"]));
}

function materializeGroupTransforms(
  source: PptxSourceModel,
  transforms: ReadonlyMap<string, string>,
): PptxSourceModel {
  const archive = unzipSync(writePptx(source));
  const partPath = "ppt/slides/slide1.xml";
  let xml = new TextDecoder().decode(requireValue(archive[partPath]));
  for (const [groupId, transformXml] of transforms) {
    const groupBlocks = xml.match(/<p:grpSp\b[^>]*>[\s\S]*?<\/p:grpSp>/g) ?? [];
    const idPattern = new RegExp(`<p:cNvPr[^>]*\\bid="${groupId}"`);
    const block = groupBlocks.find((candidate) => idPattern.test(candidate));
    if (block === undefined) throw new Error(`group '${groupId}' was not found`);
    const updatedBlock = block.replace(/<a:xfrm[\s\S]*?<\/a:xfrm>/, transformXml);
    const next = xml.replace(block, updatedBlock);
    if (next === xml) throw new Error(`group '${groupId}' xfrm was not found`);
    xml = next;
  }
  archive[partPath] = new TextEncoder().encode(xml);
  return readPptx(zipSync(archive));
}

function insertIntoGroupProperties(
  source: PptxSourceModel,
  groupId: string,
  fragment: string,
): PptxSourceModel {
  const archive = unzipSync(writePptx(source));
  const partPath = "ppt/slides/slide1.xml";
  const xml = new TextDecoder().decode(requireValue(archive[partPath]));
  const marker = `<p:cNvPr id="${groupId}"`;
  const groupXml = extractElementContaining(xml, "p:grpSp", marker);
  const updatedGroupXml = groupXml.replace("</p:grpSpPr>", `${fragment}</p:grpSpPr>`);
  if (updatedGroupXml === groupXml) throw new Error(`group '${groupId}' properties were not found`);
  archive[partPath] = new TextEncoder().encode(xml.replace(groupXml, updatedGroupXml));
  return readPptx(zipSync(archive));
}

function valueOf<T>(result: SourceTransformMatrixResult<T>): T {
  if (!result.ok) throw new Error(`unexpected matrix rejection: ${result.reason}`);
  return result.value;
}

function expectMatrixClose(actual: SourceAffineMatrix, expected: SourceAffineMatrix): void {
  for (const key of ["a", "b", "c", "d", "e", "f"] as const) {
    expect(actual[key]).toBeCloseTo(expected[key], 6);
  }
}

function expectXmlOrder(xml: string, markers: readonly string[]): void {
  const indexes = markers.map((marker) => xml.indexOf(marker));
  for (const [index, marker] of indexes.map((value, index) => [value, markers[index]] as const)) {
    expect(index, `missing XML marker ${marker}`).toBeGreaterThanOrEqual(0);
  }
  for (let index = 1; index < indexes.length; index += 1) {
    expect(indexes[index - 1]).toBeLessThan(requireValue(indexes[index]));
  }
}

function shapeIdentity(shape: SourceShapeNode): string {
  if (shape.nodeId !== undefined) return String(shape.nodeId);
  if (shape.kind === "raw") return `raw:${shape.raw.node.name.split(":").at(-1)}`;
  return `idless:${shape.kind}`;
}

function extractElementContaining(xml: string, qualifiedName: string, marker: string): string {
  const markerIndex = xml.indexOf(marker);
  if (markerIndex < 0) throw new Error(`XML marker '${marker}' was not found`);
  const start = xml.lastIndexOf(`<${qualifiedName}`, markerIndex);
  const closingTag = `</${qualifiedName}>`;
  const end = xml.indexOf(closingTag, markerIndex);
  if (start < 0 || end < 0) {
    throw new Error(`XML element '${qualifiedName}' containing '${marker}' was not found`);
  }
  return xml.slice(start, end + closingTag.length);
}

function relationshipIdentity(relationship: {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}) {
  return {
    id: relationship.id,
    type: relationship.type,
    target: relationship.target,
    targetMode: relationship.targetMode,
  };
}

function directChildren(source: PptxSourceModel, handle: SourceHandle): readonly SourceShapeNode[] {
  for (const root of [...source.slides, ...source.slideLayouts, ...source.slideMasters]) {
    if (root.handle?.partPath === handle.partPath && root.handle.nodeId === handle.nodeId) {
      return root.shapes;
    }
  }
  for (const root of [...source.slides, ...source.slideLayouts, ...source.slideMasters]) {
    const group = findGroupByHandle(root.shapes, handle);
    if (group !== undefined) return group.children;
  }
  throw new Error("test target was not found");
}

function findGroupByHandle(
  nodes: readonly SourceShapeNode[],
  handle: SourceHandle,
): SourceGroup | undefined {
  for (const node of nodes) {
    if (node.kind !== "group") continue;
    if (node.handle?.partPath === handle.partPath && node.handle.nodeId === handle.nodeId)
      return node;
    const nested = findGroupByHandle(node.children, handle);
    if (nested !== undefined) return nested;
  }
  return undefined;
}

function requireValue<T>(value: T | undefined): T {
  if (value === undefined) throw new Error("test fixture value is missing");
  return value;
}

function requireGroup(shape: SourceShapeNode | undefined): SourceGroup {
  if (shape?.kind !== "group") throw new Error("test fixture group is missing");
  return shape;
}

function requireTransform(shape: SourceShapeNode): NonNullable<SourceGroup["transform"]> {
  if (shape.kind === "raw" || shape.transform === undefined) {
    throw new Error("test fixture transform is missing");
  }
  return shape.transform;
}

function replaceFirstShape(source: PptxSourceModel, shape: SourceShapeNode): PptxSourceModel {
  const slide = requireValue(source.slides[0]);
  return { ...source, slides: [{ ...slide, shapes: [shape, ...slide.shapes.slice(1)] }] };
}

function asSmartArt(shape: SourceShapeNode): SourceShapeNode {
  return {
    kind: "smartArt",
    nodeId: shape.nodeId,
    handle: shape.handle,
    ...(shape.kind === "raw" ? {} : { transform: shape.transform }),
  };
}

function asRaw(shape: SourceShapeNode): SourceShapeNode {
  return {
    kind: "raw",
    nodeId: shape.nodeId,
    handle: shape.handle,
    raw: { id: asRawSidecarId("raw-target"), node: { name: "p:contentPart" } },
  };
}
