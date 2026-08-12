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
    const beforeGraph = source.packageGraph;
    const slide = requireValue(source.slides[0]);
    const connector = requireValue(slide.shapes.find((shape) => shape.kind === "connector"));
    const picture = requireValue(slide.shapes.find((shape) => shape.kind === "image"));
    const chart = requireValue(slide.shapes.find((shape) => shape.kind === "chart"));
    const edited = moveShapes(
      source,
      [requireValue(chart.handle), requireValue(picture.handle)],
      slideHandle,
      { beforeShapeHandle: first },
    );

    expect(edited.packageGraph).toBe(beforeGraph);
    const reread = readPptx(writePptx(edited));
    const rereadConnector = requireValue(
      reread.slides[0]?.shapes.find((shape) => shape.kind === "connector"),
    );
    const rereadPicture = requireValue(
      reread.slides[0]?.shapes.find((shape) => shape.kind === "image"),
    );
    const rereadChart = requireValue(
      reread.slides[0]?.shapes.find((shape) => shape.kind === "chart"),
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

    expect(outputXml.indexOf(`id="${third.nodeId}"`)).toBeLessThan(outputXml.indexOf('uri="slot"'));
    expect(outputXml.indexOf('uri="slot"')).toBeLessThan(outputXml.indexOf(`id="${first.nodeId}"`));
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

    expect(() => moveShapes(grouped, [handles[0]], requireValue(slide.handle))).toThrow(
      "direct child",
    );
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

function directChildren(source: PptxSourceModel, handle: SourceHandle): readonly SourceShapeNode[] {
  for (const root of [...source.slides, ...source.slideLayouts, ...source.slideMasters]) {
    if (root.handle?.partPath === handle.partPath && root.handle.nodeId === handle.nodeId) {
      return root.shapes;
    }
  }
  throw new Error("test target was not found");
}

function requireValue<T>(value: T | undefined): T {
  if (value === undefined) throw new Error("test fixture value is missing");
  return value;
}

function requireGroup(shape: SourceShapeNode | undefined): SourceGroup {
  if (shape?.kind !== "group") throw new Error("test fixture group is missing");
  return shape;
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
