import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { createPptx } from "../builder/create-pptx.js";
import { createPptxAuthoringSession, reorderShapes } from "../index.js";
import { readPptx } from "../reader/read-pptx.js";
import { writePptx } from "../writer/write-pptx.js";
import { asPartPath, asRawSidecarId, asSourceNodeId, type SourceHandle } from "./handles.js";
import type { SourceGroup, SourceShapeNode } from "./shapes.js";
import { asEmu } from "./units.js";

describe("reorderShapes", () => {
  it("writes and reads a connector before its connection targets", () => {
    const source = createPptx();
    const slideHandle = requireValue(source.slides[0]?.handle);
    const session = createPptxAuthoringSession(source);
    const target = session.target(slideHandle);
    const first = addRect(target, 0);
    const second = addRect(target, 2000);
    const connector = target.addConnector({
      preset: "straightConnector1",
      offsetX: asEmu(1000),
      offsetY: asEmu(500),
      width: asEmu(1000),
      height: asEmu(1),
      start: { shapeHandle: first, connectionSiteIndex: 3 },
      end: { shapeHandle: second, connectionSiteIndex: 1 },
    });

    target.reorderShapes([connector, first, second]);

    expect(session.source.slides[0]?.shapes.map((shape) => shape.handle?.nodeId)).toEqual([
      connector.nodeId,
      first.nodeId,
      second.nodeId,
    ]);
    const reread = readPptx(writePptx(session.source));
    expect(reread.slides[0]?.shapes.map((shape) => shape.kind)).toEqual([
      "connector",
      "shape",
      "shape",
    ]);
    const rereadConnector = reread.slides[0]?.shapes[0];
    expect(rereadConnector?.kind).toBe("connector");
    if (rereadConnector?.kind !== "connector") throw new Error("connector was not reread");
    expect(rereadConnector.connection).toEqual({
      start: { shapeId: first.nodeId, connectionSiteIndex: 3 },
      end: { shapeId: second.nodeId, connectionSiteIndex: 1 },
    });
  });

  it("reorders slide, layout, and master targets", () => {
    const source = createPptx();
    const session = createPptxAuthoringSession(source);
    const targets = [
      requireValue(source.slides[0]?.handle),
      requireValue(source.slideLayouts[0]?.handle),
      requireValue(source.slideMasters[0]?.handle),
    ];

    for (const targetHandle of targets) {
      const target = session.target(targetHandle);
      const first = addRect(target, 0);
      const second = addRect(target, 2000);
      target.reorderShapes([second, first]);
    }

    const reread = readPptx(writePptx(session.source));
    const collections = [reread.slides, reread.slideLayouts, reread.slideMasters];
    for (const collection of collections) {
      expect(collection[0]?.shapes.map((shape) => Number(shape.nodeId))).toEqual([2, 1]);
    }
  });

  it("rejects missing, duplicate, foreign, and unknown shape handles", () => {
    const source = createPptx();
    const slideHandle = requireValue(source.slides[0]?.handle);
    const session = createPptxAuthoringSession(source);
    const target = session.target(slideHandle);
    const first = addRect(target, 0);
    const second = addRect(target, 2000);

    expect(() => target.reorderShapes([first])).toThrow("every target shape exactly once");
    expect(() => target.reorderShapes([first, first])).toThrow("duplicate shape");
    expect(() =>
      target.reorderShapes([first, { ...second, partPath: asPartPath("ppt/slides/other.xml") }]),
    ).toThrow("different drawing part");
    target.reorderShapes([{ ...second, orderingSlot: 999 }, first]);
    expect(session.source.slides[0]?.shapes.map((shape) => shape.handle?.nodeId)).toEqual([
      second.nodeId,
      first.nodeId,
    ]);
    expect(() =>
      target.reorderShapes([first, { ...second, nodeId: undefined, orderingSlot: 999 }]),
    ).toThrow("every shape handle requires a node id");
    expect(() =>
      target.reorderShapes([first, { ...second, nodeId: asSourceNodeId("999") }]),
    ).toThrow("was not found in the target drawing part");
  });

  it("preserves formatted text and extLst while the root function reorders drawings", () => {
    const authored = createTwoShapeSource();
    const archive = unzipSync(writePptx(authored));
    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = new TextDecoder().decode(requireValue(archive[slidePath]));
    archive[slidePath] = new TextEncoder().encode(
      slideXml.replace("</p:spTree>", '\n  <p:extLst><p:ext uri="test"/></p:extLst>\n</p:spTree>'),
    );
    const source = readPptx(zipSync(archive));
    const slide = requireValue(source.slides[0]);
    const handles = slide.shapes.map((shape) => requireValue(shape.handle));

    const reversedHandles = [...handles].reverse();
    const reordered = reorderShapes(source, requireValue(slide.handle), reversedHandles);
    const outputArchive = unzipSync(writePptx(reordered));
    const outputXml = new TextDecoder().decode(requireValue(outputArchive[slidePath]));

    expect(outputXml).toContain('<p:extLst><p:ext uri="test"/></p:extLst>');
    expect(outputXml).toContain("\n  \n");
    expect(readPptx(writePptx(reordered)).slides[0]?.shapes.map((shape) => shape.nodeId)).toEqual(
      reversedHandles.map((handle) => handle.nodeId),
    );
  });

  it("immutably reorders only a native group's direct children and preserves identity on write/read", () => {
    const source = createPptx();
    const session = createPptxAuthoringSession(source);
    const root = session.target(requireValue(source.slides[0]?.handle));
    const firstHandle = addRect(root, 0);
    const secondHandle = addRect(root, 2000);
    const siblingHandle = addRect(root, 4000);
    const groupHandle = root.groupShapes([firstHandle, secondHandle]);
    const beforeSlide = requireValue(session.source.slides[0]);
    const beforeGroup = requireGroup(beforeSlide.shapes[0]);
    const beforeSibling = requireValue(beforeSlide.shapes[1]);
    const beforeTransform = beforeGroup.transform;
    const beforeChildTransform = beforeGroup.childTransform;

    session.target(groupHandle).reorderShapes([secondHandle, firstHandle]);

    const afterSlide = requireValue(session.source.slides[0]);
    const afterGroup = requireGroup(afterSlide.shapes[0]);
    expect(afterSlide).not.toBe(beforeSlide);
    expect(afterGroup).not.toBe(beforeGroup);
    expect(afterSlide.shapes[1]).toBe(beforeSibling);
    expect(afterGroup.children).toEqual([beforeGroup.children[1], beforeGroup.children[0]]);
    expect(afterGroup.transform).toBe(beforeTransform);
    expect(afterGroup.childTransform).toBe(beforeChildTransform);
    expect(session.source.edits?.at(-1)).toEqual({
      kind: "reorderShapes",
      targetPartPath: groupHandle.partPath,
      parentGroupId: String(groupHandle.nodeId),
      shapeIds: [String(secondHandle.nodeId), String(firstHandle.nodeId)],
    });

    const rereadSlide = requireValue(readPptx(writePptx(session.source)).slides[0]);
    const rereadGroup = requireGroup(rereadSlide.shapes[0]);
    expect(rereadSlide.shapes[1]?.nodeId).toBe(siblingHandle.nodeId);
    expect(rereadGroup.nodeId).toBe(groupHandle.nodeId);
    expect(rereadGroup.children.map((child) => child.nodeId)).toEqual([
      secondHandle.nodeId,
      firstHandle.nodeId,
    ]);
    expect(rereadGroup.children.map((child) => child.handle?.nodeId)).toEqual([
      secondHandle.nodeId,
      firstHandle.nodeId,
    ]);
    expect(rereadGroup.transform).toEqual(beforeTransform);
    expect(rereadGroup.childTransform).toEqual(beforeChildTransform);
  });

  it("atomically rejects non-direct children, foreign parts, and AlternateContent for a group", () => {
    const source = createPptx();
    const session = createPptxAuthoringSession(source);
    const root = session.target(requireValue(source.slides[0]?.handle));
    const first = addRect(root, 0);
    const second = addRect(root, 2000);
    const third = addRect(root, 4000);
    const innerGroup = root.groupShapes([first, second]);
    const outerGroup = root.groupShapes([innerGroup, third]);
    const before = session.source;

    expect(() => session.target(outerGroup).reorderShapes([first, third])).toThrow(
      "every shape must be a direct child of the target",
    );
    expect(() =>
      session
        .target(outerGroup)
        .reorderShapes([innerGroup, { ...third, partPath: asPartPath("ppt/slides/other.xml") }]),
    ).toThrow("different drawing part");
    expect(session.source).toBe(before);

    const slide = requireValue(before.slides[0]);
    const alternateContent = {
      ...before,
      slides: [
        {
          ...slide,
          shapes: slide.shapes.map((shape) =>
            shape.kind === "group"
              ? {
                  ...shape,
                  rawSidecars: [
                    {
                      id: asRawSidecarId("alternate-content"),
                      node: { name: "mc:AlternateContent" },
                    },
                  ],
                }
              : shape,
          ),
        },
      ],
    } satisfies typeof before;
    expect(() => reorderShapes(alternateContent, outerGroup, [third, innerGroup])).toThrow(
      "mc:AlternateContent shape trees are not supported",
    );
    expect(alternateContent.edits).toBe(before.edits);
  });

  it("rejects duplicate ids at the root and across group descendants without changing source", () => {
    const authored = createTwoShapeSource();
    const slide = requireValue(authored.slides[0]);
    const first = requireValue(slide.shapes[0]);
    const second = requireValue(slide.shapes[1]);
    const firstHandle = requireValue(first.handle);
    const secondHandle = requireValue(second.handle);
    const duplicateRoot = {
      ...authored,
      slides: [
        {
          ...slide,
          shapes: [
            first,
            {
              ...second,
              nodeId: first.nodeId,
              handle: { ...secondHandle, nodeId: first.nodeId },
            },
          ],
        },
      ],
    } satisfies typeof authored;
    const duplicateDescendant = {
      ...authored,
      slides: [
        {
          ...slide,
          shapes: [
            {
              kind: "group",
              nodeId: first.nodeId,
              children: [first],
              handle: { ...firstHandle, orderingSlot: 0 },
            },
            second,
          ],
        },
      ],
    } satisfies typeof authored;

    expect(() =>
      reorderShapes(duplicateRoot, requireValue(slide.handle), [firstHandle, secondHandle]),
    ).toThrow("duplicate node id");
    expect(() =>
      reorderShapes(duplicateDescendant, requireValue(slide.handle), [firstHandle, secondHandle]),
    ).toThrow("duplicate node id");
    expect(duplicateRoot.edits).toBe(authored.edits);
    expect(duplicateDescendant.edits).toBe(authored.edits);
  });

  it("rejects an AlternateContent shape tree before creating an edit", () => {
    const authored = createTwoShapeSource();
    const archive = unzipSync(writePptx(authored));
    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = new TextDecoder().decode(requireValue(archive[slidePath]));
    const firstShape = requireValue(slideXml.match(/<p:sp>[\s\S]*?<\/p:sp>/)?.[0]);
    archive[slidePath] = new TextEncoder().encode(
      slideXml
        .replace(
          "<p:sld ",
          '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ',
        )
        .replace(
          firstShape,
          `<mc:AlternateContent><mc:Choice Requires="p14">${firstShape}</mc:Choice></mc:AlternateContent>`,
        ),
    );
    const source = readPptx(zipSync(archive));
    const slide = requireValue(source.slides[0]);
    const handles = slide.shapes.map((shape) => requireValue(shape.handle));

    expect(() => reorderShapes(source, requireValue(slide.handle), [...handles].reverse())).toThrow(
      "mc:AlternateContent shape trees are not supported",
    );
    expect(source.edits).toBeUndefined();
  });
});

type Target = ReturnType<ReturnType<typeof createPptxAuthoringSession>["target"]>;

function addRect(target: Target, offsetX: number): SourceHandle {
  return target.addShape({
    geometry: { kind: "preset", preset: "rect" },
    offsetX: asEmu(offsetX),
    offsetY: asEmu(0),
    width: asEmu(1000),
    height: asEmu(1000),
  });
}

function requireValue<T>(value: T | undefined): T {
  if (value === undefined) throw new Error("test fixture value is missing");
  return value;
}

function requireGroup(shape: SourceShapeNode | undefined): SourceGroup {
  if (shape?.kind !== "group") throw new Error("test fixture group is missing");
  return shape;
}

function createTwoShapeSource() {
  const source = createPptx();
  const session = createPptxAuthoringSession(source);
  const target = session.target(requireValue(source.slides[0]?.handle));
  addRect(target, 0);
  addRect(target, 2000);
  return session.source;
}
