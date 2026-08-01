import { describe, expect, it } from "vitest";

import { createPptx } from "../builder/create-pptx.js";
import { createPptxAuthoringSession, readPptx, writePptx } from "../index.js";
import { asPartPath } from "./handles.js";
import { asEmu } from "./units.js";

const PNG_BYTES = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

describe("PptxAuthoringSession", () => {
  it("applies drawing primitives consecutively to one target and returns their handles", () => {
    const original = createPptx();
    const slideHandle = requireHandle(original.slides[0]?.handle);
    const session = createPptxAuthoringSession(original);
    const target = session.target(slideHandle);

    const handles = [
      target.addTextBox({
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        text: "Text",
      }),
      target.addShape({
        geometry: { kind: "preset", preset: "rect" },
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
      }),
      target.addConnector({
        preset: "straightConnector1",
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1),
      }),
      target.addPicture({
        bytes: PNG_BYTES,
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
      }),
      target.addTable({
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        columnWidths: [asEmu(1000)],
        rows: [{ height: asEmu(1000), cells: [{ text: "Cell" }] }],
      }),
      target.addChart({
        chartType: "bar",
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        series: [{ categories: ["A"], values: [1] }],
      }),
    ];

    const authoredSlide = session.source.slides[0];
    expect(original.slides[0].shapes).toEqual([]);
    expect(handles.map((handle) => handle.partPath)).toEqual(
      Array.from({ length: 6 }, () => authoredSlide.partPath),
    );
    expect(handles.map((handle) => handle.nodeId)).toEqual(["1", "2", "3", "4", "5", "6"]);
    expect(handles.map((handle) => handle.orderingSlot)).toEqual([0, 1, 2, 3, 4, 5]);
    expect(handles).toEqual(authoredSlide.shapes.map((shape) => shape.handle));
    expect(session.source.edits?.map((edit) => edit.kind)).toEqual([
      "addTextBox",
      "addShape",
      "addConnector",
      "addPicture",
      "addTable",
      "addChart",
    ]);
  });

  it("switches slide, layout, and master targets without part-local id or ordering collisions", () => {
    const source = createPptx({
      slideMaster: {
        background: { kind: "solid", color: { kind: "srgb", hex: "AABBCC" } },
      },
    });
    const firstSlide = requireHandle(source.slides[0]?.handle);
    const layout = source.slideLayouts[0];
    const layoutHandle = requireHandle(layout?.handle);
    const master = source.slideMasters[0];
    const masterHandle = requireHandle(master?.handle);
    const session = createPptxAuthoringSession(source);
    const secondSlide = session.addEmptySlideFromLayout({ layoutPartPath: layout.partPath });

    const targets = [firstSlide, secondSlide, layoutHandle, masterHandle];
    const handlesByPart = new Map<string, ReturnType<typeof addRect>[] | undefined>();
    for (const targetHandle of [...targets, ...targets]) {
      const handle = addRect(session.target(targetHandle));
      handlesByPart.set(handle.partPath, [...(handlesByPart.get(handle.partPath) ?? []), handle]);
    }

    for (const targetHandle of targets) {
      const handles = handlesByPart.get(targetHandle.partPath);
      const ids = handles?.map((handle) => Number(handle.nodeId));
      expect(new Set(ids).size).toBe(2);
      expect(ids?.[1]).toBe((ids?.[0] ?? 0) + 1);
      expect(handles?.map((handle) => handle.orderingSlot)).toEqual([0, 1]);
      expect(handles?.every((handle) => handle.partPath === targetHandle.partPath)).toBe(true);
    }

    const layoutSlideNumber = addSlideNumber(session.target(layoutHandle));
    const masterSlideNumber = addSlideNumber(session.target(masterHandle));
    expect(layoutSlideNumber.partPath).toBe(layoutHandle.partPath);
    expect(layoutSlideNumber.orderingSlot).toBe(2);
    expect(Number(layoutSlideNumber.nodeId)).toBe(
      Number(handlesByPart.get(layoutHandle.partPath)?.at(-1)?.nodeId) + 1,
    );
    expect(masterSlideNumber.partPath).toBe(masterHandle.partPath);
    expect(masterSlideNumber.orderingSlot).toBe(2);
    expect(Number(masterSlideNumber.nodeId)).toBe(
      Number(handlesByPart.get(masterHandle.partPath)?.at(-1)?.nodeId) + 1,
    );
    expect(secondSlide).toEqual(session.source.slides[1].handle);

    session.target(secondSlide).setSlideBackground({
      kind: "solid",
      color: { kind: "srgb", hex: "112233" },
    });
    expect(session.source.slides[1].background).toEqual({
      kind: "fill",
      fill: { kind: "solid", color: { kind: "srgb", hex: "112233" } },
    });
    session.target(layoutHandle).setBackground({
      kind: "solid",
      color: { kind: "srgb", hex: "445566" },
    });
    expect(session.source.slideLayouts[0]?.background).toMatchObject({
      kind: "fill",
      fill: { kind: "solid", color: { kind: "srgb", hex: "445566" } },
    });
    session.target(masterHandle).clearBackground();
    expect(session.source.slideMasters[0]?.background).toBeUndefined();
    expect(session.source.edits?.at(-1)).toEqual({
      kind: "setBackground",
      targetPartPath: masterHandle.partPath,
    });
  });

  it("groups authored drawing kinds and supports the same target abstraction for native groups", () => {
    const source = createPptx();
    const session = createPptxAuthoringSession(source);
    const target = session.target(requireHandle(source.slides[0]?.handle));
    const handles = [
      addRect(target),
      target.addPicture({
        bytes: PNG_BYTES,
        offsetX: asEmu(1200),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
      }),
      target.addConnector({
        preset: "straightConnector1",
        offsetX: asEmu(2300),
        offsetY: asEmu(500),
        width: asEmu(1000),
        height: asEmu(1),
      }),
      target.addTable({
        offsetX: asEmu(3400),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        columnWidths: [asEmu(1000)],
        rows: [{ height: asEmu(1000), cells: [{ text: "Cell" }] }],
      }),
      target.addChart({
        chartType: "bar",
        offsetX: asEmu(4500),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        series: [{ categories: ["A"], values: [1] }],
      }),
    ];

    const outerGroupHandle = target.groupShapes(handles);
    const outerGroup = session.source.slides[0]?.shapes[0];
    expect(outerGroup?.kind).toBe("group");
    if (outerGroup?.kind !== "group") throw new Error("authored group is missing");
    expect(outerGroup.handle).toEqual(outerGroupHandle);
    expect(outerGroup.children.map((child) => child.kind)).toEqual([
      "shape",
      "image",
      "connector",
      "table",
      "chart",
    ]);

    const rereadAuthored = readPptx(writePptx(session.source));
    const nativeOuter = rereadAuthored.slides[0]?.shapes[0];
    if (nativeOuter?.kind !== "group" || nativeOuter.handle === undefined) {
      throw new Error("reread authored group is missing");
    }
    const nativeSession = createPptxAuthoringSession(rereadAuthored);
    const innerGroupHandle = nativeSession
      .target(identityHandle(nativeOuter.handle))
      .groupShapes(nativeOuter.children.slice(0, 2).map((child) => identityHandle(child.handle)));
    const reread = readPptx(writePptx(nativeSession.source));
    const rereadOuter = reread.slides[0]?.shapes[0];
    expect(rereadOuter?.kind).toBe("group");
    if (rereadOuter?.kind !== "group") throw new Error("reread outer group is missing");
    expect(rereadOuter.children[0]?.kind).toBe("group");
    expect(rereadOuter.children[0]?.nodeId).toBe(innerGroupHandle.nodeId);
    expect(rereadOuter.children.slice(1).map((child) => child.kind)).toEqual([
      "connector",
      "table",
      "chart",
    ]);
  });

  it("preserves primitive errors and the last successful source when a target is invalid", () => {
    const source = createPptx();
    const session = createPptxAuthoringSession(source);
    const missingTarget = session.target({ partPath: asPartPath("ppt/slides/missing.xml") });

    expect(() => addRect(missingTarget)).toThrow(
      "addShape: slide, layout, or master handle was not found in PptxSourceModel source",
    );
    expect(session.source).toBe(source);

    expect(() =>
      session.target(requireHandle(source.slideLayouts[0]?.handle)).addTable({
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        columnWidths: [asEmu(1000)],
        rows: [{ height: asEmu(1000), cells: [{ text: "Cell" }] }],
      }),
    ).toThrow("addTable: slide handle was not found in PptxSourceModel source");
    expect(() =>
      session.target(requireHandle(source.slideMasters[0]?.handle)).setSlideBackground({
        kind: "solid",
        color: { kind: "srgb", hex: "112233" },
      }),
    ).toThrow("setSlideBackground: slide handle was not found in PptxSourceModel source");
    expect(session.source).toBe(source);
  });
});

type Target = ReturnType<ReturnType<typeof createPptxAuthoringSession>["target"]>;

function addRect(target: Target) {
  return target.addShape({
    geometry: { kind: "preset", preset: "rect" },
    offsetX: asEmu(0),
    offsetY: asEmu(0),
    width: asEmu(1000),
    height: asEmu(1000),
  });
}

function addSlideNumber(target: Target) {
  return target.addSlideNumber({
    offsetX: asEmu(0),
    offsetY: asEmu(0),
    width: asEmu(1000),
    height: asEmu(1000),
  });
}

function requireHandle<T>(handle: T | undefined): T {
  if (handle === undefined) throw new Error("test fixture handle is missing");
  return handle;
}

function identityHandle(handle: import("./handles.js").SourceHandle | undefined) {
  const resolved = requireHandle(handle);
  return {
    partPath: resolved.partPath,
    ...(resolved.nodeId !== undefined ? { nodeId: resolved.nodeId } : {}),
  };
}
