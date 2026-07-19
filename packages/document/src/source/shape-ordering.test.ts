import { describe, expect, it } from "vitest";

import { createPptx } from "../builder/create-pptx.js";
import { readPptx } from "../reader/read-pptx.js";
import { writePptx } from "../writer/write-pptx.js";
import { createPptxAuthoringSession } from "./authoring-session.js";
import { asPartPath, type SourceHandle } from "./handles.js";
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
    expect(() => target.reorderShapes([first, { ...second, orderingSlot: 999 }])).toThrow(
      "was not found in the target drawing part",
    );
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
