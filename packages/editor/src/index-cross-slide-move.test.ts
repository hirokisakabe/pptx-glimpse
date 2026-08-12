import {
  asEmu,
  createPptx,
  createPptxAuthoringSession,
  readPptx,
  type SourceHandle,
  writePptx,
} from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createEditorSession } from "./index.js";

describe("EditorSession cross-slide drawing move", () => {
  it("maps selection through apply, undo, and redo", () => {
    const source = twoSlideSource();
    const sourceSlide = requireValue(source.slides[0]);
    const destination = requireValue(source.slides[1]);
    const firstHandle = requireValue(sourceSlide.shapes[0]?.handle);
    const handles = sourceSlide.shapes.map((shape) => requireValue(shape.handle));
    const session = createEditorSession(source);
    expect(session.selectShape(firstHandle)).toMatchObject({ ok: true });

    const result = session.apply({
      kind: "moveShapesAcrossSlides",
      shapeHandles: handles,
      destinationSlideHandle: requireValue(destination.handle),
    });

    expect(result).toMatchObject({ ok: true });
    const movedSelection = requireValue(session.selection).shapeHandle;
    expect(movedSelection.partPath).toBe(destination.partPath);
    expect(movedSelection.nodeId).not.toBe(firstHandle.nodeId);
    expect(session.undoDepth).toBe(1);

    expect(session.undo()).toMatchObject({ ok: true });
    expect(session.selection).toEqual({ shapeHandle: firstHandle });
    expect(session.document).toBe(source);
    expect(session.redo()).toMatchObject({ ok: true });
    expect(session.selection).toEqual({ shapeHandle: movedSelection });
  });

  it("keeps document, selection, and history unchanged on boundary rejection", () => {
    const source = connectedTwoSlideSource();
    const sourceSlide = requireValue(source.slides[0]);
    const session = createEditorSession(source);
    const firstHandle = requireValue(sourceSlide.shapes[0]?.handle);
    expect(session.selectShape(firstHandle)).toMatchObject({ ok: true });

    const result = session.apply({
      kind: "moveShapesAcrossSlides",
      shapeHandles: [firstHandle],
      destinationSlideHandle: requireValue(source.slides[1]?.handle),
    });

    expect(result).toMatchObject({ ok: false, code: "invalid-command" });
    expect(session.document).toBe(source);
    expect(session.selection).toEqual({ shapeHandle: firstHandle });
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });

  it("requires the cross-slide command to be last in applyAll", () => {
    const source = twoSlideSource();
    const handle = requireValue(source.slides[0]?.shapes[0]?.handle);
    const session = createEditorSession(source);
    const result = session.applyAll([
      {
        kind: "moveShapesAcrossSlides",
        shapeHandles: [handle],
        destinationSlideHandle: requireValue(source.slides[1]?.handle),
      },
      { kind: "moveShape", handle, offsetX: asEmu(1), offsetY: asEmu(1) },
    ]);
    expect(result).toMatchObject({ ok: false, code: "invalid-command" });
    expect(session.document).toBe(source);
    expect(session.undoDepth).toBe(0);
  });
});

function twoSlideSource() {
  return cleanTwoSlideSource((session, first) => {
    session.target(first).addShape({ ...rect(0), name: "First" });
    session.target(first).addShape({ ...rect(1200), name: "Second" });
  });
}

function connectedTwoSlideSource() {
  return cleanTwoSlideSource((session, first) => {
    const target = session.target(first);
    const start = target.addShape({ ...rect(0), name: "Start" });
    const end = target.addShape({ ...rect(2400), name: "End" });
    target.addConnector({
      preset: "straightConnector1",
      offsetX: asEmu(1000),
      offsetY: asEmu(500),
      width: asEmu(1400),
      height: asEmu(1),
      start: { shapeHandle: start, connectionSiteIndex: 1 },
      end: { shapeHandle: end, connectionSiteIndex: 3 },
    });
  });
}

type AuthoringSession = ReturnType<typeof createPptxAuthoringSession>;

function cleanTwoSlideSource(author: (session: AuthoringSession, first: SourceHandle) => void) {
  const initial = createPptx();
  const session = createPptxAuthoringSession(initial);
  const first = requireValue(initial.slides[0]?.handle);
  const layout = requireValue(initial.slideLayouts[0]?.handle);
  session.addEmptySlideFromLayout({ layoutPartPath: layout.partPath });
  author(session, first);
  return readPptx(writePptx(session.source));
}

function rect(offsetX: number) {
  return {
    geometry: { kind: "preset" as const, preset: "rect" },
    offsetX: asEmu(offsetX),
    offsetY: asEmu(0),
    width: asEmu(1000),
    height: asEmu(1000),
  };
}

function requireValue<T>(value: T | undefined): T {
  if (value === undefined) throw new Error("fixture value is missing");
  return value;
}
