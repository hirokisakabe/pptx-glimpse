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

  it("keeps chart edit ordering, selection, and history atomic across a move", () => {
    const source = chartTwoSlideSource();
    const chart = source.slides[0]?.shapes[0];
    if (chart?.kind !== "chart" || chart.handle === undefined) {
      throw new Error("editor chart move fixture is missing");
    }
    const destinationHandle = requireValue(source.slides[1]?.handle);
    const rejected = createEditorSession(source);
    expect(rejected.selectShape(chart.handle)).toMatchObject({ ok: true });
    expect(
      rejected.applyAll([
        {
          kind: "moveShapesAcrossSlides",
          shapeHandles: [chart.handle],
          destinationSlideHandle: destinationHandle,
        },
        {
          kind: "updateChartData",
          handle: chart.handle,
          series: [{ name: "Rejected", categories: ["A"], values: [9] }],
        },
      ]),
    ).toMatchObject({ ok: false, code: "invalid-command" });
    expect(rejected.document).toBe(source);
    expect(rejected.selection).toEqual({ shapeHandle: chart.handle });
    expect(rejected.undoDepth).toBe(0);

    const session = createEditorSession(source);
    expect(session.selectShape(chart.handle)).toMatchObject({ ok: true });
    expect(
      session.applyAll([
        {
          kind: "updateChartData",
          handle: chart.handle,
          series: [{ name: "Before move", categories: ["A", "B"], values: [2, 3] }],
        },
        {
          kind: "moveShapesAcrossSlides",
          shapeHandles: [chart.handle],
          destinationSlideHandle: destinationHandle,
        },
      ]),
    ).toMatchObject({ ok: true });
    const movedHandle = requireValue(session.selection).shapeHandle;
    expect(movedHandle.partPath).toBe(source.slides[1]?.partPath);
    expect(session.document.edits?.map((edit) => edit.kind)).toEqual([
      "updateChartData",
      "moveShapesAcrossSlides",
    ]);
    expectPersistedChartText(session.document, 1, "Before move");
    expect(session.undoDepth).toBe(1);
    expect(
      session.apply({
        kind: "updateChartData",
        handle: movedHandle,
        series: [{ name: "After move", categories: ["A", "B"], values: [4, 5] }],
      }),
    ).toMatchObject({ ok: true });
    expectPersistedChartText(session.document, 1, "After move");
    expect(session.selection).toEqual({ shapeHandle: movedHandle });
    expect(session.undoDepth).toBe(2);
    expect(session.undo()).toMatchObject({ ok: true });
    expect(session.selection).toEqual({ shapeHandle: movedHandle });
    expectPersistedChartText(session.document, 1, "Before move");
    expect(session.undo()).toMatchObject({ ok: true });
    expect(session.selection).toEqual({ shapeHandle: chart.handle });
    expect(session.document).toBe(source);
    expect(session.redo()).toMatchObject({ ok: true });
    expect(session.selection).toEqual({ shapeHandle: movedHandle });
    expectPersistedChartText(session.document, 1, "Before move");
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

function chartTwoSlideSource() {
  return cleanTwoSlideSource((session, first) => {
    session.target(first).addChart({
      chartType: "bar",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(2_400_000),
      height: asEmu(1_800_000),
      name: "Moved chart",
      series: [{ name: "Initial", categories: ["A"], values: [1] }],
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

function expectPersistedChartText(
  source: ReturnType<typeof readPptx>,
  slideIndex: number,
  expected: string,
): void {
  const persisted = readPptx(writePptx(source));
  const slide = requireValue(persisted.slides[slideIndex]);
  const chart = slide.shapes.find((shape) => shape.kind === "chart");
  const relationshipId = requireValue(
    chart?.kind === "chart" ? chart.chartRelationshipId : undefined,
  );
  const relationship = requireValue(
    persisted.packageGraph.relationships
      .find((group) => group.sourcePartPath === slide.partPath)
      ?.relationships.find((candidate) => candidate.id === relationshipId),
  );
  const chartPartPath = resolvePackageTarget(slide.partPath, relationship.target);
  const chartPart = requireValue(
    persisted.packageGraph.rawParts?.find((part) => part.partPath === chartPartPath),
  );
  if (chartPart.kind !== "binary") throw new Error("persisted chart part is not binary");
  expect(new TextDecoder().decode(chartPart.bytes)).toContain(expected);
}

function resolvePackageTarget(sourcePartPath: string, target: string): string {
  const baseDirectory = sourcePartPath.slice(0, Math.max(0, sourcePartPath.lastIndexOf("/")));
  const combined = target.startsWith("/") ? target.slice(1) : `${baseDirectory}/${target}`;
  const segments: string[] = [];
  for (const segment of combined.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") segments.pop();
    else segments.push(segment);
  }
  return segments.join("/");
}
