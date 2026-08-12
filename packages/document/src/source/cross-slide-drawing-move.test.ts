import { describe, expect, it } from "vitest";

import {
  asEmu,
  asRawSidecarId,
  createPptx,
  createPptxAuthoringSession,
  moveShapesAcrossSlides,
  type PptxSourceModel,
  readPptx,
  type SourceHandle,
  type SourceTextBody,
  updateChartData,
  writePptx,
} from "../index.js";
import { BLUE_PNG } from "../writer/write-pptx.test-helpers.js";
import { resolveInternalRelationshipTarget } from "./package-paths.js";

describe("moveShapesAcrossSlides", () => {
  it("moves a consecutive shape/picture block before an anchor and persists remapped ids", () => {
    const source = buildCleanTwoSlideSource((session, first, second) => {
      const sourceTarget = session.target(first);
      sourceTarget.addShape({ ...rect(0), name: "Moved shape", text: "Moved text" });
      sourceTarget.addPicture({
        bytes: BLUE_PNG,
        offsetX: asEmu(1200),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        name: "Moved picture",
      });
      session.target(second).addShape({ ...rect(2400), name: "Destination anchor" });
    });
    const sourceSlide = requireValue(source.slides[0]);
    const destinationSlide = requireValue(source.slides[1]);
    const handles = sourceSlide.shapes.map((shape) => requireValue(shape.handle));
    const before = structuredClone(source);

    const result = moveShapesAcrossSlides(
      source,
      [...handles].reverse(),
      requireValue(destinationSlide.handle),
      { beforeShapeHandle: requireValue(destinationSlide.shapes[0]?.handle) },
    );

    expect(source).toEqual(before);
    expect(result.document.slides[0]?.shapes).toHaveLength(0);
    expect(result.document.slides[1]?.shapes.map(shapeName)).toEqual([
      "Moved shape",
      "Moved picture",
      "Destination anchor",
    ]);
    const rootMappings = result.moved.filter((mapping) =>
      handles.some(
        (handle) =>
          handle.partPath === mapping.before.partPath && handle.nodeId === mapping.before.nodeId,
      ),
    );
    expect(rootMappings).toHaveLength(2);
    expect(rootMappings.map((mapping) => mapping.after.orderingSlot)).toEqual([0, 1]);
    expect(
      rootMappings.every((mapping) => mapping.after.partPath === destinationSlide.partPath),
    ).toBe(true);
    expect(rootMappings.every((mapping) => mapping.before.nodeId !== mapping.after.nodeId)).toBe(
      true,
    );
    const movedShape = result.document.slides[1]?.shapes.find(
      (shape) => shape.kind === "shape" && shape.name === "Moved shape",
    );
    if (movedShape?.kind !== "shape" || movedShape.textBody === undefined) {
      throw new Error("moved text shape is missing");
    }
    expect(
      textHandles(movedShape.textBody).every(
        (handle) => handle.partPath === destinationSlide.partPath,
      ),
    ).toBe(true);

    const sourceRelationships = requireValue(
      result.document.packageGraph.relationships.find(
        (group) => group.sourcePartPath === sourceSlide.partPath,
      ),
    );
    const destinationRelationships = requireValue(
      result.document.packageGraph.relationships.find(
        (group) => group.sourcePartPath === destinationSlide.partPath,
      ),
    );
    expect(
      sourceRelationships.relationships.some((relationship) =>
        relationship.type.endsWith("/image"),
      ),
    ).toBe(false);
    expect(
      destinationRelationships.relationships.some((relationship) =>
        relationship.type.endsWith("/image"),
      ),
    ).toBe(true);
    expect(result.document.packageGraph.media).toEqual(source.packageGraph.media);

    const persisted = readPptx(writePptx(result.document));
    expect(persisted.slides[0]?.shapes).toHaveLength(0);
    expect(persisted.slides[1]?.shapes.map(shapeName)).toEqual([
      "Moved shape",
      "Moved picture",
      "Destination anchor",
    ]);
    const persistedShape = persisted.slides[1]?.shapes.find(
      (shape) => shape.kind === "shape" && shape.name === "Moved shape",
    );
    if (persistedShape?.kind !== "shape" || persistedShape.textBody === undefined) {
      throw new Error("persisted text shape is missing");
    }
    expect(
      textHandles(persistedShape.textBody).every(
        (handle) => handle.partPath === destinationSlide.partPath,
      ),
    ).toBe(true);
    const persistedPicture = persisted.slides[1]?.shapes.find((shape) => shape.kind === "image");
    expect(persistedPicture?.blipRelationshipId).toBe(
      result.document.slides[1]?.shapes.find((shape) => shape.kind === "image")?.blipRelationshipId,
    );
  });

  it("remaps an internal connector closure and rejects a selection-crossing boundary atomically", () => {
    const source = buildCleanTwoSlideSource((session, first) => {
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
        name: "Connector",
      });
    });
    const sourceSlide = requireValue(source.slides[0]);
    const destinationSlide = requireValue(source.slides[1]);
    const handles = sourceSlide.shapes.map((shape) => requireValue(shape.handle));
    const before = structuredClone(source);

    expect(() =>
      moveShapesAcrossSlides(source, [handles[0]], requireValue(destinationSlide.handle)),
    ).toThrow("connector references must stay inside the moved closure");
    expect(source).toEqual(before);

    const result = moveShapesAcrossSlides(source, handles, requireValue(destinationSlide.handle));
    const connector = result.document.slides[1]?.shapes.find((shape) => shape.kind === "connector");
    const movedIds = new Set(result.document.slides[1]?.shapes.map((shape) => shape.nodeId));
    expect(connector?.kind).toBe("connector");
    if (connector?.kind !== "connector") throw new Error("connector result is missing");
    expect(movedIds.has(connector.connection?.start?.shapeId)).toBe(true);
    expect(movedIds.has(connector.connection?.end?.shapeId)).toBe(true);

    const persisted = readPptx(writePptx(result.document));
    const persistedConnector = persisted.slides[1]?.shapes.find(
      (shape) => shape.kind === "connector",
    );
    expect(persistedConnector).toMatchObject({
      kind: "connector",
      connection: {
        start: { shapeId: connector.connection?.start?.shapeId },
        end: { shapeId: connector.connection?.end?.shapeId },
      },
    });
  });

  it("retains a source relationship while a non-moved picture still references it", () => {
    const source = buildCleanTwoSlideSource((session, first) => {
      const target = session.target(first);
      for (const name of ["Move shared", "Keep shared"]) {
        target.addPicture({
          bytes: BLUE_PNG,
          offsetX: asEmu(name.startsWith("Move") ? 0 : 1200),
          offsetY: asEmu(0),
          width: asEmu(1000),
          height: asEmu(1000),
          name,
        });
      }
    });
    const sourceSlide = requireValue(source.slides[0]);
    const [move, keep] = sourceSlide.shapes;
    if (
      move?.kind !== "image" ||
      keep?.kind !== "image" ||
      move.blipRelationshipId === undefined ||
      keep.blipRelationshipId === undefined ||
      keep.handle === undefined
    ) {
      throw new Error("shared relationship fixture pictures are missing");
    }
    const sharedSource = sharePictureRelationship(
      source,
      sourceSlide.partPath,
      keep.nodeId,
      keep.blipRelationshipId,
      move.blipRelationshipId,
    );

    const result = moveShapesAcrossSlides(
      sharedSource,
      [requireValue(move.handle)],
      requireValue(sharedSource.slides[1]?.handle),
    );
    const sourceRelationships = requireValue(
      result.document.packageGraph.relationships.find(
        (group) => group.sourcePartPath === sourceSlide.partPath,
      ),
    );
    expect(sourceRelationships.relationships.map((relationship) => relationship.id)).toContain(
      move.blipRelationshipId,
    );
    const persisted = readPptx(writePptx(result.document));
    const kept = persisted.slides[0]?.shapes.find((shape) => shape.kind === "image");
    expect(kept?.kind === "image" ? kept.blipRelationshipId : undefined).toBe(
      move.blipRelationshipId,
    );
  });

  it("retains a moved picture relationship while the source background still references it", () => {
    const source = buildCleanTwoSlideSource((session, first) => {
      session.target(first).addPicture({
        bytes: BLUE_PNG,
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        name: "Move shared background",
      });
      session.target(first).setSlideBackground({ kind: "image", bytes: BLUE_PNG });
    });
    const sourceSlide = requireValue(source.slides[0]);
    const picture = sourceSlide.shapes[0];
    const background = sourceSlide.background;
    if (
      picture?.kind !== "image" ||
      picture.handle === undefined ||
      picture.blipRelationshipId === undefined ||
      background?.kind !== "fill" ||
      background.fill.kind !== "image" ||
      background.fill.blipRelationshipId === undefined
    ) {
      throw new Error("shared background relationship fixture is missing");
    }
    const oldBackgroundRelationshipId = background.fill.blipRelationshipId;
    const sharedRelationshipId = picture.blipRelationshipId;
    const sharedSource: PptxSourceModel = {
      ...source,
      packageGraph: {
        ...source.packageGraph,
        rawParts: source.packageGraph.rawParts?.map((part) =>
          part.partPath === sourceSlide.partPath && part.kind === "binary"
            ? {
                ...part,
                bytes: new TextEncoder().encode(
                  new TextDecoder()
                    .decode(part.bytes)
                    .replaceAll(oldBackgroundRelationshipId, sharedRelationshipId),
                ),
              }
            : part,
        ),
      },
      slides: source.slides.map((slide) =>
        slide.partPath === sourceSlide.partPath
          ? {
              ...slide,
              background: {
                kind: "fill" as const,
                fill: { kind: "image" as const, blipRelationshipId: sharedRelationshipId },
              },
            }
          : slide,
      ),
    };

    const result = moveShapesAcrossSlides(
      sharedSource,
      [picture.handle],
      requireValue(sharedSource.slides[1]?.handle),
    );
    const sourceRelationships = requireValue(
      result.document.packageGraph.relationships.find(
        (group) => group.sourcePartPath === sourceSlide.partPath,
      ),
    );
    expect(sourceRelationships.relationships.map((relationship) => relationship.id)).toContain(
      sharedRelationshipId,
    );

    const persisted = readPptx(writePptx(result.document));
    const persistedBackground = persisted.slides[0]?.background;
    expect(
      persistedBackground?.kind === "fill" && persistedBackground.fill.kind === "image"
        ? persistedBackground.fill.blipRelationshipId
        : undefined,
    ).toBe(sharedRelationshipId);

    const rawFillSource: PptxSourceModel = {
      ...sharedSource,
      slides: sharedSource.slides.map((slide) =>
        slide.partPath === sourceSlide.partPath
          ? {
              ...slide,
              background: {
                kind: "fill" as const,
                fill: {
                  kind: "raw" as const,
                  raw: {
                    id: asRawSidecarId("background-fill"),
                    node: {
                      name: "a:blipFill",
                      children: [
                        {
                          name: "a:blip",
                          attributes: { "r:embed": sharedRelationshipId },
                        },
                      ],
                    },
                  },
                },
              },
            }
          : slide,
      ),
    };
    const rawFillResult = moveShapesAcrossSlides(
      rawFillSource,
      [picture.handle],
      requireValue(rawFillSource.slides[1]?.handle),
    );
    const rawFillSourceRelationships = requireValue(
      rawFillResult.document.packageGraph.relationships.find(
        (group) => group.sourcePartPath === sourceSlide.partPath,
      ),
    );
    expect(
      rawFillSourceRelationships.relationships.map((relationship) => relationship.id),
    ).toContain(sharedRelationshipId);
  });

  it("moves a chart while byte-preserving its chart/workbook subgraph and package metadata", () => {
    const source = buildCleanTwoSlideSource((session, first, second) => {
      session.target(first).addChart({
        chartType: "bar",
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(2_400_000),
        height: asEmu(1_800_000),
        name: "Moved chart",
        series: [{ name: "Revenue", categories: ["A", "B"], values: [10, 20] }],
      });
      session.target(second).addChart({
        chartType: "line",
        offsetX: asEmu(3_000_000),
        offsetY: asEmu(0),
        width: asEmu(2_400_000),
        height: asEmu(1_800_000),
        name: "Destination chart",
        series: [{ name: "Cost", categories: ["A", "B"], values: [5, 8] }],
      });
    });
    const sourceSlide = requireValue(source.slides[0]);
    const destinationSlide = requireValue(source.slides[1]);
    const chart = sourceSlide.shapes[0];
    if (
      chart?.kind !== "chart" ||
      chart.handle === undefined ||
      chart.chartRelationshipId === undefined
    ) {
      throw new Error("chart move fixture is missing its source chart");
    }
    const chartPartPath = requireInternalTarget(
      source,
      sourceSlide.partPath,
      chart.chartRelationshipId,
    );
    const workbookRelationship = requireValue(
      source.packageGraph.relationships
        .find((group) => group.sourcePartPath === chartPartPath)
        ?.relationships.find((relationship) => relationship.type.endsWith("/package")),
    );
    const workbookPartPath = requireValue(
      resolveInternalRelationshipTarget(chartPartPath, workbookRelationship),
    );
    const chartBytes = rawPartBytes(source, chartPartPath);
    const workbookBytes = rawPartBytes(source, workbookPartPath);
    const contentTypes = structuredClone(source.packageGraph.contentTypes);
    const partPaths = source.packageGraph.parts.map((part) => part.partPath);

    const result = moveShapesAcrossSlides(
      source,
      [chart.handle],
      requireValue(destinationSlide.handle),
    );
    const movedChart = result.document.slides[1]?.shapes.find(
      (shape) => shape.kind === "chart" && shape.name === "Moved chart",
    );
    if (movedChart?.kind !== "chart" || movedChart.chartRelationshipId === undefined) {
      throw new Error("moved chart is missing");
    }
    expect(movedChart.handle?.partPath).toBe(destinationSlide.partPath);
    expect(movedChart.chartRelationshipId).not.toBe(chart.chartRelationshipId);
    expect(
      result.document.packageGraph.relationships
        .find((group) => group.sourcePartPath === sourceSlide.partPath)
        ?.relationships.some((relationship) => relationship.id === chart.chartRelationshipId),
    ).toBe(false);
    expect(
      requireInternalTarget(
        result.document,
        destinationSlide.partPath,
        movedChart.chartRelationshipId,
      ),
    ).toBe(chartPartPath);
    expect(rawPartBytes(result.document, chartPartPath)).toEqual(chartBytes);
    expect(rawPartBytes(result.document, workbookPartPath)).toEqual(workbookBytes);
    expect(result.document.packageGraph.contentTypes).toEqual(contentTypes);
    expect(result.document.packageGraph.parts.map((part) => part.partPath)).toEqual(partPaths);

    const persisted = readPptx(writePptx(result.document));
    expect(persisted.slides[0]?.shapes.some((shape) => shape.kind === "chart")).toBe(false);
    const persistedMovedChart = persisted.slides[1]?.shapes.find(
      (shape) => shape.kind === "chart" && shape.name === "Moved chart",
    );
    if (persistedMovedChart?.kind !== "chart") throw new Error("persisted chart is missing");
    expect(
      requireInternalTarget(
        persisted,
        destinationSlide.partPath,
        requireValue(persistedMovedChart.chartRelationshipId),
      ),
    ).toBe(chartPartPath);
    expect(rawPartBytes(persisted, chartPartPath)).toEqual(chartBytes);
    expect(rawPartBytes(persisted, workbookPartPath)).toEqual(workbookBytes);
    expect(persisted.packageGraph.contentTypes).toEqual(contentTypes);
    expect(new Set(persisted.packageGraph.parts.map((part) => part.partPath))).toEqual(
      new Set(partPaths),
    );
  });

  it("allows chronological chart data edits before and after a move using the current handle", () => {
    const source = buildCleanTwoSlideSource((session, first) => {
      session.target(first).addChart({
        chartType: "bar",
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(2_400_000),
        height: asEmu(1_800_000),
        name: "Edited moved chart",
        series: [{ name: "Before", categories: ["A"], values: [1] }],
      });
    });
    const sourceChart = source.slides[0]?.shapes[0];
    if (sourceChart?.kind !== "chart" || sourceChart.handle === undefined) {
      throw new Error("chronological chart fixture is missing");
    }
    const editedBeforeMove = updateChartData(source, sourceChart.handle, {
      series: [{ name: "Before move", categories: ["A", "B"], values: [2, 3] }],
    });
    const moved = moveShapesAcrossSlides(
      editedBeforeMove,
      [sourceChart.handle],
      requireValue(editedBeforeMove.slides[1]?.handle),
    );
    const movedHandle = requireValue(
      moved.moved.find((mapping) => mapping.before.nodeId === sourceChart.nodeId)?.after,
    );
    expect(expectPersistedChartXml(moved.document, 1)).toContain("Before move");
    expect(() =>
      updateChartData(moved.document, sourceChart.handle, {
        series: [{ name: "Stale", categories: ["A"], values: [9] }],
      }),
    ).toThrow("chart handle was not found");

    const editedAfterMove = updateChartData(moved.document, movedHandle, {
      series: [{ name: "After move", categories: ["A", "B"], values: [4, 5] }],
    });
    const persisted = readPptx(writePptx(editedAfterMove));
    const persistedChart = persisted.slides[1]?.shapes.find((shape) => shape.kind === "chart");
    expect(persistedChart?.kind).toBe("chart");
    const chartPartPath = requireInternalTarget(
      persisted,
      requireValue(persisted.slides[1]).partPath,
      requireValue(
        persistedChart?.kind === "chart" ? persistedChart.chartRelationshipId : undefined,
      ),
    );
    expect(new TextDecoder().decode(rawPartBytes(persisted, chartPartPath))).toContain(
      "After move",
    );
  });

  it("retains a source chart relationship while another source chart still uses it", () => {
    const source = buildCleanTwoSlideSource((session, first) => {
      for (const [index, name] of ["Move shared chart", "Keep shared chart"].entries()) {
        session.target(first).addChart({
          chartType: "bar",
          offsetX: asEmu(index * 2_600_000),
          offsetY: asEmu(0),
          width: asEmu(2_400_000),
          height: asEmu(1_800_000),
          name,
          series: [{ name, categories: ["A"], values: [index + 1] }],
        });
      }
    });
    const [move, keep] = requireValue(source.slides[0]).shapes;
    if (
      move?.kind !== "chart" ||
      keep?.kind !== "chart" ||
      move.handle === undefined ||
      keep.handle === undefined ||
      move.chartRelationshipId === undefined ||
      keep.chartRelationshipId === undefined
    ) {
      throw new Error("shared chart fixture is missing");
    }
    const sharedSource = shareChartRelationship(
      source,
      requireValue(source.slides[0]).partPath,
      requireValue(keep.nodeId),
      keep.chartRelationshipId,
      move.chartRelationshipId,
    );
    const moved = moveShapesAcrossSlides(
      sharedSource,
      [move.handle],
      requireValue(sharedSource.slides[1]?.handle),
    );
    expect(
      moved.document.packageGraph.relationships
        .find((group) => group.sourcePartPath === sharedSource.slides[0]?.partPath)
        ?.relationships.map((relationship) => relationship.id),
    ).toContain(move.chartRelationshipId);
    const persisted = readPptx(writePptx(moved.document));
    const keptChart = persisted.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    expect(keptChart?.kind === "chart" ? keptChart.chartRelationshipId : undefined).toBe(
      move.chartRelationshipId,
    );
  });

  it("rejects placeholders, non-root nodes, and unsupported typed drawing kinds", () => {
    const source = buildCleanTwoSlideSource((session, first) => {
      session.target(first).addTable({
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        columnWidths: [asEmu(1000)],
        rows: [{ height: asEmu(1000), cells: [{ text: "Unsupported table" }] }],
      });
    });
    expect(() =>
      moveShapesAcrossSlides(
        source,
        [requireValue(source.slides[0]?.shapes[0]?.handle)],
        requireValue(source.slides[1]?.handle),
      ),
    ).toThrow("outside the slide-to-slide typed slice");
  });

  it("rejects conflicting slide-root namespace prefixes atomically", () => {
    const source = buildCleanTwoSlideSource((session, first) => {
      session.target(first).addShape({ ...rect(0), name: "Moved shape" });
    });
    const sourceSlide = requireValue(source.slides[0]);
    const destinationSlide = requireValue(source.slides[1]);
    const conflicted = withSlideRootNamespace(
      withSlideRootNamespace(source, sourceSlide.partPath, "vnd", "urn:source-vendor"),
      destinationSlide.partPath,
      "vnd",
      "urn:destination-vendor",
    );
    const before = structuredClone(conflicted);

    expect(() =>
      moveShapesAcrossSlides(
        conflicted,
        [requireValue(conflicted.slides[0]?.shapes[0]?.handle)],
        requireValue(conflicted.slides[1]?.handle),
      ),
    ).toThrow("bind namespace prefix 'vnd' differently");
    expect(conflicted).toEqual(before);
  });
});

type AuthoringSession = ReturnType<typeof createPptxAuthoringSession>;

function buildCleanTwoSlideSource(
  author: (session: AuthoringSession, first: SourceHandle, second: SourceHandle) => void,
) {
  const initial = createPptx();
  const session = createPptxAuthoringSession(initial);
  const first = requireValue(initial.slides[0]?.handle);
  const layout = requireValue(initial.slideLayouts[0]?.handle);
  const second = session.addEmptySlideFromLayout({ layoutPartPath: layout.partPath });
  author(session, first, second);
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

function shapeName(shape: { readonly kind: string }): string | undefined {
  return "name" in shape && typeof shape.name === "string" ? shape.name : undefined;
}

function sharePictureRelationship(
  source: PptxSourceModel,
  slidePartPath: PptxSourceModel["slides"][number]["partPath"],
  keepNodeId: PptxSourceModel["slides"][number]["shapes"][number]["nodeId"],
  oldRelationshipId: NonNullable<
    Extract<
      PptxSourceModel["slides"][number]["shapes"][number],
      { kind: "image" }
    >["blipRelationshipId"]
  >,
  sharedRelationshipId: typeof oldRelationshipId,
): PptxSourceModel {
  const rawParts = source.packageGraph.rawParts?.map((part) =>
    part.partPath === slidePartPath && part.kind === "binary"
      ? {
          ...part,
          bytes: new TextEncoder().encode(
            new TextDecoder()
              .decode(part.bytes)
              .replaceAll(oldRelationshipId, sharedRelationshipId),
          ),
        }
      : part,
  );
  return {
    ...source,
    packageGraph: {
      ...source.packageGraph,
      ...(rawParts !== undefined ? { rawParts } : {}),
    },
    slides: source.slides.map((slide) =>
      slide.partPath !== slidePartPath
        ? slide
        : {
            ...slide,
            shapes: slide.shapes.map((shape) =>
              shape.kind === "image" && shape.nodeId === keepNodeId
                ? {
                    ...shape,
                    blipRelationshipId: sharedRelationshipId,
                    handle: {
                      ...requireValue(shape.handle),
                      relationshipId: sharedRelationshipId,
                    },
                  }
                : shape,
            ),
          },
    ),
  };
}

function shareChartRelationship(
  source: PptxSourceModel,
  slidePartPath: PptxSourceModel["slides"][number]["partPath"],
  keepNodeId: NonNullable<PptxSourceModel["slides"][number]["shapes"][number]["nodeId"]>,
  oldRelationshipId: NonNullable<
    Extract<
      PptxSourceModel["slides"][number]["shapes"][number],
      { kind: "chart" }
    >["chartRelationshipId"]
  >,
  sharedRelationshipId: typeof oldRelationshipId,
): PptxSourceModel {
  return {
    ...source,
    packageGraph: {
      ...source.packageGraph,
      rawParts: source.packageGraph.rawParts?.map((part) =>
        part.partPath === slidePartPath && part.kind === "binary"
          ? {
              ...part,
              bytes: new TextEncoder().encode(
                new TextDecoder()
                  .decode(part.bytes)
                  .replaceAll(`r:id="${oldRelationshipId}"`, `r:id="${sharedRelationshipId}"`),
              ),
            }
          : part,
      ),
    },
    slides: source.slides.map((slide) =>
      slide.partPath !== slidePartPath
        ? slide
        : {
            ...slide,
            shapes: slide.shapes.map((shape) =>
              shape.kind === "chart" && shape.nodeId === keepNodeId
                ? {
                    ...shape,
                    chartRelationshipId: sharedRelationshipId,
                    handle: {
                      ...requireValue(shape.handle),
                      relationshipId: sharedRelationshipId,
                    },
                  }
                : shape,
            ),
          },
    ),
  };
}

function textHandles(textBody: SourceTextBody): readonly SourceHandle[] {
  return [
    ...(textBody.handle === undefined ? [] : [textBody.handle]),
    ...textBody.paragraphs.flatMap((paragraph) => [
      ...(paragraph.handle === undefined ? [] : [paragraph.handle]),
      ...paragraph.runs.flatMap((run) => (run.handle === undefined ? [] : [run.handle])),
    ]),
  ];
}

function requireInternalTarget(
  source: PptxSourceModel,
  ownerPartPath: PptxSourceModel["slides"][number]["partPath"],
  relationshipId: NonNullable<
    Extract<
      PptxSourceModel["slides"][number]["shapes"][number],
      { kind: "chart" }
    >["chartRelationshipId"]
  >,
) {
  const relationship = requireValue(
    source.packageGraph.relationships
      .find((group) => group.sourcePartPath === ownerPartPath)
      ?.relationships.find((item) => item.id === relationshipId),
  );
  return requireValue(resolveInternalRelationshipTarget(ownerPartPath, relationship));
}

function rawPartBytes(source: PptxSourceModel, partPath: string): Uint8Array {
  const part = source.packageGraph.rawParts?.find((item) => item.partPath === partPath);
  if (part?.kind !== "binary") throw new Error(`raw part '${partPath}' is missing`);
  return part.bytes;
}

function expectPersistedChartXml(source: PptxSourceModel, slideIndex: number): string {
  const persisted = readPptx(writePptx(source));
  const slide = requireValue(persisted.slides[slideIndex]);
  const chart = slide.shapes.find((shape) => shape.kind === "chart");
  const relationshipId = requireValue(
    chart?.kind === "chart" ? chart.chartRelationshipId : undefined,
  );
  const chartPartPath = requireInternalTarget(persisted, slide.partPath, relationshipId);
  return new TextDecoder().decode(rawPartBytes(persisted, chartPartPath));
}

function withSlideRootNamespace(
  source: PptxSourceModel,
  slidePartPath: PptxSourceModel["slides"][number]["partPath"],
  prefix: string,
  namespace: string,
): PptxSourceModel {
  return {
    ...source,
    packageGraph: {
      ...source.packageGraph,
      rawParts: source.packageGraph.rawParts?.map((part) =>
        part.partPath === slidePartPath && part.kind === "binary"
          ? {
              ...part,
              bytes: new TextEncoder().encode(
                new TextDecoder()
                  .decode(part.bytes)
                  .replace("<p:sld ", `<p:sld xmlns:${prefix}="${namespace}" `),
              ),
            }
          : part,
      ),
    },
  };
}

function requireValue<T>(value: T | undefined): T {
  if (value === undefined) throw new Error("fixture value is missing");
  return value;
}
