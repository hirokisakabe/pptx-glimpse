import { describe, expect, it } from "vitest";

import {
  asEmu,
  createPptx,
  createPptxAuthoringSession,
  moveShapesAcrossSlides,
  type PptxSourceModel,
  readPptx,
  type SourceHandle,
  type SourceTextBody,
  writePptx,
} from "../index.js";
import { BLUE_PNG } from "../writer/write-pptx.test-helpers.js";

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
  });

  it("rejects placeholders, non-root nodes, and unsupported typed drawing kinds", () => {
    const source = buildCleanTwoSlideSource((session, first) => {
      session.target(first).addChart({
        chartType: "bar",
        offsetX: asEmu(0),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        series: [{ categories: ["A"], values: [1] }],
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

function textHandles(textBody: SourceTextBody): readonly SourceHandle[] {
  return [
    ...(textBody.handle === undefined ? [] : [textBody.handle]),
    ...textBody.paragraphs.flatMap((paragraph) => [
      ...(paragraph.handle === undefined ? [] : [paragraph.handle]),
      ...paragraph.runs.flatMap((run) => (run.handle === undefined ? [] : [run.handle])),
    ]),
  ];
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
