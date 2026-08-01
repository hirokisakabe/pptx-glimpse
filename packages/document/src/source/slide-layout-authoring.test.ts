import { readFileSync } from "node:fs";

import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  addEmptySlideFromLayout,
  addSlideLayout,
  asEmu,
  cloneSlideLayout,
  createPptx,
  createPptxAuthoringSession,
  readPptx,
  writePptx,
} from "../index.js";

const PNG_BYTES = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

describe("addSlideLayout", () => {
  it("allocates layout package bookkeeping and preserves authoring order through write/read", () => {
    const original = createPptx();
    const masterHandle = requireValue(original.slideMasters[0]?.handle);
    let source = addSlideLayout(original, masterHandle, {
      name: "Title & Body",
      type: "titleOnly",
      show: false,
      background: { kind: "solid", color: { kind: "srgb", hex: "f1f5f9" } },
      margin: {
        left: asEmu(100000),
        right: asEmu(110000),
        top: asEmu(120000),
        bottom: asEmu(130000),
      },
    });
    source = addSlideLayout(source, masterHandle, {
      name: "Image Layout",
      background: { kind: "image", bytes: PNG_BYTES },
    });

    const master = source.slideMasters[0];
    expect(master?.layoutPartPaths).toEqual([
      "ppt/slideLayouts/slideLayout1.xml",
      "ppt/slideLayouts/slideLayout2.xml",
      "ppt/slideLayouts/slideLayout3.xml",
    ]);
    expect(source.slideLayouts.map((layout) => layout.partPath)).toEqual(master?.layoutPartPaths);
    expect(source.packageGraph.contentTypes.overrides).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ partName: "ppt/slideLayouts/slideLayout2.xml" }),
        expect.objectContaining({ partName: "ppt/slideLayouts/slideLayout3.xml" }),
      ]),
    );
    const masterRels = requireValue(
      source.packageGraph.relationships.find(
        (relationships) => relationships.sourcePartPath === master?.partPath,
      ),
    );
    expect(masterRels.relationships.map((relationship) => relationship.id)).toEqual([
      "rId1",
      "rId2",
      "rId3",
      "rId4",
    ]);
    expect(new Set(masterRels.relationships.map((relationship) => relationship.id)).size).toBe(4);

    const output = writePptx(source);
    const files = unzipSync(output);
    const masterXml = strFromU8(requireValue(files["ppt/slideMasters/slideMaster1.xml"]));
    expect(masterXml).toContain('<p:sldLayoutId id="2147483649" r:id="rId1"');
    expect(masterXml).toContain('<p:sldLayoutId id="2147483650" r:id="rId3"');
    expect(masterXml).toContain('<p:sldLayoutId id="2147483651" r:id="rId4"');
    expect(masterXml.indexOf('r:id="rId3"')).toBeLessThan(masterXml.indexOf('r:id="rId4"'));

    const reread = readPptx(output);
    expect(reread.slideMasters[0]?.layoutPartPaths).toEqual(master?.layoutPartPaths);
    const solidLayout = reread.slideLayouts[1];
    expect(solidLayout).toMatchObject({
      name: "Title & Body",
      type: "titleOnly",
      show: false,
      masterPartPath: master?.partPath,
      background: {
        kind: "fill",
        fill: { kind: "solid", color: { kind: "srgb", hex: "F1F5F9" } },
      },
    });
    const imageLayout = reread.slideLayouts[2];
    expect(imageLayout).toMatchObject({ name: "Image Layout", type: "blank", show: true });
    expect(imageLayout?.background?.kind).toBe("fill");
    expect(reread.packageGraph.media).toHaveLength(1);
  });

  it("returns a session handle that accepts drawing APIs and drives slide authoring defaults", () => {
    const session = createPptxAuthoringSession(createPptx());
    const masterHandle = requireValue(session.source.slideMasters[0]?.handle);
    const layoutHandle = session.addSlideLayout(masterHandle, {
      name: "Session Layout",
      margin: {
        left: asEmu(100000),
        right: asEmu(110000),
        top: asEmu(120000),
        bottom: asEmu(130000),
      },
    });
    const layoutShape = session.target(layoutHandle).addShape({
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000000),
      height: asEmu(500000),
    });
    const slideHandle = session.addEmptySlideFromLayout({
      layoutPartPath: layoutHandle.partPath,
    });
    const textHandle = session.target(slideHandle).addTextBox({
      offsetX: asEmu(100000),
      offsetY: asEmu(100000),
      width: asEmu(2000000),
      height: asEmu(500000),
      text: "Uses added layout",
    });

    expect(layoutShape.partPath).toBe(layoutHandle.partPath);
    expect(textHandle.partPath).toBe(slideHandle.partPath);
    const reread = readPptx(writePptx(session.source));
    const authoredSlide = reread.slides.at(-1);
    expect(authoredSlide?.layoutPartPath).toBe(layoutHandle.partPath);
    const textShape = authoredSlide?.shapes.find((shape) => shape.kind === "shape");
    expect(textShape?.kind).toBe("shape");
    if (textShape?.kind !== "shape") throw new Error("authored text box was not reread");
    expect(textShape.textBody?.properties).toMatchObject({
      marginLeft: 100000,
      marginRight: 110000,
      marginTop: 120000,
      marginBottom: 130000,
    });
  });

  it("rejects values outside the public input contract without changing the source", () => {
    const source = createPptx();
    const masterHandle = requireValue(source.slideMasters[0]?.handle);
    expect(() => addSlideLayout(source, masterHandle, { name: " " })).toThrow(
      "name must be a non-empty string",
    );
    expect(() =>
      addSlideLayout(source, masterHandle, {
        name: "Invalid type",
        // @ts-expect-error Runtime validation protects JavaScript callers.
        type: "not-a-layout-type",
      }),
    ).toThrow("type must be a supported OOXML slide layout type");
    expect(() =>
      addSlideLayout(source, masterHandle, {
        name: "Invalid show",
        // @ts-expect-error Runtime validation protects JavaScript callers.
        show: "yes",
      }),
    ).toThrow("show must be a boolean");
    expect(() =>
      addSlideLayout(source, masterHandle, {
        name: "Invalid margin",
        margin: {
          left: asEmu(-1),
          right: asEmu(0),
          top: asEmu(0),
          bottom: asEmu(0),
        },
      }),
    ).toThrow("margin.left must be a finite non-negative EMU value");
    for (const margin of [null, 1, [], { left: 0 }, { left: 0, right: 0, top: 0 }]) {
      expect(() =>
        addSlideLayout(source, masterHandle, {
          name: "Malformed margin",
          // @ts-expect-error Runtime validation protects JavaScript callers.
          margin,
        }),
      ).toThrow();
    }
    for (const left of [Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(() =>
        addSlideLayout(source, masterHandle, {
          name: "Non-finite margin",
          margin: {
            left: asEmu(left),
            right: asEmu(0),
            top: asEmu(0),
            bottom: asEmu(0),
          },
        }),
      ).toThrow("margin.left must be a finite non-negative EMU value");
    }
    expect(source.slideLayouts).toHaveLength(1);
  });

  it("materializes relationship-order layouts before transition content when the id list is absent", () => {
    const original = createPptx();
    const master = requireValue(original.slideMasters[0]);
    const masterHandle = requireValue(master.handle);
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const source = {
      ...original,
      packageGraph: {
        ...original.packageGraph,
        rawParts: requireValue(original.packageGraph.rawParts).map((part) => {
          if (part.partPath !== master.partPath || part.kind !== "binary") return part;
          const xml = decoder
            .decode(part.bytes)
            .replace(/<p:sldLayoutIdLst>[\s\S]*?<\/p:sldLayoutIdLst>/, "")
            .replace("<p:txStyles>", "<p:transition/><p:timing/><p:hf/><p:txStyles>");
          return { ...part, bytes: encoder.encode(xml) };
        }),
      },
    };

    const edited = addSlideLayout(source, masterHandle, { name: "Added after fallback" });
    const output = writePptx(edited);
    const masterXml = strFromU8(
      requireValue(unzipSync(output)["ppt/slideMasters/slideMaster1.xml"]),
    );
    const layoutListIndex = masterXml.indexOf("<p:sldLayoutIdLst>");
    expect(layoutListIndex).toBeGreaterThan(masterXml.indexOf("<p:clrMap"));
    expect(layoutListIndex).toBeLessThan(masterXml.indexOf("<p:transition"));
    expect(masterXml.indexOf("<p:transition")).toBeLessThan(masterXml.indexOf("<p:timing"));
    expect(masterXml.indexOf("<p:timing")).toBeLessThan(masterXml.indexOf("<p:hf"));
    expect(masterXml).toContain('<p:sldLayoutId id="2147483649" r:id="rId1"');
    expect(masterXml).toContain('<p:sldLayoutId id="2147483650" r:id="rId3"');
    expect(readPptx(output).slideMasters[0]?.layoutPartPaths).toEqual([
      "ppt/slideLayouts/slideLayout1.xml",
      "ppt/slideLayouts/slideLayout2.xml",
    ]);
  });

  it("writes the first inserted entry into an existing empty layout id list", () => {
    const original = createPptx();
    const master = requireValue(original.slideMasters[0]);
    const layout = requireValue(original.slideLayouts[0]);
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const source = {
      ...original,
      packageGraph: {
        ...original.packageGraph,
        rawParts: requireValue(original.packageGraph.rawParts).map((part) => {
          if (part.partPath !== master.partPath || part.kind !== "binary") return part;
          const xml = decoder
            .decode(part.bytes)
            .replace(/<p:sldLayoutIdLst>[\s\S]*?<\/p:sldLayoutIdLst>/, "<p:sldLayoutIdLst/>");
          return { ...part, bytes: encoder.encode(xml) };
        }),
      },
    };

    const edited = cloneSlideLayout(source, requireValue(layout.handle), { name: "First Entry" });
    const cloneEdit = requireValue(edited.edits?.at(-1));
    if (cloneEdit.kind !== "cloneSlideLayout") throw new Error("clone edit was not recorded");
    const masterXml = strFromU8(requireValue(unzipSync(writePptx(edited))[master.partPath]));
    expect(masterXml).toContain(
      `<p:sldLayoutId id="${cloneEdit.newLayoutNumericId}" r:id="${cloneEdit.newRelationshipId}"`,
    );
  });

  it("supports the immutable function flow when adding a slide from the new layout", () => {
    const source = createPptx();
    const masterHandle = requireValue(source.slideMasters[0]?.handle);
    const withLayout = addSlideLayout(source, masterHandle, { name: "Function Layout" });
    const addedLayout = requireValue(withLayout.slideLayouts.at(-1));
    const withSlide = addEmptySlideFromLayout(withLayout, {
      layoutPartPath: addedLayout.partPath,
    });
    expect(withSlide.slides.at(-1)?.layoutPartPath).toBe(addedLayout.partPath);
  });

  it("adds a layout to a master loaded from an existing PPTX", () => {
    const source = readPptx(
      readFileSync(new URL("../../../../shared-fixtures/real-basic-theme.pptx", import.meta.url)),
    );
    const master = requireValue(source.slideMasters[0]);
    const masterHandle = requireValue(master.handle);
    const edited = addSlideLayout(source, masterHandle, {
      name: "Existing Template Addition",
      type: "cust",
    });
    const added = requireValue(edited.slideLayouts.at(-1));
    const reread = readPptx(writePptx(edited));

    expect(added.masterPartPath).toBe(master.partPath);
    expect(edited.slideMasters[0]?.layoutPartPaths.at(-1)).toBe(added.partPath);
    expect(reread.slideMasters[0]?.layoutPartPaths).toEqual(
      edited.slideMasters[0]?.layoutPartPaths,
    );
    expect(reread.slideLayouts.at(-1)).toMatchObject({
      partPath: added.partPath,
      masterPartPath: master.partPath,
      name: "Existing Template Addition",
      type: "cust",
      show: true,
    });
  });
});

describe("cloneSlideLayout", () => {
  it("preserves raw layout XML and child order while inserting a renamed clone", () => {
    const input = readFileSync(
      new URL("../../../../shared-fixtures/real-basic-theme.pptx", import.meta.url),
    );
    const source = readPptx(input);
    const layout = requireValue(source.slideLayouts[0]);
    const layoutHandle = requireValue(layout.handle);
    const edited = cloneSlideLayout(source, layoutHandle, {
      name: "Cloned & $ Layout",
      insertAt: 0,
    });
    const cloned = requireValue(edited.slideLayouts[0]);

    expect(cloned).toMatchObject({
      name: "Cloned & $ Layout",
      masterPartPath: layout.masterPartPath,
    });
    expect(cloned.shapes).toHaveLength(layout.shapes.length);
    expect(cloned.partPath).not.toBe(layout.partPath);
    expect(edited.slideMasters[0]?.layoutPartPaths[0]).toBe(cloned.partPath);

    const originalFiles = unzipSync(input);
    const output = writePptx(edited);
    const outputFiles = unzipSync(output);
    const originalXml = strFromU8(requireValue(originalFiles[layout.partPath]));
    const clonedXml = strFromU8(requireValue(outputFiles[cloned.partPath]));
    expect(clonedXml).toBe(
      originalXml.replace('<p:cSld name="TITLE">', '<p:cSld name="Cloned &amp; $ Layout">'),
    );

    const reread = readPptx(output);
    expect(reread.slideMasters[0]?.layoutPartPaths[0]).toBe(cloned.partPath);
    expect(
      reread.slideLayouts.find((candidate) => candidate.partPath === cloned.partPath),
    ).toMatchObject({
      partPath: cloned.partPath,
      masterPartPath: layout.masterPartPath,
      name: "Cloned & $ Layout",
    });
    expect(
      reread.slideLayouts.find((candidate) => candidate.partPath === cloned.partPath)?.shapes,
    ).toHaveLength(layout.shapes.length);
  });

  it("renames the real cSld element when a comment contains a decoy start tag", () => {
    const input = readFileSync(
      new URL("../../../../shared-fixtures/real-basic-theme.pptx", import.meta.url),
    );
    const original = readPptx(input);
    const layout = requireValue(original.slideLayouts[0]);
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const decoy = '<!-- <p:cSld name="DECOY"> -->';
    const source = {
      ...original,
      packageGraph: {
        ...original.packageGraph,
        rawParts: requireValue(original.packageGraph.rawParts).map((part) => {
          if (part.partPath !== layout.partPath || part.kind !== "binary") return part;
          return {
            ...part,
            bytes: encoder.encode(decoder.decode(part.bytes).replace("<p:cSld", `${decoy}<p:cSld`)),
          };
        }),
      },
    };

    const edited = cloneSlideLayout(source, requireValue(layout.handle), { name: "Real Name" });
    const cloned = requireValue(
      edited.slideLayouts.find((candidate) => candidate.name === "Real Name"),
    );
    const clonedXml = strFromU8(requireValue(unzipSync(writePptx(edited))[cloned.partPath]));
    expect(clonedXml).toContain(decoy);
    expect(clonedXml).toContain('<p:cSld name="Real Name">');
    expect(clonedXml).not.toContain('<p:cSld name="DECOY" name="Real Name">');
  });

  it("rebases raw shape handles so the cloned layout accepts shape ordering edits", () => {
    const input = readFileSync(
      new URL("../../../../shared-fixtures/real-basic-theme.pptx", import.meta.url),
    );
    const original = readPptx(input);
    const layout = requireValue(original.slideLayouts[0]);
    const group = requireValue(layout.shapes.find((shape) => shape.kind === "group"));
    if (group.kind !== "group") throw new Error("test fixture group shape is missing");
    const rawShape = {
      kind: "raw" as const,
      nodeId: group.nodeId,
      raw: requireValue(group.rawSidecars?.[0]),
      handle: requireValue(group.handle),
    };
    const source = {
      ...original,
      slideLayouts: original.slideLayouts.map((candidate) =>
        candidate.partPath === layout.partPath
          ? {
              ...candidate,
              shapes: candidate.shapes.map((shape) => (shape === group ? rawShape : shape)),
            }
          : candidate,
      ),
    };
    const session = createPptxAuthoringSession(source);
    const clonedHandle = session.cloneSlideLayout(requireValue(layout.handle), {
      name: "Raw Clone",
    });
    const cloned = requireValue(
      session.source.slideLayouts.find((candidate) => candidate.partPath === clonedHandle.partPath),
    );
    const clonedRaw = requireValue(cloned.shapes.find((shape) => shape.kind === "raw"));

    expect(clonedRaw.handle?.partPath).toBe(cloned.partPath);
    const handles = cloned.shapes.map((shape) => requireValue(shape.handle));
    expect(() => session.target(clonedHandle).reorderShapes([...handles].reverse())).not.toThrow();
  });

  it("shares relationship targets, allocates collisions, and returns an authoring handle", () => {
    let source = createPptx();
    const masterHandle = requireValue(source.slideMasters[0]?.handle);
    source = addSlideLayout(source, masterHandle, {
      name: "Image Source",
      background: { kind: "image", bytes: PNG_BYTES },
    });
    const sourceLayout = requireValue(source.slideLayouts.at(-1));
    const sourceLayoutHandle = requireValue(sourceLayout.handle);
    const mediaBefore = source.packageGraph.media;
    const existingPartPaths = new Set(source.packageGraph.parts.map((part) => part.partPath));
    const master = requireValue(source.slideMasters[0]);
    const existingMasterRelationshipIds = new Set(
      requireValue(
        source.packageGraph.relationships.find(
          (relationships) => relationships.sourcePartPath === master.partPath,
        ),
      ).relationships.map((relationship) => relationship.id),
    );
    const masterXmlBefore = strFromU8(requireValue(unzipSync(writePptx(source))[master.partPath]));
    const existingNumericIds = new Set(
      [...masterXmlBefore.matchAll(/<p:sldLayoutId id="(\d+)"/g)].map((match) => Number(match[1])),
    );

    source = cloneSlideLayout(source, sourceLayoutHandle, {
      name: "Image Clone One",
      insertAt: 1,
    });
    source = cloneSlideLayout(source, sourceLayoutHandle, {
      name: "Image Clone Two",
      insertAt: 3,
    });

    const clones = source.slideLayouts.filter((layout) => layout.name?.startsWith("Image Clone"));
    expect(clones.map((layout) => layout.partPath)).toEqual([
      "ppt/slideLayouts/slideLayout3.xml",
      "ppt/slideLayouts/slideLayout4.xml",
    ]);
    expect(source.packageGraph.media).toEqual(mediaBefore);
    const sourceRelationships = requireValue(
      source.packageGraph.relationships.find(
        (relationships) => relationships.sourcePartPath === sourceLayout.partPath,
      ),
    );
    for (const clone of clones) {
      const relationships = requireValue(
        source.packageGraph.relationships.find(
          (candidate) => candidate.sourcePartPath === clone.partPath,
        ),
      );
      expect(relationships.relationships).toEqual(sourceRelationships.relationships);
    }

    const session = createPptxAuthoringSession(source);
    const clonedHandle = session.cloneSlideLayout(sourceLayoutHandle, {
      name: "Session Clone",
      insertAt: 0,
    });
    session.target(clonedHandle).addShape({
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000000),
      height: asEmu(500000),
    });

    const reread = readPptx(writePptx(session.source));
    expect(
      reread.slideLayouts.find((layout) => layout.partPath === clonedHandle.partPath),
    ).toMatchObject({
      partPath: clonedHandle.partPath,
      name: "Session Clone",
    });
    expect(
      reread.slideLayouts.find((layout) => layout.partPath === clonedHandle.partPath)?.shapes,
    ).toHaveLength(sourceLayout.shapes.length + 1);
    const cloneEdits = session.source.edits?.filter((edit) => edit.kind === "cloneSlideLayout");
    expect(new Set(cloneEdits?.map((edit) => edit.newLayoutPartPath)).size).toBe(3);
    expect(new Set(cloneEdits?.map((edit) => edit.newRelationshipId)).size).toBe(3);
    expect(new Set(cloneEdits?.map((edit) => edit.newLayoutNumericId)).size).toBe(3);
    for (const edit of requireValue(cloneEdits)) {
      expect(existingPartPaths.has(edit.newLayoutPartPath)).toBe(false);
      expect(existingMasterRelationshipIds.has(edit.newRelationshipId)).toBe(false);
      expect(existingNumericIds.has(edit.newLayoutNumericId)).toBe(false);
    }
  });

  it("inserts immediately after the source layout when insertAt is omitted", () => {
    const original = createPptx();
    const layout = requireValue(original.slideLayouts[0]);
    const edited = cloneSlideLayout(original, requireValue(layout.handle), {
      name: "Default Position",
    });
    const cloned = requireValue(
      edited.slideLayouts.find((candidate) => candidate.name === "Default Position"),
    );
    expect(edited.slideMasters[0]?.layoutPartPaths).toEqual([layout.partPath, cloned.partPath]);
  });

  it("rejects invalid positions and layouts with pending raw-part edits", () => {
    const original = createPptx();
    const layout = requireValue(original.slideLayouts[0]);
    const layoutHandle = requireValue(layout.handle);
    for (const insertAt of [2, -1, 0.5, Number.NaN]) {
      expect(() =>
        cloneSlideLayout(original, layoutHandle, { name: "Invalid position", insertAt }),
      ).toThrow("insertAt must be an integer insertion index");
    }

    const dirty = createPptxAuthoringSession(original);
    dirty.target(layoutHandle).addShape({
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000000),
      height: asEmu(500000),
    });
    expect(() => cloneSlideLayout(dirty.source, layoutHandle, { name: "Dirty clone" })).toThrow(
      "pending dirty part edits is unsupported",
    );
  });
});

function requireValue<T>(value: T | undefined): T {
  if (value === undefined) throw new Error("test fixture value is missing");
  return value;
}
