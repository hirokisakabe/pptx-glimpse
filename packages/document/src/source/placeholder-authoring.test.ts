import { strFromU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  addEmptySlideFromLayout,
  addPlaceholder,
  asEmu,
  type ComputedShapeElement,
  createComputedView,
  createPptx,
  createPptxAuthoringSession,
  readPptx,
  replaceParagraphPlainText,
  type SourceImage,
  type SourceShape,
  type SourceShapeNode,
  writePptx,
} from "../index.js";

describe("native placeholder authoring", () => {
  it("authors master/layout placeholders and materializes inherited slide shells", () => {
    const session = createPptxAuthoringSession(createPptx());
    const masterHandle = requireValue(session.source.slideMasters[0]?.handle);
    const layoutHandle = requireValue(session.source.slideLayouts[0]?.handle);
    const masterPlaceholderHandle = session.target(masterHandle).addPlaceholder({
      type: "title",
      index: 7,
      name: "Master title prompt",
      transform: {
        offsetX: asEmu(100000),
        offsetY: asEmu(200000),
        width: asEmu(4000000),
        height: asEmu(800000),
      },
      geometry: { kind: "preset", preset: "roundRect" },
      promptText: "Master prompt",
      body: { anchor: "middle" },
    });
    const layoutPlaceholderHandle = session.target(layoutHandle).addPlaceholder({
      type: "ctrTitle",
      index: 7,
      name: "Layout title prompt",
      orientation: "vertical",
      size: "half",
      promptText: "Layout prompt",
      promptProperties: { bold: true },
    });
    const slideHandle = session.addEmptySlideFromLayout({ layoutPartPath: layoutHandle.partPath });

    expect(masterPlaceholderHandle.partPath).toBe(masterHandle.partPath);
    expect(layoutPlaceholderHandle.partPath).toBe(layoutHandle.partPath);
    const slide = requireValue(
      session.source.slides.find(
        (candidate) => candidate.handle?.partPath === slideHandle.partPath,
      ),
    );
    const shell = requireShape(slide.shapes[0]);
    expect(shell).toMatchObject({
      placeholder: {
        type: "ctrTitle",
        index: 7,
        orientation: "vert",
        size: "half",
        hasCustomPrompt: true,
      },
    });
    expect(shell.transform).toBeUndefined();
    expect(shell.geometry).toBeUndefined();
    expect(shell.textBody?.paragraphs[0]?.runs).toEqual([]);

    const paragraphHandle = requireValue(shell.textBody?.paragraphs[0]?.handle);
    const populated = replaceParagraphPlainText(session.source, paragraphHandle, "User title");
    const output = writePptx(populated);
    const slideXml = strFromU8(requireValue(unzipSync(output)[slide.partPath]));
    expect(slideXml).toContain(
      '<p:ph type="ctrTitle" idx="7" orient="vert" sz="half" hasCustomPrompt="1"/>',
    );
    expect(slideXml).not.toContain("Layout prompt");
    expect(slideXml).not.toContain("Master prompt");
    const shellXml = slideXml.slice(slideXml.indexOf("Layout title prompt"));
    expect(shellXml).not.toContain("<a:xfrm>");
    expect(shellXml).not.toContain("<a:prstGeom");

    const reread = readPptx(output);
    const computedSlide = requireValue(createComputedView(reread).slides.at(-1));
    const computedTitle = computedSlide.elements.find(
      (element) => element.kind === "shape" && element.sourceNode.name === "Layout title prompt",
    );
    expect(computedTitle).toMatchObject({
      kind: "shape",
      transform: {
        offsetX: 100000,
        offsetY: 200000,
        width: 4000000,
        height: 800000,
      },
      geometry: { preset: "roundRect" },
      textBody: { properties: { anchor: "middle" } },
      placeholder: {
        type: "ctrTitle",
        index: 7,
        orientation: "vert",
        size: "half",
        hasCustomPrompt: true,
      },
      placeholderMatch: {
        layout: { name: "Layout title prompt", placeholder: { type: "ctrTitle", index: 7 } },
        master: { name: "Master title prompt", placeholder: { type: "title", index: 7 } },
      },
    });
    expect(
      computedSlide.elements.some((element) => element.sourceNode.name === "Master title prompt"),
    ).toBe(false);

    const rereadLayout = requireValue(
      reread.slideLayouts.find((candidate) => candidate.partPath === layoutHandle.partPath),
    );
    const layoutPlaceholder = requireShape(rereadLayout.shapes[0]);
    expect(layoutPlaceholder.textBody?.paragraphs[0]?.runs[0]?.properties?.bold).toBe(true);
    const ambiguous = {
      ...reread,
      slideLayouts: reread.slideLayouts.map((candidate) =>
        candidate.partPath === rereadLayout.partPath
          ? {
              ...candidate,
              shapes: [
                ...candidate.shapes,
                {
                  ...layoutPlaceholder,
                  name: "Duplicate idx",
                  nodeId: undefined,
                  handle: undefined,
                },
              ],
            }
          : candidate,
      ),
    };
    const unresolved = requireValue(createComputedView(ambiguous).slides.at(-1)).elements.find(
      (element) => element.kind === "shape" && element.sourceNode.name === "Layout title prompt",
    );
    expect(unresolved).toMatchObject({ kind: "shape" });
    if (unresolved?.kind !== "shape") throw new Error("expected unresolved placeholder shape");
    expect(unresolved.placeholderMatch).toBeUndefined();
    expect(unresolved.transform).toBeUndefined();
  });

  it("inherits transform from a placeholder-capable layout picture node", () => {
    const source = createPptx();
    const layout = requireValue(source.slideLayouts[0]);
    const picturePlaceholder: SourceImage = {
      kind: "image",
      nodeId: undefined,
      name: "Picture definition",
      transform: {
        offsetX: asEmu(300000),
        offsetY: asEmu(400000),
        width: asEmu(2000000),
        height: asEmu(1500000),
      },
      placeholder: { type: "pic", index: 8 },
      handle: { partPath: layout.partPath, orderingSlot: 0 },
    };
    const withPictureDefinition = {
      ...source,
      slideLayouts: source.slideLayouts.map((candidate) =>
        candidate.partPath === layout.partPath
          ? { ...candidate, shapes: [picturePlaceholder] }
          : candidate,
      ),
    };
    const withSlide = addEmptySlideFromLayout(withPictureDefinition, {
      layoutPartPath: layout.partPath,
    });
    const shell = requireShape(requireValue(withSlide.slides.at(-1)).shapes[0]);
    const populated = replaceParagraphPlainText(
      withSlide,
      requireValue(shell.textBody?.paragraphs[0]?.handle),
      "Picture user content",
    );
    const computed = requireValue(createComputedView(populated).slides.at(-1)).elements.find(
      (element) => element.kind === "shape" && element.sourceNode.name === "Picture definition",
    );
    expect(computed).toMatchObject({
      kind: "shape",
      transform: { offsetX: 300000, offsetY: 400000, width: 2000000, height: 1500000 },
      placeholderMatch: {
        layoutNode: { kind: "image", name: "Picture definition" },
      },
    });
  });

  it("rejects unsupported targets, duplicate effective indexes, and ambiguous inheritance atomically", () => {
    const source = createPptx();
    const masterHandle = requireValue(source.slideMasters[0]?.handle);
    const layoutHandle = requireValue(source.slideLayouts[0]?.handle);
    const slideHandle = requireValue(source.slides[0]?.handle);
    expect(() =>
      addPlaceholder(source, slideHandle, {
        type: "title",
        index: 0,
        transform: transform(),
      }),
    ).toThrow("layout or master handle was not found");
    expect(() => addPlaceholder(source, masterHandle, { type: "title", index: 0 })).toThrow(
      "master placeholders require a transform",
    );
    expect(() =>
      addPlaceholder(source, masterHandle, {
        // @ts-expect-error Runtime validation protects JavaScript callers.
        type: "ctrTitle",
        index: 0,
        transform: transform(),
      }),
    ).toThrow("unsupported for master");
    expect(() => addPlaceholder(source, layoutHandle, { type: "title", index: 0 })).toThrow(
      "one compatible master placeholder transform",
    );

    const withMaster = addPlaceholder(source, masterHandle, {
      type: "title",
      index: 0,
      transform: transform(),
    });
    const withLayout = addPlaceholder(withMaster, layoutHandle, { type: "title", index: 0 });
    expect(() =>
      addPlaceholder(withLayout, layoutHandle, {
        type: "ctrTitle",
        index: 0,
        transform: transform(),
      }),
    ).toThrow("effective index '0' is already in use");
    expect(
      withLayout.slideLayouts.find((candidate) => candidate.partPath === layoutHandle.partPath)
        ?.shapes,
    ).toHaveLength(1);

    const ambiguousMaster = addPlaceholder(withMaster, masterHandle, {
      type: "title",
      index: 1,
      transform: transform(),
    });
    expect(() =>
      addPlaceholder(ambiguousMaster, layoutHandle, {
        type: "ctrTitle",
        index: 2,
        transform: transform(),
      }),
    ).toThrow("compatible master placeholder category is ambiguous");

    const specialCollision: SourceShape = {
      kind: "shape",
      name: "Date collision",
      placeholder: { type: "dt", index: 0 },
      handle: { partPath: layoutHandle.partPath, orderingSlot: 1 },
    };
    const conflictingLayout = {
      ...withLayout,
      slideLayouts: withLayout.slideLayouts.map((candidate) =>
        candidate.partPath === layoutHandle.partPath
          ? { ...candidate, shapes: [...candidate.shapes, specialCollision] }
          : candidate,
      ),
    };
    expect(() =>
      addEmptySlideFromLayout(conflictingLayout, { layoutPartPath: layoutHandle.partPath }),
    ).toThrow("ambiguous effective placeholder index '0'");
    expect(source.edits).toBeUndefined();
    expect(source.slideMasters[0]?.shapes).toEqual([]);
    expect(source.slideLayouts[0]?.shapes).toEqual([]);
  });

  it("normalizes preset geometry and rejects forbidden XML code points", () => {
    const source = createPptx();
    const masterHandle = requireValue(source.slideMasters[0]?.handle);
    const normalized = addPlaceholder(source, masterHandle, {
      type: "title",
      index: 0,
      transform: transform(),
      geometry: { kind: "preset", preset: " roundRect " },
    });
    expect(requireShape(normalized.slideMasters[0]?.shapes[0]).geometry).toEqual({
      preset: "roundRect",
    });
    expect(() =>
      addPlaceholder(source, masterHandle, {
        type: "title",
        index: 0,
        transform: transform(),
        promptText: "invalid\ufffe",
      }),
    ).toThrow("invalid XML character");
    expect(() =>
      addPlaceholder(source, masterHandle, {
        type: "title",
        index: 0,
        transform: transform(),
        geometry: {
          kind: "preset",
          preset: "roundRect",
          adjustValues: { "adj\u0001": 50000 },
        },
      }),
    ).toThrow("invalid XML character");
  });

  it("diagnoses master category ambiguity even when no slide shell references the layout", () => {
    let source = createPptx();
    const master = requireValue(source.slideMasters[0]);
    const layout = requireValue(source.slideLayouts[0]);
    source = addPlaceholder(source, requireValue(master.handle), {
      type: "title",
      index: 0,
      transform: transform(),
    });
    source = addPlaceholder(source, requireValue(master.handle), {
      type: "body",
      index: 1,
      transform: transform(),
    });
    source = addPlaceholder(source, requireValue(layout.handle), {
      type: "ctrTitle",
      index: 2,
      transform: transform(),
    });
    const archive = unzipSync(writePptx(source));
    const masterXml = strFromU8(requireValue(archive[master.partPath])).replace(
      'type="body" idx="1"',
      'type="title" idx="1"',
    );
    const reread = readPptx(
      zipSync({ ...archive, [master.partPath]: new TextEncoder().encode(masterXml) }),
    );
    expect(reread.slides[0]?.shapes).toEqual([]);
    const diagnostic = reread.diagnostics.find(
      (candidate) => candidate.code === "placeholder-master-match-ambiguous",
    );
    expect(diagnostic).toMatchObject({
      code: "placeholder-master-match-ambiguous",
      handle: { partPath: layout.partPath },
    });
  });

  it("retains source-local p:ph attributes and identity on sp, pic, and graphicFrame", () => {
    const output = writePptx(createPptx());
    const archive = unzipSync(output);
    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = strFromU8(requireValue(archive[slidePath])).replace(
      "</p:spTree>",
      `<p:sp><p:nvSpPr><p:cNvPr id="20" name="Defaulted"/><p:cNvSpPr/>` +
        `<p:nvPr><p:ph/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>` +
        `<a:lstStyle/><a:p><a:r><a:t>Default user</a:t></a:r><a:endParaRPr/></a:p>` +
        `</p:txBody></p:sp>` +
        `<p:pic><p:nvPicPr><p:cNvPr id="21" name="Picture placeholder"/><p:cNvPicPr/>` +
        `<p:nvPr><p:ph type="pic" idx="5" orient="vert" sz="half" ` +
        `hasCustomPrompt="0"/></p:nvPr></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>` +
        `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="22" name="Raw frame"/>` +
        `<p:cNvGraphicFramePr/><p:nvPr><p:ph type="obj" idx="6" orient="horz" ` +
        `sz="quarter" hasCustomPrompt="1"/></p:nvPr></p:nvGraphicFramePr>` +
        `<p:xfrm/><a:graphic><a:graphicData uri="urn:unsupported"/></a:graphic>` +
        `</p:graphicFrame></p:spTree>`,
    );
    const reread = readPptx(
      zipSync({ ...archive, [slidePath]: new TextEncoder().encode(slideXml) }),
    );
    const shapes = reread.slides[0]?.shapes ?? [];
    expect(shapes.find((shape) => String(shape.nodeId) === "20")?.placeholder).toEqual({});
    expect(shapes.find((shape) => String(shape.nodeId) === "21")?.placeholder).toEqual({
      type: "pic",
      index: 5,
      orientation: "vert",
      size: "half",
      hasCustomPrompt: false,
    });
    expect(shapes.find((shape) => String(shape.nodeId) === "22")).toMatchObject({
      kind: "raw",
      placeholder: {
        type: "obj",
        index: 6,
        orientation: "horz",
        size: "quarter",
        hasCustomPrompt: true,
      },
      handle: { partPath: slidePath, nodeId: "22" },
    });
    const rawSlidePlaceholder = requireValue(shapes.find((shape) => String(shape.nodeId) === "22"));
    const rawMatchSource = {
      ...reread,
      slideLayouts: reread.slideLayouts.map((layout, index) =>
        index === 0 ? { ...layout, shapes: [rawSlidePlaceholder] } : layout,
      ),
    };
    const rawComputed = createComputedView(rawMatchSource).slides[0]?.elements.find(
      (element) => element.kind === "raw" && String(element.sourceNode.nodeId) === "22",
    );
    expect(rawComputed).toMatchObject({
      kind: "raw",
      placeholderMatch: { layoutNode: { kind: "raw", placeholder: { index: 6 } } },
    });
    const defaulted = createComputedView(reread).slides[0]?.elements.find(
      (element) => element.sourceNode.name === "Defaulted",
    );
    expect(defaulted).toMatchObject({
      placeholder: {
        type: "obj",
        index: 0,
        orientation: "horz",
        size: "full",
        hasCustomPrompt: false,
      },
    });
  });

  it("inherits master body transforms for content placeholders and layout-node transforms", () => {
    const source = createPptx();
    const bodyTransform = transform();
    const layoutImageTransform = {
      offsetX: asEmu(200000),
      offsetY: asEmu(300000),
      width: asEmu(1200000),
      height: asEmu(600000),
    };
    const bodyTypes = ["chart", "clipArt", "dgm", "media", "pic", "tbl"] as const;
    const layoutShapes: SourceShapeNode[] = [
      ...bodyTypes.map((type, index) => ({
        kind: "shape" as const,
        name: `Layout ${type}`,
        placeholder: { type, index: index + 1 },
      })),
      {
        kind: "image",
        name: "Layout picture node",
        placeholder: { type: "pic", index: 99 },
        transform: layoutImageTransform,
      },
    ];
    const slideShapes: SourceShapeNode[] = [
      ...bodyTypes.map((type, index) => ({
        kind: "shape" as const,
        name: `Slide ${type}`,
        placeholder: { type, index: index + 1 },
        textBody: { paragraphs: [{ runs: [{ text: type }] }] },
      })),
      {
        kind: "shape",
        name: "Slide picture shell",
        placeholder: { type: "pic", index: 99 },
        textBody: { paragraphs: [{ runs: [{ text: "picture" }] }] },
      },
    ];
    const computed = createComputedView({
      ...source,
      slideMasters: source.slideMasters.map((master, index) =>
        index === 0
          ? {
              ...master,
              shapes: [
                {
                  kind: "shape",
                  name: "Master body",
                  placeholder: { type: "body", index: 1 },
                  transform: bodyTransform,
                },
              ],
            }
          : master,
      ),
      slideLayouts: source.slideLayouts.map((layout, index) =>
        index === 0 ? { ...layout, shapes: layoutShapes } : layout,
      ),
      slides: source.slides.map((slide, index) =>
        index === 0 ? { ...slide, shapes: slideShapes } : slide,
      ),
    });

    for (const type of bodyTypes) {
      const element = computed.slides[0]?.elements.find(
        (candidate): candidate is ComputedShapeElement =>
          candidate.kind === "shape" && candidate.sourceNode.name === `Slide ${type}`,
      );
      expect(element?.transform).toEqual(bodyTransform);
      expect(element?.placeholderMatch?.masterNode.name).toBe("Master body");
    }
    const pictureShell = computed.slides[0]?.elements.find(
      (candidate): candidate is ComputedShapeElement =>
        candidate.kind === "shape" && candidate.sourceNode.name === "Slide picture shell",
    );
    expect(pictureShell?.transform).toEqual(layoutImageTransform);
    expect(pictureShell?.placeholderMatch?.layoutNode.kind).toBe("image");
  });
});

function transform() {
  return {
    offsetX: asEmu(0),
    offsetY: asEmu(0),
    width: asEmu(1000000),
    height: asEmu(500000),
  };
}

function requireShape(shape: SourceShapeNode | undefined): SourceShape {
  if (shape?.kind !== "shape") {
    throw new Error("expected a SourceShape");
  }
  return shape;
}

function requireValue<T>(value: T | undefined): T {
  if (value === undefined) throw new Error("expected value");
  return value;
}
