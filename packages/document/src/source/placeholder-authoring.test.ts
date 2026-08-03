import { strFromU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  addPlaceholder,
  asEmu,
  createComputedView,
  createPptx,
  createPptxAuthoringSession,
  readPptx,
  replaceParagraphPlainText,
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
    expect(source.edits).toBeUndefined();
    expect(source.slideMasters[0]?.shapes).toEqual([]);
    expect(source.slideLayouts[0]?.shapes).toEqual([]);
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
