import { describe, expect, it } from "vitest";

import {
  addConnector,
  addEmptySlideFromLayout,
  addPicture,
  addShape,
  addTextBox,
  clearParagraphProperties,
  clearTextRunProperties,
  deleteShape,
  deleteSlide,
  duplicateSlide,
  findParagraphBySourceHandle,
  findShapeNodeBySourceHandle,
  findTextRunBySourceHandle,
  moveSlide,
  replaceImageBytes,
  replaceParagraphPlainText,
  replaceTextRunPlainText,
  setParagraphProperties,
  setShapeFill,
  setShapeOutline,
  setTextRunProperties,
  updateShapeTransform,
} from "./editing.js";
import { asPartPath, asRelationshipId, asSourceNodeId, type SourceHandle } from "./handles.js";
import type { PptxSourceModel } from "./pptx-source-model.js";
import type { SourceConnector, SourceImage, SourceShape } from "./shapes.js";
import { asEmu, asHundredthPt, asOoxmlAngle, asOoxmlPercent, asPt } from "./units.js";

const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const PRESENTATION_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const SLIDE_LAYOUT_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_LAYOUT_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3]);
const JPEG = new Uint8Array([0xff, 0xd8, 0xff, 1, 2, 3]);

describe("editing source helpers", () => {
  it("finds text runs, paragraphs, and shape nodes by source handle", () => {
    const source = buildSourceModel();
    const run = firstRun(source);
    const paragraph = firstParagraph(source);
    const shape = shapeByName(source, "Title");

    expect(findTextRunBySourceHandle(source, requireHandle(run.handle))).toBe(run);
    expect(findParagraphBySourceHandle(source, requireHandle(paragraph.handle))).toBe(paragraph);
    expect(findShapeNodeBySourceHandle(source, requireHandle(shape.handle))).toBe(shape);
    expect(findTextRunBySourceHandle(source, handle("ppt/slides/slide1.xml", "missing"))).toBe(
      undefined,
    );
  });
});

describe("editing text operations", () => {
  it("records a text run replacement without mutating the input source", () => {
    const source = buildSourceModel();
    const runHandle = requireHandle(firstRun(source).handle);

    const edited = expectNonMutating(source, () =>
      replaceTextRunPlainText(source, runHandle, "Edited"),
    );

    expect(firstRun(source).text).toBe("Original");
    expect(firstRun(edited).text).toBe("Edited");
    expect(edited.edits).toEqual([
      { kind: "replaceTextRunPlainText", handle: runHandle, text: "Edited" },
    ]);
  });

  it("returns the original source when replacing a text run with the same text", () => {
    const source = buildSourceModel();
    const runHandle = requireHandle(firstRun(source).handle);

    const edited = replaceTextRunPlainText(source, runHandle, firstRun(source).text);

    expect(edited).toBe(source);
    expect(source.edits).toBeUndefined();
  });

  it("does not append another text run edit when the replacement is already applied", () => {
    const source = buildSourceModel();
    const runHandle = requireHandle(firstRun(source).handle);
    const edited = replaceTextRunPlainText(source, runHandle, "Edited");

    const repeated = replaceTextRunPlainText(edited, runHandle, "Edited");

    expect(repeated).toBe(edited);
    expect(repeated.edits).toHaveLength(1);
  });

  it("sets and clears text run properties in source and edit records", () => {
    const source = buildSourceModel();
    const runHandle = requireHandle(firstRun(source).handle);

    const set = expectNonMutating(source, () =>
      setTextRunProperties(source, runHandle, {
        bold: false,
        underline: true,
        fontSize: asPt(18),
        color: { kind: "srgb", hex: "00aa44" },
        typeface: "Aptos",
      }),
    );
    const cleared = expectNonMutating(set, () =>
      clearTextRunProperties(set, runHandle, ["italic", "typeface"]),
    );

    expect(firstRun(set).properties).toMatchObject({
      bold: false,
      italic: true,
      underline: true,
      fontSize: 18,
      color: { kind: "srgb", hex: "00aa44" },
      typeface: "Aptos",
    });
    expect(firstRun(cleared).properties).toMatchObject({
      bold: false,
      underline: true,
      fontSize: 18,
      color: { kind: "srgb", hex: "00aa44" },
    });
    expect(firstRun(cleared).properties?.italic).toBeUndefined();
    expect(firstRun(cleared).properties?.typeface).toBeUndefined();
    expect(cleared.edits?.map((edit) => edit.kind)).toEqual([
      "updateTextRunProperties",
      "updateTextRunProperties",
    ]);
    expect(cleared.edits?.[0]).toMatchObject({
      kind: "updateTextRunProperties",
      set: {
        bold: false,
        underline: true,
        fontSize: 18,
        color: { kind: "srgb", hex: "00aa44" },
        typeface: "Aptos",
      },
    });
    expect(cleared.edits?.[1]).toMatchObject({
      kind: "updateTextRunProperties",
      clear: ["italic", "typeface"],
    });
  });

  it("returns the original source when clearing properties that are already absent", () => {
    const source = buildSourceModel();
    const runHandle = requireHandle(firstRun(source).handle);

    const edited = clearTextRunProperties(source, runHandle, ["underline"]);

    expect(edited).toBe(source);
    expect(source.edits).toBeUndefined();
  });

  it("returns the original source for equivalent text properties with different key order", () => {
    const source = structuredClone(buildSourceModel());
    const runHandle = requireHandle(firstRun(source).handle);
    (firstRun(source) as { properties?: unknown }).properties = {
      color: { hex: "00aa44", kind: "srgb" },
      bold: true,
      italic: true,
      fontSize: asPt(24),
      typeface: "Calibri",
    };

    const edited = setTextRunProperties(source, runHandle, {
      color: { kind: "srgb", hex: "00aa44" },
    });

    expect(edited).toBe(source);
    expect(source.edits).toBeUndefined();
  });

  it("supports text edits addressed by shapeSlot handles", () => {
    const source = buildSourceModel();
    const slotShape = shapeByName(source, "Slot Text");
    const slotRun = slotShape.textBody?.paragraphs[0]?.runs[0];
    const slotRunHandle = requireHandle(slotRun?.handle);

    expect(slotRunHandle).toMatchObject({
      partPath: "ppt/slides/slide1.xml",
      nodeId: "text:shapeSlot:3:p:0:r:0",
      orderingSlot: 0,
    });
    expect(findTextRunBySourceHandle(source, slotRunHandle)).toBe(slotRun);

    const edited = expectNonMutating(source, () =>
      replaceTextRunPlainText(source, slotRunHandle, "Slot Edited"),
    );

    expect(shapeByName(edited, "Slot Text").textBody?.paragraphs[0]?.runs[0]?.text).toBe(
      "Slot Edited",
    );
    expect(edited.edits).toEqual([
      { kind: "replaceTextRunPlainText", handle: slotRunHandle, text: "Slot Edited" },
    ]);
  });

  it("normalizes a paragraph replacement to one run and records the edit", () => {
    const source = buildSourceModel();
    const paragraphHandle = requireHandle(firstParagraph(source).handle);

    const edited = expectNonMutating(source, () =>
      replaceParagraphPlainText(source, paragraphHandle, "Paragraph"),
    );

    expect(firstParagraph(source).runs).toHaveLength(2);
    expect(firstParagraph(edited).runs).toEqual([
      {
        kind: "textRun",
        text: "Paragraph",
        properties: firstRun(source).properties,
        handle: firstRun(source).handle,
      },
    ]);
    expect(edited.edits).toEqual([
      { kind: "replaceParagraphPlainText", handle: paragraphHandle, text: "Paragraph" },
    ]);
  });

  it("sets and clears paragraph properties in source and edit records", () => {
    const source = buildSourceModel();
    const paragraphHandle = requireHandle(firstParagraph(source).handle);

    const set = expectNonMutating(source, () =>
      setParagraphProperties(source, paragraphHandle, {
        align: "right",
        level: 2,
        bullet: { type: "char", char: "\u2022" },
      }),
    );
    const cleared = expectNonMutating(set, () =>
      clearParagraphProperties(set, paragraphHandle, ["align"]),
    );

    expect(firstParagraph(set).properties).toMatchObject({
      align: "right",
      level: 2,
      bullet: { type: "char", char: "\u2022" },
    });
    expect(firstParagraph(cleared).properties).toMatchObject({
      level: 2,
      bullet: { type: "char", char: "\u2022" },
    });
    expect(firstParagraph(cleared).properties?.align).toBeUndefined();
    expect(cleared.edits?.map((edit) => edit.kind)).toEqual([
      "updateParagraphProperties",
      "updateParagraphProperties",
    ]);
    expect(cleared.edits?.[0]).toMatchObject({
      kind: "updateParagraphProperties",
      set: { align: "right", level: 2, bullet: { type: "char", char: "\u2022" } },
    });
    expect(cleared.edits?.[1]).toMatchObject({
      kind: "updateParagraphProperties",
      clear: ["align"],
    });
  });

  it("returns the original source for equivalent paragraph properties", () => {
    const source = buildSourceModel();
    const paragraphHandle = requireHandle(firstParagraph(source).handle);
    const set = setParagraphProperties(source, paragraphHandle, {
      align: "center",
      bullet: { type: "none" },
    });

    const repeated = setParagraphProperties(set, paragraphHandle, {
      align: "center",
      bullet: { type: "none" },
    });

    expect(repeated).toBe(set);
    expect(repeated.edits).toHaveLength(1);
  });

  it("returns the original source when clearing paragraph properties that are already absent", () => {
    const source = buildSourceModel();
    const paragraphHandle = requireHandle(firstParagraph(source).handle);

    const edited = clearParagraphProperties(source, paragraphHandle, ["bullet"]);

    expect(edited).toBe(source);
    expect(source.edits).toBeUndefined();
  });

  it("rejects invalid text property patches with operation-specific errors", () => {
    const source = buildSourceModel();
    const runHandle = requireHandle(firstRun(source).handle);

    expect(() => setTextRunProperties(source, runHandle, { fontSize: asPt(0) })).toThrow(
      "updateTextRunProperties: fontSize must be a finite positive pt value",
    );
    expect(() =>
      setTextRunProperties(source, runHandle, { color: { kind: "srgb", hex: "abc" } }),
    ).toThrow("updateTextRunProperties: srgb text run color must be a 6-digit hex value");
    expect(() =>
      // @ts-expect-error exercises runtime validation for JS callers.
      clearTextRunProperties(source, runHandle, ["strikethrough"]),
    ).toThrow("updateTextRunProperties: unsupported text run property 'strikethrough'");
  });

  it("rejects invalid paragraph property patches with operation-specific errors", () => {
    const source = buildSourceModel();
    const paragraphHandle = requireHandle(firstParagraph(source).handle);

    expect(() =>
      // @ts-expect-error exercises runtime validation for JS callers.
      setParagraphProperties(source, paragraphHandle, { align: "middle" }),
    ).toThrow("updateParagraphProperties: align must be left, center, right, or justify");
    expect(() => setParagraphProperties(source, paragraphHandle, { level: 9 })).toThrow(
      "updateParagraphProperties: level must be an integer from 0 to 8",
    );
    expect(() =>
      setParagraphProperties(source, paragraphHandle, { bullet: { type: "char", char: "" } }),
    ).toThrow("updateParagraphProperties: bullet.char must be a non-empty string");
    expect(() =>
      // @ts-expect-error exercises runtime validation for JS callers.
      clearParagraphProperties(source, paragraphHandle, ["spacing"]),
    ).toThrow("updateParagraphProperties: unsupported paragraph property 'spacing'");
  });
});

describe("editing shape operations", () => {
  it("updates a top-level shape transform and records the edit", () => {
    const source = buildSourceModel();
    const shapeHandle = requireHandle(shapeByName(source, "Title").handle);
    const transform = {
      offsetX: asEmu(11),
      offsetY: asEmu(22),
      width: asEmu(33),
      height: asEmu(44),
    };

    const edited = expectNonMutating(source, () =>
      updateShapeTransform(source, shapeHandle, transform),
    );

    expect(shapeByName(source, "Title").transform).toMatchObject({
      offsetX: 100,
      offsetY: 200,
      width: 3000,
      height: 1000,
    });
    expect(shapeByName(edited, "Title").transform).toMatchObject(transform);
    expect(edited.edits).toEqual([
      { kind: "updateShapeTransform", handle: shapeHandle, ...transform },
    ]);
  });

  it("returns the original source when updating a shape transform to the same values", () => {
    const source = buildSourceModel();
    const shape = shapeByName(source, "Title");
    const shapeHandle = requireHandle(shape.handle);
    const transform = {
      offsetX: shape.transform!.offsetX,
      offsetY: shape.transform!.offsetY,
      width: shape.transform!.width,
      height: shape.transform!.height,
    };

    const edited = updateShapeTransform(source, shapeHandle, transform);

    expect(edited).toBe(source);
    expect(source.edits).toBeUndefined();
  });

  it("does not append another shape transform edit when the transform is already applied", () => {
    const source = buildSourceModel();
    const shapeHandle = requireHandle(shapeByName(source, "Title").handle);
    const transform = {
      offsetX: asEmu(11),
      offsetY: asEmu(22),
      width: asEmu(33),
      height: asEmu(44),
    };
    const edited = updateShapeTransform(source, shapeHandle, transform);

    const repeated = updateShapeTransform(edited, shapeHandle, transform);

    expect(repeated).toBe(edited);
    expect(repeated.edits).toHaveLength(1);
  });

  it("sets shape fill and shape or connector outlines idempotently", () => {
    const source = buildSourceModel();
    const shapeHandle = requireHandle(shapeByName(source, "Title").handle);
    const withFill = expectNonMutating(source, () =>
      setShapeFill(source, shapeHandle, {
        kind: "solid",
        color: { kind: "srgb", hex: "00aa44" },
      }),
    );
    const withOutline = setShapeOutline(withFill, shapeHandle, {
      width: asEmu(12700),
      fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } },
    });
    const withoutFillOrOutline = setShapeOutline(
      setShapeFill(withOutline, shapeHandle, { kind: "none" }),
      shapeHandle,
      { fill: { kind: "none" } },
    );
    const repeated = setShapeOutline(withoutFillOrOutline, shapeHandle, {
      fill: { kind: "none" },
    });

    expect(shapeByName(source, "Title").fill).toBeUndefined();
    expect(shapeByName(withFill, "Title").fill).toEqual({
      kind: "solid",
      color: { kind: "srgb", hex: "00aa44" },
    });
    expect(shapeByName(withOutline, "Title").outline).toEqual({
      width: 12700,
      fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } },
    });
    expect(shapeByName(withoutFillOrOutline, "Title").fill).toEqual({ kind: "none" });
    expect(shapeByName(withoutFillOrOutline, "Title").outline).toEqual({
      width: 12700,
      fill: { kind: "none" },
    });
    expect(repeated).toBe(withoutFillOrOutline);
    expect(withoutFillOrOutline.edits).toEqual([
      { kind: "updateShapeFill", handle: shapeHandle, fill: { kind: "none" } },
      {
        kind: "updateShapeOutline",
        handle: shapeHandle,
        outline: { width: asEmu(12700), fill: { kind: "none" } },
      },
    ]);
  });

  it("sets connector outlines without changing fill", () => {
    const source = buildSourceModel();
    const withConnector = addConnector(source, requireHandle(source.slides[0].handle), {
      preset: "straightConnector1",
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
      start: {
        shapeHandle: requireHandle(shapeByName(source, "Title").handle),
        connectionSiteIndex: 0,
      },
      end: {
        shapeHandle: requireHandle(shapeByName(source, "Body").handle),
        connectionSiteIndex: 0,
      },
    });
    const connectorHandle = requireHandle(connectorByName(withConnector, "Connector 31").handle);

    const edited = setShapeOutline(withConnector, connectorHandle, {
      width: asEmu(19050),
      fill: { kind: "solid", color: { kind: "srgb", hex: "FF00AA" } },
    });

    expect(connectorByName(edited, "Connector 31").outline).toEqual({
      width: 19050,
      fill: { kind: "solid", color: { kind: "srgb", hex: "FF00AA" } },
    });
    expect(() => setShapeFill(edited, connectorHandle, { kind: "none" })).toThrow(
      "setShapeFill: only top-level sp shapes support fill edits",
    );
  });

  it("adds a text box with a collision-free id and finalized XML", () => {
    const source = buildSourceModel();

    const edited = expectNonMutating(source, () =>
      addTextBox(source, requireHandle(source.slides[0].handle), {
        offsetX: asEmu(500),
        offsetY: asEmu(600),
        width: asEmu(700),
        height: asEmu(800),
        text: "007",
      }),
    );
    const added = shapeByName(edited, "TextBox 31");

    expect(
      source.slides[0].shapes.map((shape) => shape.kind !== "raw" && shape.name),
    ).not.toContain("TextBox 31");
    expect(added.nodeId).toBe("31");
    expect(added.textBody?.paragraphs[0]?.runs[0]?.text).toBe("007");
    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "addTextBox",
      slidePartPath: "ppt/slides/slide1.xml",
      shapeId: "31",
    });
    expect(edited.edits?.at(-1)).toHaveProperty("xml", expect.stringContaining("TextBox 31"));
  });

  it("adds a formatted text box through finalized XML and source model parsing", () => {
    const source = buildSourceModel();

    const edited = expectNonMutating(source, () =>
      addTextBox(source, requireHandle(source.slides[0].handle), {
        offsetX: asEmu(500),
        offsetY: asEmu(600),
        width: asEmu(700),
        height: asEmu(800),
        rotation: asOoxmlAngle(900000),
        name: "Formatted TextBox",
        body: {
          anchor: "middle",
          marginLeft: asEmu(1000),
          marginRight: asEmu(2000),
          marginTop: asEmu(3000),
          marginBottom: asEmu(4000),
        },
        paragraphs: [
          {
            properties: {
              align: "center",
              lineSpacing: { type: "points", value: asHundredthPt(1800) },
            },
            runs: [
              {
                text: "Gradient",
                properties: {
                  fontFace: "Aptos",
                  fontSize: asPt(24),
                  bold: true,
                  gradientFill: {
                    gradientType: "linear",
                    angle: asOoxmlAngle(5400000),
                    stops: [
                      { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FF0000" } },
                      {
                        position: asOoxmlPercent(100000),
                        color: { kind: "srgb", hex: "0000FF" },
                      },
                    ],
                  },
                },
              },
              {
                text: " solid",
                properties: {
                  italic: true,
                  color: { kind: "srgb", hex: "00AA44" },
                },
              },
            ],
          },
          {
            runs: [{ text: "Second paragraph" }],
          },
        ],
      }),
    );
    const added = shapeByName(edited, "Formatted TextBox");
    const lastEdit = edited.edits?.at(-1);
    if (lastEdit?.kind !== "addTextBox") {
      throw new Error("expected addTextBox edit");
    }
    const xml = lastEdit.xml;

    expect(added.transform).toMatchObject({
      offsetX: 500,
      offsetY: 600,
      width: 700,
      height: 800,
      rotation: 900000,
    });
    expect(added.textBody?.properties).toMatchObject({
      anchor: "middle",
      marginLeft: 1000,
      marginRight: 2000,
      marginTop: 3000,
      marginBottom: 4000,
    });
    expect(added.textBody?.paragraphs).toHaveLength(2);
    expect(added.textBody?.paragraphs[0]?.properties).toMatchObject({
      align: "center",
      lineSpacing: { type: "pts", value: 1800 },
    });
    expect(added.textBody?.paragraphs[0]?.runs.map((run) => run.text)).toEqual([
      "Gradient",
      " solid",
    ]);
    expect(added.textBody?.paragraphs[0]?.runs[0]?.properties).toMatchObject({
      bold: true,
      fontSize: 24,
      typeface: "Aptos",
    });
    expect(added.textBody?.paragraphs[0]?.runs[1]?.properties).toMatchObject({
      italic: true,
      color: { kind: "srgb", hex: "00AA44" },
    });
    expect(added.textBody?.paragraphs[1]?.runs[0]?.text).toBe("Second paragraph");
    expect(xml).toContain("<a:gradFill");
    expect(xml).toContain('<a:pPr algn="ctr">');
    expect(xml).toContain('anchor="ctr" lIns="1000" rIns="2000"');
  });

  it("adds preset geometry shapes with finalized XML and parsed source styling", () => {
    const source = buildSourceModel();

    const withRect = addShape(source, requireHandle(source.slides[0].handle), {
      geometry: { kind: "preset", preset: " rect " },
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(300),
      height: asEmu(400),
      fill: { kind: "solid", color: { kind: "srgb", hex: "112233" } },
      name: "Shape rect",
    });
    const withRoundRect = addShape(withRect, requireHandle(withRect.slides[0].handle), {
      geometry: { kind: "preset", preset: "roundRect" },
      offsetX: asEmu(500),
      offsetY: asEmu(600),
      width: asEmu(700),
      height: asEmu(800),
      name: "Shape roundRect",
    });
    const withEllipse = addShape(withRoundRect, requireHandle(withRoundRect.slides[0].handle), {
      geometry: { kind: "preset", preset: "ellipse" },
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(1100),
      height: asEmu(1200),
      name: "Shape ellipse",
    });
    const edited = expectNonMutating(withEllipse, () =>
      addShape(withEllipse, requireHandle(withEllipse.slides[0].handle), {
        geometry: { kind: "preset", preset: "line" },
        offsetX: asEmu(1300),
        offsetY: asEmu(1400),
        width: asEmu(1500),
        height: asEmu(1600),
        rotation: asOoxmlAngle(900000),
        outline: {
          width: asEmu(12700),
          fill: { kind: "solid", color: { kind: "srgb", hex: "00AAFF" } },
          dash: "lgDashDotDot",
          headEnd: { type: "oval", width: "sm", length: "sm" },
          tailEnd: { type: "triangle", width: "med", length: "lg" },
        },
        effects: {
          glow: { radius: asEmu(25400), color: { kind: "srgb", hex: "FF00AA" } },
        },
        paragraphs: [{ runs: [{ text: "Line label" }] }],
        name: "Shape line",
      }),
    );
    const rect = shapeByName(edited, "Shape rect");
    const line = shapeByName(edited, "Shape line");
    const lastEdit = edited.edits?.at(-1);
    if (lastEdit?.kind !== "addShape") throw new Error("expected addShape edit");

    expect(
      ["Shape rect", "Shape roundRect", "Shape ellipse", "Shape line"].map(
        (name) => shapeByName(edited, name).geometry,
      ),
    ).toEqual([
      { preset: "rect" },
      { preset: "roundRect" },
      { preset: "ellipse" },
      { preset: "line" },
    ]);
    expect(rect.fill).toEqual({ kind: "solid", color: { kind: "srgb", hex: "112233" } });
    expect(line.transform).toMatchObject({ rotation: 900000 });
    expect(line.outline).toMatchObject({
      width: 12700,
      fill: { kind: "solid", color: { kind: "srgb", hex: "00AAFF" } },
      dashStyle: "lgDashDotDot",
      headEnd: { type: "oval", width: "sm", length: "sm" },
      tailEnd: { type: "triangle", width: "med", length: "lg" },
    });
    expect(line.effects?.glow).toEqual({
      radius: 25400,
      color: { kind: "srgb", hex: "FF00AA" },
    });
    expect(line.textBody?.paragraphs[0]?.runs[0]?.text).toBe("Line label");
    expect(lastEdit.xml).toContain(`<a:prstGeom prst="line"`);
    expect(edited.edits?.find((edit) => edit.kind === "addShape")?.xml).toContain(
      `<a:prstGeom prst="rect"`,
    );
    expect(lastEdit.xml).toContain(`<a:prstDash val="lgDashDotDot"`);
    expect(lastEdit.xml).toContain(`<a:tailEnd type="triangle" w="med" len="lg"`);
  });

  it("rejects invalid shape and picture shadow effects", () => {
    const source = buildSourceModel();
    const slideHandle = requireHandle(source.slides[0].handle);
    const shapeInput = {
      geometry: { kind: "preset" as const, preset: "rect" },
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
    };
    const outerShadow = {
      blurRadius: asEmu(0),
      distance: asEmu(0),
      direction: asOoxmlAngle(0),
      color: { kind: "srgb" as const, hex: "000000" },
      alignment: "b" as const,
      rotateWithShape: false,
    };

    expect(() =>
      addShape(source, slideHandle, {
        ...shapeInput,
        effects: { outerShadow: { ...outerShadow, blurRadius: asEmu(-1) } },
      }),
    ).toThrow(
      "addShape: effects.outerShadow.blurRadius must be an integer between 0 and 2147483647",
    );
    expect(() =>
      addShape(source, slideHandle, {
        ...shapeInput,
        effects: { outerShadow: { ...outerShadow, distance: asEmu(1.5) } },
      }),
    ).toThrow("addShape: effects.outerShadow.distance must be an integer between 0 and 2147483647");
    expect(() =>
      addShape(source, slideHandle, {
        ...shapeInput,
        effects: {
          outerShadow: { ...outerShadow, blurRadius: asEmu(2147483647) },
        },
      }),
    ).not.toThrow();
    expect(() =>
      addShape(source, slideHandle, {
        ...shapeInput,
        effects: {
          outerShadow: { ...outerShadow, blurRadius: asEmu(2147483648) },
        },
      }),
    ).toThrow(
      "addShape: effects.outerShadow.blurRadius must be an integer between 0 and 2147483647",
    );
    expect(() =>
      addShape(source, slideHandle, {
        ...shapeInput,
        effects: {
          outerShadow: { ...outerShadow, direction: asOoxmlAngle(21600000) },
        },
      }),
    ).toThrow("addShape: effects.outerShadow.direction must be an integer between 0 and 21599999");
    expect(() =>
      addShape(source, slideHandle, {
        ...shapeInput,
        effects: {
          outerShadow: {
            ...outerShadow,
            // @ts-expect-error Runtime validation rejects unsupported rectangle alignments.
            alignment: "middle",
          },
        },
      }),
    ).toThrow("addShape: effects.outerShadow.alignment is not supported");
    expect(() =>
      addPicture(source, slideHandle, {
        bytes: PNG,
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
        effects: {
          innerShadow: {
            blurRadius: asEmu(0),
            distance: asEmu(2147483647),
            direction: asOoxmlAngle(0),
            color: { kind: "srgb", hex: "000000" },
          },
        },
      }),
    ).not.toThrow();
    expect(() =>
      addPicture(source, slideHandle, {
        bytes: PNG,
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
        effects: {
          innerShadow: {
            blurRadius: asEmu(0),
            distance: asEmu(2147483648),
            direction: asOoxmlAngle(0),
            color: { kind: "srgb", hex: "000000" },
          },
        },
      }),
    ).toThrow(
      "addPicture: effects.innerShadow.distance must be an integer between 0 and 2147483647",
    );
  });

  it("adds adjusted and custom geometry with flips and zero line extents", () => {
    let edited = buildSourceModel();
    const slideHandle = requireHandle(edited.slides[0].handle);
    edited = addShape(edited, slideHandle, {
      geometry: { kind: "preset", preset: "roundRect", adjustValues: { adj: 25000 } },
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(300),
      height: asEmu(400),
      flipHorizontal: true,
      name: "Adjusted",
    });
    edited = addShape(edited, slideHandle, {
      geometry: {
        kind: "custom",
        paths: [
          {
            width: 100,
            height: 100,
            commands: [
              { kind: "moveTo", x: 0, y: 100 },
              { kind: "lineTo", x: 50, y: 0 },
              { kind: "lineTo", x: 100, y: 100 },
              { kind: "close" },
            ],
          },
        ],
      },
      offsetX: asEmu(500),
      offsetY: asEmu(600),
      width: asEmu(700),
      height: asEmu(800),
      flipVertical: true,
      name: "Custom",
    });
    edited = addShape(edited, slideHandle, {
      geometry: { kind: "preset", preset: "line" },
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(0),
      height: asEmu(1200),
      flipHorizontal: true,
      flipVertical: true,
      name: "Vertical Line",
    });
    edited = addShape(edited, slideHandle, {
      geometry: { kind: "preset", preset: "line" },
      offsetX: asEmu(1300),
      offsetY: asEmu(1400),
      width: asEmu(1500),
      height: asEmu(0),
      name: "Horizontal Line",
    });

    expect(shapeByName(edited, "Adjusted")).toMatchObject({
      geometry: { preset: "roundRect", adjustValues: { adj: 25000 } },
      transform: { flipHorizontal: true },
    });
    expect(shapeByName(edited, "Custom")).toMatchObject({
      geometry: {
        kind: "custom",
        paths: [{ width: 100, height: 100, commands: "M 0 100 L 50 0 L 100 100 Z" }],
      },
      transform: { flipVertical: true },
    });
    expect(shapeByName(edited, "Vertical Line").transform).toMatchObject({
      width: 0,
      height: 1200,
      flipHorizontal: true,
      flipVertical: true,
    });
    expect(shapeByName(edited, "Horizontal Line").transform).toMatchObject({
      width: 1500,
      height: 0,
    });

    const xml = edited.edits
      ?.filter((edit) => edit.kind === "addShape")
      .map((edit) => edit.xml)
      .join("");
    expect(xml).toContain(
      `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>`,
    );
    expect(xml).toContain(
      `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="50" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>`,
    );
    expect(xml).toContain(`<a:xfrm flipH="1" flipV="1">`);
    expect(xml).toContain(`<a:ext cx="0" cy="1200"/>`);
    expect(xml).toContain(`<a:ext cx="1500" cy="0"/>`);
  });

  it("rejects invalid formatted text box inputs before editing", () => {
    const source = buildSourceModel();
    const slideHandle = requireHandle(source.slides[0].handle);
    const formattedInput: Parameters<typeof addTextBox>[2] = {
      offsetX: asEmu(500),
      offsetY: asEmu(600),
      width: asEmu(700),
      height: asEmu(800),
      paragraphs: [
        {
          properties: {
            align: "center",
            lineSpacing: asHundredthPt(1800),
          },
          runs: [
            {
              text: "formatted",
              properties: {
                bold: true,
                gradientFill: {
                  gradientType: "linear",
                  angle: asOoxmlAngle(0),
                  stops: [
                    { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FF0000" } },
                    {
                      position: asOoxmlPercent(100000),
                      color: { kind: "srgb", hex: "0000FF" },
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    };

    const invalidCases: readonly {
      readonly input: Parameters<typeof addTextBox>[2];
      readonly expected: RegExp;
    }[] = [
      {
        input: { ...formattedInput, rotation: asOoxmlAngle(1.5) },
        expected: /rotation must be a finite integer/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [
                {
                  text: "bad boolean",
                  properties: {
                    // @ts-expect-error Runtime validation rejects non-boolean flags from JS callers.
                    bold: "false",
                  },
                },
              ],
            },
          ],
        },
        expected: /bold must be a boolean value/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [
                {
                  text: "bad underline",
                  properties: {
                    // @ts-expect-error Runtime validation rejects non-object underline values.
                    underline: "sng",
                  },
                },
              ],
            },
          ],
        },
        expected: /underline must be a boolean or underline options object/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              properties: {
                lineSpacing: { type: "points", value: asHundredthPt(1.5) },
              },
              runs: [{ text: "bad line spacing" }],
            },
          ],
        },
        expected: /lineSpacing.value must be a finite integer/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              properties: {
                lineSpacing: { type: "points", value: asHundredthPt(158401) },
              },
              runs: [{ text: "point spacing out of range" }],
            },
          ],
        },
        expected: /lineSpacing.value must be between 0 and 158400/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              properties: {
                lineSpacing: { type: "percent", value: asOoxmlPercent(13200001) },
              },
              runs: [{ text: "percent spacing out of range" }],
            },
          ],
        },
        expected: /lineSpacing.value must be between 0 and 13200000/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              properties: {
                bullet: {
                  type: "character",
                  character: "•",
                  size: asOoxmlPercent(24000),
                },
              },
              runs: [{ text: "bad bullet size" }],
            },
          ],
        },
        expected: /bullet.size must be between 25000 and 400000/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              properties: {
                bullet: {
                  type: "auto-number",
                  scheme: "arabicPeriod",
                  startAt: 0,
                },
              },
              runs: [{ text: "bad start" }],
            },
          ],
        },
        expected: /bullet.startAt must be between 1 and 32767/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [{ text: "bad char spacing", properties: { charSpacing: 1.5 } }],
            },
          ],
        },
        expected: /charSpacing must be a finite integer/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [{ text: "bad char spacing range", properties: { charSpacing: 400001 } }],
            },
          ],
        },
        expected: /charSpacing must be between -400000 and 400000/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [
                {
                  text: "bad gradient",
                  properties: {
                    gradientFill: {
                      gradientType: "linear",
                      stops: [
                        { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FF0000" } },
                        {
                          position: asOoxmlPercent(100.5),
                          color: { kind: "srgb", hex: "0000FF" },
                        },
                      ],
                    },
                  },
                },
              ],
            },
          ],
        },
        expected: /position must be a finite integer/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [
                {
                  text: "bad baseline",
                  properties: {
                    // @ts-expect-error Runtime validation rejects ambiguous numeric baseline values.
                    baseline: 30,
                  },
                },
              ],
            },
          ],
        },
        expected: /baseline must be subscript, superscript, or a percent/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [
                {
                  text: "bad baseline percentage",
                  properties: {
                    baseline: { type: "percent", value: asOoxmlPercent(400001) },
                  },
                },
              ],
            },
          ],
        },
        expected: /baseline must be subscript, superscript, or a percent/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [
                {
                  text: "bad color",
                  properties: { color: { kind: "srgb", hex: "#112233" } },
                },
              ],
            },
          ],
        },
        expected: /color must be an srgb 6-digit hex color/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [
                {
                  text: "bad negative alpha",
                  properties: {
                    color: {
                      kind: "srgb",
                      hex: "112233",
                      transforms: [{ kind: "alpha", value: asOoxmlPercent(-1) }],
                    },
                  },
                },
              ],
            },
          ],
        },
        expected: /transforms\[0\]\.value must be between 0 and 100000/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [
                {
                  text: "bad large alpha",
                  properties: {
                    color: {
                      kind: "srgb",
                      hex: "112233",
                      transforms: [{ kind: "alpha", value: asOoxmlPercent(100001) }],
                    },
                  },
                },
              ],
            },
          ],
        },
        expected: /transforms\[0\]\.value must be between 0 and 100000/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [
                {
                  text: "bad fractional alpha",
                  properties: {
                    color: {
                      kind: "srgb",
                      hex: "112233",
                      transforms: [{ kind: "alpha", value: asOoxmlPercent(1.5) }],
                    },
                  },
                },
              ],
            },
          ],
        },
        expected: /transforms\[0\]\.value must be a finite integer/,
      },
      {
        input: {
          ...formattedInput,
          // @ts-expect-error Runtime validation rejects null body values from JS callers.
          body: null,
        },
        expected: /body must be an object/,
      },
      {
        input: {
          ...formattedInput,
          body: {
            // @ts-expect-error Runtime validation rejects unsupported auto-fit modes.
            autoFit: "normal",
          },
        },
        expected: /body.autoFit is not supported/,
      },
      {
        input: {
          ...formattedInput,
          // @ts-expect-error Runtime validation rejects null paragraph values from JS callers.
          paragraphs: [null],
        },
        expected: /paragraphs\[0\] must be an object/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [
                {
                  text: "bad properties",
                  // @ts-expect-error Runtime validation rejects null properties from JS callers.
                  properties: null,
                },
              ],
            },
          ],
        },
        expected: /properties must be an object/,
      },
      {
        input: {
          ...formattedInput,
          paragraphs: [
            {
              runs: [
                {
                  text: "bad color object",
                  properties: {
                    // @ts-expect-error Runtime validation rejects null colors from JS callers.
                    color: null,
                  },
                },
              ],
            },
          ],
        },
        expected: /color must be an srgb 6-digit hex color/,
      },
    ];

    invalidCases.forEach(({ input, expected }) => {
      expectInvalidEdit(source, () => addTextBox(source, slideHandle, input), expected);
    });
  });

  it("adds a connector using source target handles and records finalized XML", () => {
    const source = buildSourceModel();
    const start = shapeByName(source, "Title");
    const end = shapeByName(source, "Body");

    const edited = expectNonMutating(source, () =>
      addConnector(source, requireHandle(source.slides[0].handle), {
        preset: "bentConnector3",
        offsetX: asEmu(10),
        offsetY: asEmu(20),
        width: asEmu(30),
        height: asEmu(40),
        start: { shapeHandle: requireHandle(start.handle), connectionSiteIndex: 1 },
        end: { shapeHandle: requireHandle(end.handle), connectionSiteIndex: 3 },
        outline: {
          width: asEmu(12700),
          fill: { kind: "solid", color: { kind: "srgb", hex: "00AAFF" } },
          dash: "lgDashDotDot",
          headEnd: { type: "oval", width: "sm", length: "sm" },
          tailEnd: { type: "triangle", width: "med", length: "lg" },
        },
      }),
    );
    const connector = edited.slides[0].shapes.at(-1);

    expect(connector).toMatchObject({
      kind: "connector",
      nodeId: "31",
      name: "Connector 31",
      connection: {
        start: { shapeId: "10", connectionSiteIndex: 1 },
        end: { shapeId: "30", connectionSiteIndex: 3 },
      },
      geometry: { preset: "bentConnector3" },
      outline: {
        width: 12700,
        fill: { kind: "solid", color: { kind: "srgb", hex: "00AAFF" } },
        dashStyle: "lgDashDotDot",
        headEnd: { type: "oval", width: "sm", length: "sm" },
        tailEnd: { type: "triangle", width: "med", length: "lg" },
      },
    });
    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "addConnector",
      slidePartPath: "ppt/slides/slide1.xml",
      shapeId: "31",
      startShapeId: "10",
      endShapeId: "30",
    });
    expect(edited.edits?.at(-1)).toHaveProperty("xml", expect.stringContaining("bentConnector3"));
    expect(edited.edits?.at(-1)).toHaveProperty("xml", expect.stringContaining('<a:ln w="12700">'));
    expect(edited.edits?.at(-1)).toHaveProperty(
      "xml",
      expect.stringContaining('<a:prstDash val="lgDashDotDot"'),
    );
  });

  it("rejects invalid connector outline styles", () => {
    const source = buildSourceModel();
    const slideHandle = requireHandle(source.slides[0].handle);
    const baseInput = {
      preset: "straightConnector1" as const,
      offsetX: asEmu(10),
      offsetY: asEmu(20),
      width: asEmu(30),
      height: asEmu(40),
    };

    expect(() =>
      addConnector(source, slideHandle, {
        ...baseInput,
        outline: { width: asEmu(0) },
      }),
    ).toThrow("addConnector: outline.width must be a finite positive EMU value");
    expect(() =>
      addConnector(source, slideHandle, {
        ...baseInput,
        outline: {
          // @ts-expect-error Runtime validation rejects unsupported dash styles from JS callers.
          dash: "invalid",
        },
      }),
    ).toThrow("addConnector: outline.dash is not supported");
    expect(() =>
      addConnector(source, slideHandle, {
        ...baseInput,
        outline: {
          fill: {
            kind: "solid",
            // @ts-expect-error Runtime validation rejects malformed colors from JS callers.
            color: { kind: "srgb", hex: "xyz" },
          },
        },
      }),
    ).toThrow("addConnector: outline.fill.color must be an srgb 6-digit hex color");
  });

  it("adds a free connector without native connection sites", () => {
    const source = buildSourceModel();

    const edited = expectNonMutating(source, () =>
      addConnector(source, requireHandle(source.slides[0].handle), {
        preset: "straightConnector1",
        offsetX: asEmu(10),
        offsetY: asEmu(20),
        width: asEmu(30),
        height: asEmu(40),
        outline: { tailEnd: { type: "triangle", width: "med", length: "med" } },
      }),
    );
    const connector = edited.slides[0].shapes.at(-1);

    expect(connector).toMatchObject({
      kind: "connector",
      nodeId: "31",
      name: "Connector 31",
      geometry: { preset: "straightConnector1" },
      outline: { tailEnd: { type: "triangle" } },
    });
    expect(connector).not.toHaveProperty("connection.start");
    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "addConnector",
      slidePartPath: "ppt/slides/slide1.xml",
      shapeId: "31",
    });
    expect(edited.edits?.at(-1)).not.toHaveProperty("startShapeId");
    expect(edited.edits?.at(-1)).not.toHaveProperty("endShapeId");
  });

  it("preserves empty connector outline compatibility with the default black line", () => {
    const source = buildSourceModel();
    const edited = addConnector(source, requireHandle(source.slides[0].handle), {
      preset: "straightConnector1",
      offsetX: asEmu(10),
      offsetY: asEmu(20),
      width: asEmu(30),
      height: asEmu(40),
      outline: {},
    });

    expect(edited.slides[0].shapes.at(-1)).toMatchObject({
      kind: "connector",
      outline: { fill: { kind: "solid", color: { kind: "srgb", hex: "000000" } } },
    });
    expect(edited.edits?.at(-1)).toHaveProperty(
      "xml",
      expect.stringContaining('<a:solidFill><a:srgbClr val="000000"'),
    );
  });

  it("rejects connector endpoints outside the target part or OOXML index range", () => {
    const source = buildSourceModel();
    const slideHandle = requireHandle(source.slides[0].handle);
    const titleHandle = requireHandle(shapeByName(source, "Title").handle);
    const secondSlideShapeHandle = requireHandle(shapeByName(source, "Second").handle);
    const baseInput = {
      preset: "straightConnector1" as const,
      offsetX: asEmu(10),
      offsetY: asEmu(20),
      width: asEmu(30),
      height: asEmu(40),
    };

    expect(() =>
      addConnector(source, slideHandle, {
        ...baseInput,
        start: { shapeHandle: secondSlideShapeHandle, connectionSiteIndex: 0 },
      }),
    ).toThrow("addConnector: start target shape belongs to a different drawing part");
    expect(() =>
      addConnector(source, slideHandle, {
        ...baseInput,
        start: {
          shapeHandle: handle("ppt/slides/slide1.xml", "missing"),
          connectionSiteIndex: 0,
        },
      }),
    ).toThrow("addConnector: start shape handle was not found in the target drawing part");
    expect(() =>
      addConnector(source, slideHandle, {
        ...baseInput,
        start: { shapeHandle: titleHandle, connectionSiteIndex: -1 },
      }),
    ).toThrow("addConnector: start.connectionSiteIndex must be an unsigned 32-bit integer");
    expect(() =>
      addConnector(source, slideHandle, {
        ...baseInput,
        end: { shapeHandle: titleHandle, connectionSiteIndex: 0x1_0000_0000 },
      }),
    ).toThrow("addConnector: end.connectionSiteIndex must be an unsigned 32-bit integer");
  });

  it("rejects an existing connector as a connection target", () => {
    const source = buildSourceModel();
    const slideHandle = requireHandle(source.slides[0].handle);
    const withConnector = addConnector(source, slideHandle, {
      preset: "straightConnector1",
      offsetX: asEmu(10),
      offsetY: asEmu(20),
      width: asEmu(30),
      height: asEmu(40),
      name: "Existing connector",
    });

    expect(() =>
      addConnector(withConnector, slideHandle, {
        preset: "straightConnector1",
        offsetX: asEmu(50),
        offsetY: asEmu(60),
        width: asEmu(70),
        height: asEmu(80),
        start: {
          shapeHandle: requireHandle(connectorByName(withConnector, "Existing connector").handle),
          connectionSiteIndex: 0,
        },
      }),
    ).toThrow("addConnector: start target must be a top-level sp shape");
  });

  it("deletes an existing shape, drops stale text edits, and cancels added shapes followed by deleteShape", () => {
    const source = buildSourceModel();
    const existingHandle = requireHandle(shapeByName(source, "Title").handle);

    const deletedExisting = expectNonMutating(source, () => deleteShape(source, existingHandle));
    const withTextEdit = replaceTextRunPlainText(
      source,
      requireHandle(firstRun(source).handle),
      "Stale",
    );
    const deletedTextEditedShape = expectNonMutating(withTextEdit, () =>
      deleteShape(withTextEdit, existingHandle),
    );
    const withTextBox = addTextBox(source, requireHandle(source.slides[0].handle), {
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
      text: "Temporary",
      name: "Temporary",
    });
    const deletedAdded = expectNonMutating(withTextBox, () =>
      deleteShape(withTextBox, requireHandle(shapeByName(withTextBox, "Temporary").handle)),
    );
    const withConnector = addConnector(source, requireHandle(source.slides[0].handle), {
      preset: "straightConnector1",
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
      name: "Temporary connector",
    });
    const deletedAddedConnector = expectNonMutating(withConnector, () =>
      deleteShape(
        withConnector,
        requireHandle(
          withConnector.slides[0].shapes.find(
            (shape) => shape.kind === "connector" && shape.name === "Temporary connector",
          )?.handle,
        ),
      ),
    );
    const withShape = addShape(source, requireHandle(source.slides[0].handle), {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
      name: "Temporary shape",
    });
    const deletedAddedShape = expectNonMutating(withShape, () =>
      deleteShape(withShape, requireHandle(shapeByName(withShape, "Temporary shape").handle)),
    );

    expect(
      deletedExisting.slides[0].shapes.map((shape) => shape.kind !== "raw" && shape.name),
    ).not.toContain("Title");
    expect(deletedExisting.edits).toEqual([{ kind: "deleteShape", handle: existingHandle }]);
    expect(deletedTextEditedShape.edits).toEqual([{ kind: "deleteShape", handle: existingHandle }]);
    expect(deletedAdded.edits).toEqual([]);
    expect(
      deletedAdded.slides[0].shapes.map((shape) => shape.kind !== "raw" && shape.name),
    ).not.toContain("Temporary");
    expect(deletedAddedConnector.edits).toEqual([]);
    expect(
      deletedAddedConnector.slides[0].shapes.map((shape) => shape.kind !== "raw" && shape.name),
    ).not.toContain("Temporary connector");
    expect(deletedAddedShape.edits).toEqual([]);
    expect(
      deletedAddedShape.slides[0].shapes.map((shape) => shape.kind !== "raw" && shape.name),
    ).not.toContain("Temporary shape");
  });

  it("rejects unsupported shape edits with stable error messages", () => {
    const source = buildSourceModel();
    const imageHandle = requireHandle(imageByName(source, "Picture").handle);
    const noNodeHandle = { partPath: asPartPath("ppt/slides/slide1.xml") };

    expect(() =>
      updateShapeTransform(source, noNodeHandle, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow("updateShapeTransform: shape transform edit requires a node id");
    expect(() => deleteShape(source, imageHandle)).toThrow(
      "deleteShape: only top-level sp or cxnSp shapes can be deleted",
    );
    expect(() =>
      addConnector(source, requireHandle(source.slides[0].handle), {
        // @ts-expect-error exercises runtime validation for JS callers.
        preset: "arc",
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
        start: {
          shapeHandle: requireHandle(shapeByName(source, "Title").handle),
          connectionSiteIndex: 0,
        },
        end: {
          shapeHandle: requireHandle(shapeByName(source, "Body").handle),
          connectionSiteIndex: 0,
        },
      }),
    ).toThrow(
      "addConnector: preset must be straightConnector1, bentConnector3, or curvedConnector3",
    );
    expect(() =>
      addShape(source, requireHandle(source.slides[0].handle), {
        geometry: { kind: "preset", preset: "" },
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow("addShape: geometry.preset must be a non-empty string");
    expect(() =>
      addShape(source, requireHandle(source.slides[0].handle), {
        geometry: { kind: "preset", preset: "rect" },
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
        outline: {
          // @ts-expect-error Runtime validation rejects unsupported dash values.
          dash: "badDash",
        },
      }),
    ).toThrow("addShape: outline.dash is not supported");
    expect(() =>
      addShape(source, requireHandle(source.slides[0].handle), {
        geometry: { kind: "preset", preset: "rect" },
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
        fill: {
          kind: "gradient",
          gradientType: "linear",
          stops: [{ position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FF0000" } }],
        },
      }),
    ).toThrow("addShape: fill.stops must contain at least two stops");
    expect(() =>
      addShape(source, requireHandle(source.slides[0].handle), {
        geometry: { kind: "preset", preset: "rect" },
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
        fill: {
          kind: "gradient",
          gradientType: "radial",
          centerX: asOoxmlPercent(100001),
          centerY: asOoxmlPercent(50000),
          stops: [
            { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FF0000" } },
            {
              position: asOoxmlPercent(100000),
              color: { kind: "srgb", hex: "0000FF" },
            },
          ],
        },
      }),
    ).toThrow("addShape: fill.centerX must be between 0 and 100000");
    expect(() =>
      addShape(source, requireHandle(source.slides[0].handle), {
        geometry: { kind: "preset", preset: "rect" },
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
        body: { anchor: "middle" },
      }),
    ).toThrow("addShape: body requires text or paragraphs");
    expect(() =>
      addShape(source, requireHandle(source.slides[0].handle), {
        geometry: { kind: "preset", preset: "rect" },
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
        // @ts-expect-error Runtime validation rejects empty effects from JS callers.
        effects: {},
      }),
    ).toThrow("addShape: effects must set glow, outerShadow, or innerShadow");

    const baseShape = {
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
    };
    expect(() =>
      addShape(source, requireHandle(source.slides[0].handle), {
        ...baseShape,
        geometry: { kind: "preset", preset: "roundRect", adjustValues: { adj: 1.5 } },
      }),
    ).toThrow("addShape: geometry.adjustValues.adj must be a finite integer");
    expect(() =>
      addShape(source, requireHandle(source.slides[0].handle), {
        ...baseShape,
        geometry: { kind: "preset", preset: "roundRect", adjustValues: { " adj ": 25000 } },
      }),
    ).toThrow(
      "addShape: geometry.adjustValues names must be non-empty and have no surrounding whitespace",
    );
    expect(() =>
      addShape(source, requireHandle(source.slides[0].handle), {
        ...baseShape,
        geometry: { kind: "custom", paths: [] },
      }),
    ).toThrow("addShape: geometry.paths must contain at least one path");
    expect(() =>
      addShape(source, requireHandle(source.slides[0].handle), {
        ...baseShape,
        geometry: {
          kind: "custom",
          paths: [
            {
              width: 100,
              height: 100,
              commands: [{ kind: "lineTo", x: 0, y: 0 }],
            },
          ],
        },
      }),
    ).toThrow("addShape: geometry.paths[0].commands must start with one moveTo command");
    expect(() =>
      addShape(source, requireHandle(source.slides[0].handle), {
        ...baseShape,
        width: asEmu(0),
        geometry: { kind: "preset", preset: "rect" },
      }),
    ).toThrow("addShape: width must be a finite positive EMU value");
    expect(() =>
      addShape(source, requireHandle(source.slides[0].handle), {
        ...baseShape,
        width: asEmu(0),
        height: asEmu(0),
        geometry: { kind: "preset", preset: "line" },
      }),
    ).toThrow("addShape: line width and height must not both be zero");
  });
});

describe("editing media and slide topology operations", () => {
  it("replaces image bytes after validating content type and records shared reference count", () => {
    const source = buildSourceModel();
    const imageHandle = requireHandle(imageByName(source, "Picture").handle);

    const edited = expectNonMutating(source, () => replaceImageBytes(source, imageHandle, PNG));

    expect(source.packageGraph.media[0].bytes).not.toBe(PNG);
    expect(edited.packageGraph.media[0]).toMatchObject({
      partPath: "ppt/media/image1.png",
      contentType: "image/png",
      bytes: PNG,
    });
    expect(edited.packageGraph.media[0].bytes).not.toBe(PNG);
    expect(edited.edits).toEqual([
      {
        kind: "replaceImage",
        handle: imageHandle,
        mediaPartPath: "ppt/media/image1.png",
        contentType: "image/png",
        sharedReferenceCount: 1,
      },
    ]);
    expect(() => replaceImageBytes(source, imageHandle, JPEG)).toThrow(
      "replaceImageBytes: replacement image content type 'image/jpeg' does not match existing media content type 'image/png'",
    );
  });

  it("adds an empty slide from a layout with reserved part, relationship, and numeric ids", () => {
    const source = buildSourceModel();

    const edited = expectNonMutating(source, () =>
      addEmptySlideFromLayout(source, {
        layoutPartPath: asPartPath("ppt/slideLayouts/slideLayout1.xml"),
      }),
    );

    expect(source.slides).toHaveLength(2);
    expect(edited.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
      "ppt/slides/slide3.xml",
    ]);
    expect(edited.slides[2]).toMatchObject({
      partPath: "ppt/slides/slide3.xml",
      layoutPartPath: "ppt/slideLayouts/slideLayout1.xml",
      shapes: [],
      handle: { partPath: "ppt/slides/slide3.xml" },
    });
    expect(edited.edits?.at(-1)).toEqual({
      kind: "addEmptySlideFromLayout",
      layoutPartPath: "ppt/slideLayouts/slideLayout1.xml",
      newSlidePartPath: "ppt/slides/slide3.xml",
      newRelationshipId: "rId3",
      newSlideNumericId: 301,
    });
  });

  it("duplicates a slide after the source slide and rewrites cloned handles to the new part", () => {
    const source = buildSourceModel();

    const edited = expectNonMutating(source, () =>
      duplicateSlide(source, requireHandle(source.slides[0].handle)),
    );

    expect(source.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(edited.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide3.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(edited.slides[1].partPath).toBe("ppt/slides/slide3.xml");
    expect(edited.slides[1].shapes[0]?.handle?.partPath).toBe("ppt/slides/slide3.xml");
    const clonedShape = edited.slides[1].shapes[0];
    if (clonedShape?.kind !== "shape") throw new Error("expected duplicated text shape");
    expect(clonedShape.textBody?.paragraphs[0]?.handle?.partPath).toBe("ppt/slides/slide3.xml");
    expect(clonedShape.textBody?.paragraphs[0]?.runs[0]?.handle?.partPath).toBe(
      "ppt/slides/slide3.xml",
    );
    expect(edited.edits?.at(-1)).toEqual({
      kind: "duplicateSlide",
      sourceSlidePartPath: "ppt/slides/slide1.xml",
      sourceRelationshipId: "rIdSlide1",
      newSlidePartPath: "ppt/slides/slide3.xml",
      newRelationshipId: "rId3",
      newSlideNumericId: 301,
    });
  });

  it("moves a slide to the requested final index without changing package parts", () => {
    const source = buildSourceModel();
    const edited = expectNonMutating(source, () =>
      moveSlide(source, requireHandle(source.slides[0].handle), { toIndex: 1 }),
    );

    expect(source.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(edited.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide1.xml",
    ]);
    expect(edited.slides.map((slide) => slide.partPath)).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide1.xml",
    ]);
    expect(edited.packageGraph.parts).toEqual(source.packageGraph.parts);
    expect(edited.edits?.at(-1)).toEqual({
      kind: "moveSlide",
      slidePartPath: "ppt/slides/slide1.xml",
      relationshipId: "rIdSlide1",
      toIndex: 1,
    });
  });

  it("treats moving a slide to its current index as a no-op", () => {
    const source = buildSourceModel();
    const edited = moveSlide(source, requireHandle(source.slides[1].handle), { toIndex: 1 });

    expect(edited).toBe(source);
    expect(edited.edits).toBeUndefined();
  });

  it("deletes a slide while dropping pending edits invalidated by that part", () => {
    const source = buildSourceModel();
    const dirty = replaceTextRunPlainText(source, requireHandle(firstRun(source).handle), "Dirty");

    const edited = expectNonMutating(dirty, () =>
      deleteSlide(dirty, requireHandle(dirty.slides[0].handle)),
    );

    expect(dirty.edits?.map((edit) => edit.kind)).toEqual(["replaceTextRunPlainText"]);
    expect(edited.presentation.slidePartPaths).toEqual(["ppt/slides/slide2.xml"]);
    expect(edited.edits).toEqual([
      {
        kind: "deleteSlide",
        slidePartPath: "ppt/slides/slide1.xml",
        relationshipId: "rIdSlide1",
      },
    ]);
    expect(edited.packageGraph.parts.map((part) => part.partPath)).not.toContain(
      "ppt/slides/slide1.xml",
    );
  });

  it("cancels addEmptySlideFromLayout followed by deleteSlide", () => {
    const source = buildSourceModel();
    const added = addEmptySlideFromLayout(source, {
      layoutPartPath: asPartPath("ppt/slideLayouts/slideLayout1.xml"),
    });
    const edited = expectNonMutating(added, () =>
      deleteSlide(added, requireHandle(added.slides[2].handle)),
    );

    expect(edited.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(edited.edits).toEqual([]);
    expect(edited.packageGraph.parts.map((part) => part.partPath)).not.toContain(
      "ppt/slides/slide3.xml",
    );
  });

  it("rejects invalid slide topology operations with operation-specific errors", () => {
    const source = buildSourceModel();
    const singleSlide = { ...source, slides: [source.slides[0]] };

    expect(() =>
      addEmptySlideFromLayout(source, {
        layoutPartPath: asPartPath("ppt/slideLayouts/missing.xml"),
      }),
    ).toThrow("addEmptySlideFromLayout: slide layout part path was not found");
    expect(() =>
      duplicateSlide(
        replaceTextRunPlainText(source, requireHandle(firstRun(source).handle), "Dirty"),
        requireHandle(source.slides[0].handle),
      ),
    ).toThrow("duplicateSlide: duplicating a slide with pending dirty part edits is unsupported");
    expect(() => deleteSlide(singleSlide, requireHandle(singleSlide.slides[0].handle))).toThrow(
      "deleteSlide: cannot delete the last slide",
    );
    expect(() =>
      moveSlide(source, { partPath: asPartPath("ppt/slides/missing.xml") }, { toIndex: 1 }),
    ).toThrow("moveSlide: slide handle was not found in PptxSourceModel source");
    expect(() => moveSlide(source, requireHandle(source.slides[0].handle), { toIndex: 2 })).toThrow(
      "moveSlide: toIndex must be an integer slide index in range",
    );
  });
});

function expectNonMutating(source: PptxSourceModel, edit: () => PptxSourceModel): PptxSourceModel {
  const before = structuredClone(source);
  const edited = edit();

  expect(source).toEqual(before);
  if (edited !== source) expect(edited).not.toBe(source);
  return edited;
}

function expectInvalidEdit(
  source: PptxSourceModel,
  edit: () => PptxSourceModel,
  expected: RegExp | string,
): void {
  const before = structuredClone(source);

  expect(edit).toThrow(expected);
  expect(source).toEqual(before);
}

function buildSourceModel(): PptxSourceModel {
  const presentationPath = asPartPath("ppt/presentation.xml");
  const slide1Path = asPartPath("ppt/slides/slide1.xml");
  const slide2Path = asPartPath("ppt/slides/slide2.xml");
  const layoutPath = asPartPath("ppt/slideLayouts/slideLayout1.xml");
  const imagePath = asPartPath("ppt/media/image1.png");

  return {
    packageGraph: {
      contentTypes: {
        defaults: [
          {
            extension: "rels",
            contentType: "application/vnd.openxmlformats-package.relationships+xml",
          },
          { extension: "png", contentType: "image/png" },
        ],
        overrides: [
          { partName: presentationPath, contentType: PRESENTATION_CONTENT_TYPE },
          { partName: slide1Path, contentType: SLIDE_CONTENT_TYPE },
          { partName: slide2Path, contentType: SLIDE_CONTENT_TYPE },
          { partName: layoutPath, contentType: SLIDE_LAYOUT_CONTENT_TYPE },
        ],
      },
      parts: [
        { partPath: presentationPath, contentType: PRESENTATION_CONTENT_TYPE },
        { partPath: slide1Path, contentType: SLIDE_CONTENT_TYPE },
        { partPath: slide2Path, contentType: SLIDE_CONTENT_TYPE },
        { partPath: layoutPath, contentType: SLIDE_LAYOUT_CONTENT_TYPE },
        { partPath: imagePath, contentType: "image/png" },
      ],
      relationships: [
        {
          sourcePartPath: presentationPath,
          relationships: [
            {
              id: asRelationshipId("rIdSlide1"),
              type: SLIDE_REL_TYPE,
              target: "slides/slide1.xml",
            },
            {
              id: asRelationshipId("rIdSlide2"),
              type: SLIDE_REL_TYPE,
              target: "slides/slide2.xml",
            },
          ],
        },
        {
          sourcePartPath: slide1Path,
          relationships: [
            {
              id: asRelationshipId("rIdLayout"),
              type: SLIDE_LAYOUT_REL_TYPE,
              target: "../slideLayouts/slideLayout1.xml",
            },
            {
              id: asRelationshipId("rIdImage"),
              type: IMAGE_REL_TYPE,
              target: "../media/image1.png",
            },
          ],
        },
        {
          sourcePartPath: slide2Path,
          relationships: [
            {
              id: asRelationshipId("rIdLayout"),
              type: SLIDE_LAYOUT_REL_TYPE,
              target: "../slideLayouts/slideLayout1.xml",
            },
          ],
        },
      ],
      media: [{ partPath: imagePath, contentType: "image/png", bytes: new Uint8Array(PNG) }],
      rawParts: [
        {
          kind: "binary",
          partPath: presentationPath,
          contentType: PRESENTATION_CONTENT_TYPE,
          bytes: xml(
            `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
              `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
              `<p:sldIdLst>` +
              `<p:sldId id="256" r:id="rIdSlide1"/>` +
              `<p:sldId id="300" r:id="rIdSlide2"/>` +
              `</p:sldIdLst>` +
              `</p:presentation>`,
          ),
        },
        {
          kind: "binary",
          partPath: slide1Path,
          contentType: SLIDE_CONTENT_TYPE,
          bytes: xml(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`),
        },
        {
          kind: "binary",
          partPath: slide2Path,
          contentType: SLIDE_CONTENT_TYPE,
          bytes: xml(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`),
        },
      ],
    },
    presentation: {
      partPath: presentationPath,
      slideSize: { width: asEmu(9144000), height: asEmu(5143500) },
      slidePartPaths: [slide1Path, slide2Path],
    },
    slides: [
      {
        partPath: slide1Path,
        layoutPartPath: layoutPath,
        shapes: [
          textShape(slide1Path, "10", "Title", 0),
          imageShape(slide1Path),
          textShape(slide1Path, "30", "Body", 2),
          slotTextShape(slide1Path, 3),
        ],
        handle: { partPath: slide1Path },
      },
      {
        partPath: slide2Path,
        layoutPartPath: layoutPath,
        shapes: [textShape(slide2Path, "40", "Second", 0)],
        handle: { partPath: slide2Path },
      },
    ],
    slideLayouts: [
      {
        partPath: layoutPath,
        masterPartPath: asPartPath("ppt/slideMasters/slideMaster1.xml"),
        type: "blank",
        shapes: [],
      },
    ],
    slideMasters: [],
    themes: [],
    diagnostics: [],
  };
}

function textShape(
  partPath: ReturnType<typeof asPartPath>,
  id: string,
  name: string,
  orderingSlot: number,
): SourceShape {
  const nodeId = asSourceNodeId(id);
  return {
    kind: "shape",
    nodeId,
    name,
    transform: {
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(3000),
      height: asEmu(1000),
    },
    geometry: { preset: "rect" },
    textBody: {
      paragraphs: [
        {
          handle: handle(partPath, `text:shape:${id}:p:0`, 0),
          runs: [
            {
              kind: "textRun",
              text: "Original",
              properties: { bold: true, italic: true, fontSize: asPt(24), typeface: "Calibri" },
              handle: handle(partPath, `text:shape:${id}:p:0:r:0`, 0),
            },
            {
              kind: "textRun",
              text: " Keep",
              handle: handle(partPath, `text:shape:${id}:p:0:r:1`, 1),
            },
          ],
        },
      ],
    },
    handle: handle(partPath, id, orderingSlot),
  };
}

function slotTextShape(partPath: ReturnType<typeof asPartPath>, orderingSlot: number): SourceShape {
  return {
    kind: "shape",
    name: "Slot Text",
    transform: {
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(3000),
      height: asEmu(1000),
    },
    geometry: { preset: "rect" },
    textBody: {
      paragraphs: [
        {
          handle: handle(partPath, `text:shapeSlot:${orderingSlot}:p:0`, 0),
          runs: [
            {
              kind: "textRun",
              text: "Slot Original",
              handle: handle(partPath, `text:shapeSlot:${orderingSlot}:p:0:r:0`, 0),
            },
          ],
        },
      ],
    },
    handle: { partPath, orderingSlot },
  };
}

function imageShape(partPath: ReturnType<typeof asPartPath>): SourceImage {
  return {
    kind: "image",
    nodeId: asSourceNodeId("20"),
    name: "Picture",
    blipRelationshipId: asRelationshipId("rIdImage"),
    handle: {
      partPath,
      nodeId: asSourceNodeId("20"),
      relationshipId: asRelationshipId("rIdImage"),
      orderingSlot: 1,
    },
  };
}

function handle(
  partPath: string | ReturnType<typeof asPartPath>,
  nodeId: string,
  orderingSlot?: number,
): SourceHandle {
  return {
    partPath: typeof partPath === "string" ? asPartPath(partPath) : partPath,
    nodeId: asSourceNodeId(nodeId),
    ...(orderingSlot === undefined ? {} : { orderingSlot }),
  };
}

function xml(content: string): Uint8Array {
  return new TextEncoder().encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`,
  );
}

function requireHandle(handle: SourceHandle | undefined): SourceHandle {
  if (handle === undefined) throw new Error("test fixture handle is missing");
  return handle;
}

function firstRun(source: PptxSourceModel) {
  return firstParagraph(source).runs[0];
}

function firstParagraph(source: PptxSourceModel) {
  const shape = shapeByName(source, "Title");
  const paragraph = shape.textBody?.paragraphs[0];
  if (paragraph === undefined) throw new Error("test fixture paragraph is missing");
  return paragraph;
}

function shapeByName(source: PptxSourceModel, name: string): SourceShape {
  for (const slide of source.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind === "shape" && shape.name === name) return shape;
    }
  }
  throw new Error(`test fixture shape '${name}' is missing`);
}

function imageByName(source: PptxSourceModel, name: string): SourceImage {
  for (const slide of source.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind === "image" && shape.name === name) return shape;
    }
  }
  throw new Error(`test fixture image '${name}' is missing`);
}

function connectorByName(source: PptxSourceModel, name: string): SourceConnector {
  for (const slide of source.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind === "connector" && shape.name === name) return shape;
    }
  }
  throw new Error(`test fixture connector '${name}' is missing`);
}
