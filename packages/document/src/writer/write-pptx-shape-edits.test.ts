import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

// Import via the actual public surface (`@pptx-glimpse/document`).
import {
  addChart,
  addConnector,
  addPicture,
  addShape,
  addTable,
  addTextBox,
  asEmu,
  asSourceNodeId,
  createPptx,
  deleteShape,
  findShapeNodeBySourceHandle,
  groupShapes,
  readPptx,
  replaceImageBytes,
  replaceTextRunPlainText,
  setShapeFill,
  setShapeOutline,
  updateShapeTransform,
  writePptx,
} from "../index.js";
import { resolveInternalRelationshipTarget } from "../source/package-paths.js";
import {
  BLUE_PNG,
  buildConnectedShapeFixture,
  buildMediaReplacementFixture,
  buildShapeDeleteFixture,
  buildShapeStyleFixture,
  buildTextEditFixture,
  buildTextEditFixtureFromSlide,
  decoder,
  encoder,
  findConnectorByName,
  findConnectorByNameOptional,
  findShapeByName,
  firstShape,
  getEntry,
  GREEN_PNG,
  RED_PNG,
  requireHandle,
  requireShape,
  xml,
} from "./write-pptx.test-helpers.js";

function buildAuthoredDrawingDeleteSource(): ReturnType<typeof readPptx> {
  let source = createPptx();
  const slideHandle = source.slides[0]?.handle;
  if (slideHandle === undefined) throw new Error("authored drawing fixture slide is missing");
  source = addShape(source, slideHandle, {
    geometry: { kind: "preset", preset: "rect" },
    offsetX: asEmu(100),
    offsetY: asEmu(100),
    width: asEmu(1000),
    height: asEmu(1000),
    name: "Keep Before",
  });
  source = addTable(source, slideHandle, {
    offsetX: asEmu(1200),
    offsetY: asEmu(100),
    width: asEmu(2000),
    height: asEmu(1000),
    columnWidths: [asEmu(2000)],
    rows: [{ height: asEmu(1000), cells: [{ text: "Table" }] }],
    name: "Delete Table",
  });
  source = addChart(source, slideHandle, {
    chartType: "bar",
    series: [{ name: "Series", categories: ["A"], values: [1] }],
    offsetX: asEmu(3400),
    offsetY: asEmu(100),
    width: asEmu(2000),
    height: asEmu(1000),
    name: "Delete Chart",
  });
  source = addShape(source, slideHandle, {
    geometry: { kind: "preset", preset: "rect" },
    offsetX: asEmu(5600),
    offsetY: asEmu(100),
    width: asEmu(1000),
    height: asEmu(1000),
    name: "Group Child A",
  });
  source = addShape(source, slideHandle, {
    geometry: { kind: "preset", preset: "rect" },
    offsetX: asEmu(6800),
    offsetY: asEmu(100),
    width: asEmu(1000),
    height: asEmu(1000),
    name: "Group Child B",
  });
  const children = source.slides[0].shapes.filter(
    (shape) => shape.kind !== "raw" && shape.name?.startsWith("Group Child"),
  );
  source = groupShapes(
    source,
    children.map((shape) => requireHandle(shape.handle)),
  );
  source = addShape(source, slideHandle, {
    geometry: { kind: "preset", preset: "rect" },
    offsetX: asEmu(8000),
    offsetY: asEmu(100),
    width: asEmu(1000),
    height: asEmu(1000),
    name: "Keep After",
  });
  const archive = unzipSync(writePptx(source));
  return readPptx(
    zipSync({
      ...archive,
      "docProps/custom.xml": xml(`<Properties><custom value="unrelated-orphan"/></Properties>`),
    }),
  );
}

describe("writePptx - shape add/delete edits", () => {
  it("adds a text box with a collision-free shape id and persists it", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const edited = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(914400),
      offsetY: asEmu(457200),
      width: asEmu(2743200),
      height: asEmu(914400),
      text: "Added text box",
    });
    const output = writePptx(edited);
    const reread = readPptx(output);
    const added = requireShape(findShapeByName(reread, "TextBox 31"));
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(added).toMatchObject({
      nodeId: "31",
      name: "TextBox 31",
      transform: {
        offsetX: 914400,
        offsetY: 457200,
        width: 2743200,
        height: 914400,
      },
    });
    expect(added.textBody?.paragraphs[0]?.runs[0]?.text).toBe("Added text box");
    expect(slideXml).toContain(`<p:cNvPr id="31" name="TextBox 31"`);
    expect(slideXml).toContain(`<p:cNvSpPr txBox="1"`);
    expect(decoder.decode(getEntry(output, "docProps/custom.xml"))).toContain("preserve-me");
  });

  it("adds a connector with connection sites, preset geometry, and arrow endpoints", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const start = findShapeByName(source, "Delete Me");
    const end = findShapeByName(source, "Keep Shape");
    const edited = addConnector(source, source.slides[0].handle!, {
      preset: "bentConnector3",
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(700),
      height: asEmu(800),
      flipHorizontal: true,
      flipVertical: true,
      start: {
        shapeHandle: requireHandle(start.handle),
        connectionSiteIndex: 1,
      },
      end: {
        shapeHandle: requireHandle(end.handle),
        connectionSiteIndex: 3,
      },
      outline: {
        width: asEmu(12700),
        fill: { kind: "solid", color: { kind: "srgb", hex: "00AAFF" } },
        dash: "lgDashDotDot",
        headEnd: { type: "oval", width: "sm", length: "sm" },
        tailEnd: { type: "triangle", width: "med", length: "lg" },
      },
    });
    const output = writePptx(edited);
    const reread = readPptx(output);
    const added = findConnectorByName(reread, "Connector 31");
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(added).toMatchObject({
      nodeId: "31",
      name: "Connector 31",
      connection: {
        start: { shapeId: "10", connectionSiteIndex: 1 },
        end: { shapeId: "30", connectionSiteIndex: 3 },
      },
      geometry: { preset: "bentConnector3" },
      transform: { flipHorizontal: true, flipVertical: true },
      outline: {
        width: 12700,
        fill: { kind: "solid", color: { kind: "srgb", hex: "00AAFF" } },
        dashStyle: "lgDashDotDot",
        headEnd: { type: "oval", width: "sm", length: "sm" },
        tailEnd: { type: "triangle", width: "med", length: "lg" },
      },
    });
    expect(findConnectorByName(edited, "Connector 31").outline).toMatchObject({
      width: 12700,
      fill: { kind: "solid", color: { kind: "srgb", hex: "00AAFF" } },
      dashStyle: "lgDashDotDot",
      headEnd: { type: "oval", width: "sm", length: "sm" },
      tailEnd: { type: "triangle", width: "med", length: "lg" },
    });
    expect(slideXml).toContain(`<p:cxnSp>`);
    expect(slideXml).toContain(`<a:stCxn id="10" idx="1"`);
    expect(slideXml).toContain(`<a:endCxn id="30" idx="3"`);
    expect(slideXml).toContain(`<a:prstGeom prst="bentConnector3"`);
    expect(slideXml).toContain(`<a:xfrm flipH="1" flipV="1"`);
    expect(slideXml).toContain(`<a:ln w="12700"`);
    expect(slideXml).toContain(`<a:srgbClr val="00AAFF"`);
    expect(slideXml).toContain(`<a:prstDash val="lgDashDotDot"`);
    expect(slideXml).toContain(`<a:headEnd type="oval" w="sm" len="sm"`);
    expect(slideXml).toContain(`<a:tailEnd type="triangle" w="med" len="lg"`);
    expect(decoder.decode(getEntry(output, "docProps/custom.xml"))).toContain("preserve-me");
  });

  it("adds and deletes a free connector without native connection sites", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const edited = addConnector(source, source.slides[0].handle!, {
      preset: "straightConnector1",
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(700),
      height: asEmu(800),
      outline: {
        tailEnd: { type: "triangle", width: "med", length: "med" },
      },
    });
    const output = writePptx(edited);
    const reread = readPptx(output);
    const added = findConnectorByName(reread, "Connector 31");
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(added).toMatchObject({
      nodeId: "31",
      name: "Connector 31",
      geometry: { preset: "straightConnector1" },
      outline: { tailEnd: { type: "triangle", width: "med", length: "med" } },
    });
    expect(added.connection).toBeUndefined();
    expect(slideXml).toContain(`<p:cxnSp>`);
    expect(slideXml).not.toContain(`<a:stCxn`);
    expect(slideXml).not.toContain(`<a:endCxn`);

    const persisted = readPptx(output);
    const deleted = deleteShape(
      persisted,
      requireHandle(findConnectorByName(persisted, "Connector 31").handle),
    );
    const deletedOutput = writePptx(deleted);
    expect(findConnectorByNameOptional(readPptx(deletedOutput), "Connector 31")).toBeUndefined();
  });

  it("rejects deleting a shape referenced by a connector", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const start = findShapeByName(source, "Delete Me");
    const end = findShapeByName(source, "Keep Shape");
    const withConnector = addConnector(source, source.slides[0].handle!, {
      preset: "straightConnector1",
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(700),
      height: asEmu(800),
      start: {
        shapeHandle: requireHandle(start.handle),
        connectionSiteIndex: 1,
      },
      end: {
        shapeHandle: requireHandle(end.handle),
        connectionSiteIndex: 3,
      },
    });

    expect(() => deleteShape(withConnector, requireHandle(start.handle))).toThrow(
      /referenced by connector/,
    );
  });

  it("allows deleting a connector before its connected target in one edit journal", () => {
    const source = readPptx(buildConnectedShapeFixture());
    const connector = findConnectorByName(source, "Connected Shapes");
    const withoutConnector = deleteShape(source, requireHandle(connector.handle));
    const target = findShapeByName(withoutConnector, "Delete Me");
    const deleted = deleteShape(withoutConnector, requireHandle(target.handle));
    const reread = readPptx(writePptx(deleted));

    expect(findConnectorByNameOptional(reread, "Connected Shapes")).toBeUndefined();
    expect(reread.slides[0].shapes.map((shape) => shape.nodeId)).not.toContain("10");
  });

  it("allows an added text box to be edited before write", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const withTextBox = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(1000),
      offsetY: asEmu(2000),
      width: asEmu(3000),
      height: asEmu(4000),
      text: "Initial",
      name: "Editable Added",
    });
    const added = requireShape(findShapeByName(withTextBox, "Editable Added"));
    const runHandle = added.textBody?.paragraphs[0]?.runs[0]?.handle;
    if (runHandle === undefined || added.handle === undefined) {
      throw new Error("added text box handles not found");
    }

    const edited = updateShapeTransform(
      replaceTextRunPlainText(withTextBox, runHandle, "Edited Added"),
      added.handle,
      {
        offsetX: asEmu(5000),
        offsetY: asEmu(6000),
        width: asEmu(7000),
        height: asEmu(8000),
      },
    );
    const rereadAdded = requireShape(
      findShapeByName(readPptx(writePptx(edited)), "Editable Added"),
    );

    expect(rereadAdded.textBody?.paragraphs[0]?.runs[0]?.text).toBe("Edited Added");
    expect(rereadAdded.transform).toMatchObject({
      offsetX: 5000,
      offsetY: 6000,
      width: 7000,
      height: 8000,
    });
  });

  it("does not reuse a pending-deleted shape id when adding a text box", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const deletedMaxIdShape = deleteShape(
      source,
      requireHandle(findShapeByName(source, "Keep Shape").handle),
    );
    const edited = addTextBox(deletedMaxIdShape, deletedMaxIdShape.slides[0].handle!, {
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(1100),
      height: asEmu(1200),
      text: "Added after delete",
    });
    const output = writePptx(edited);
    const reread = readPptx(output);

    expect(findShapeByName(reread, "TextBox 31").textBody?.paragraphs[0]?.runs[0]?.text).toBe(
      "Added after delete",
    );
    expect(() => findShapeByName(reread, "Keep Shape")).toThrow(/shape not found/);
  });

  it("does not reuse a pending-deleted shape id when adding a picture", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const deletedMaxIdShape = deleteShape(
      source,
      requireHandle(findShapeByName(source, "Keep Shape").handle),
    );
    const edited = addPicture(deletedMaxIdShape, deletedMaxIdShape.slides[0].handle!, {
      bytes: RED_PNG,
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(1100),
      height: asEmu(1200),
    });
    const output = writePptx(edited);
    const reread = readPptx(output);
    const picture = reread.slides[0]?.shapes.find(
      (shape) => shape.kind === "image" && shape.name === "Picture 31",
    );

    expect(picture).toMatchObject({
      kind: "image",
      nodeId: "31",
      name: "Picture 31",
      blipRelationshipId: "rId1",
    });
    expect(() => findShapeByName(reread, "Keep Shape")).toThrow(/shape not found/);
  });

  it("cancels the add edit when a newly-added text box is deleted before write", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const withTextBox = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(300),
      height: asEmu(400),
      text: "Temporary",
      name: "Temporary TextBox",
    });
    const added = findShapeByName(withTextBox, "Temporary TextBox");
    const edited = deleteShape(withTextBox, requireHandle(added.handle));
    const output = writePptx(edited);

    expect(
      edited.edits?.filter((edit) => edit.kind === "addTextBox" || edit.kind === "deleteShape"),
    ).toEqual([]);
    expect(decoder.decode(getEntry(output, "ppt/slides/slide1.xml"))).not.toContain(
      "Temporary TextBox",
    );
  });

  it("finalizes added shape XML on the edit record and the writer only splices it", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const start = findShapeByName(source, "Delete Me");
    const end = findShapeByName(source, "Keep Shape");
    const withTextBox = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(914400),
      offsetY: asEmu(457200),
      width: asEmu(2743200),
      height: asEmu(914400),
      text: "Added text box",
    });
    const edited = addConnector(withTextBox, withTextBox.slides[0].handle!, {
      preset: "straightConnector1",
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(700),
      height: asEmu(800),
      start: { shapeHandle: requireHandle(start.handle), connectionSiteIndex: 1 },
      end: { shapeHandle: requireHandle(end.handle), connectionSiteIndex: 3 },
    });
    const textBoxEdit = edited.edits?.find((edit) => edit.kind === "addTextBox");
    const connectorEdit = edited.edits?.find((edit) => edit.kind === "addConnector");
    if (textBoxEdit?.kind !== "addTextBox" || connectorEdit?.kind !== "addConnector") {
      throw new Error("expected addTextBox and addConnector edits to be recorded");
    }
    const slideXml = decoder.decode(getEntry(writePptx(edited), "ppt/slides/slide1.xml"));

    expect(textBoxEdit.xml).toContain(`<p:cNvPr id="31" name="TextBox 31"/>`);
    expect(textBoxEdit.xml).toContain(`<a:t>Added text box</a:t>`);
    expect(connectorEdit.xml).toContain(`<a:stCxn id="10" idx="1"/>`);
    expect(slideXml).toContain(textBoxEdit.xml);
    expect(slideXml).toContain(connectorEdit.xml);
  });

  it("round-trips added text box text that needs XML escaping and space preservation", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const text = ` A & B <C> "quoted" `;
    const edited = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(914400),
      offsetY: asEmu(457200),
      width: asEmu(2743200),
      height: asEmu(914400),
      text,
      name: "Escaped TextBox",
    });
    const output = writePptx(edited);
    const reread = readPptx(output);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(
      requireShape(findShapeByName(edited, "Escaped TextBox")).textBody?.paragraphs[0]?.runs[0]
        ?.text,
    ).toBe(text);
    expect(
      requireShape(findShapeByName(reread, "Escaped TextBox")).textBody?.paragraphs[0]?.runs[0]
        ?.text,
    ).toBe(text);
    expect(slideXml).toContain(`xml:space="preserve"`);
    expect(slideXml).toContain(`A &amp; B &lt;C&gt;`);
  });

  it("deletes only the targeted sp shape while preserving other shapes and invisible slide material", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const deleted = deleteShape(source, requireHandle(source.slides[0].shapes[0]?.handle));
    const output = writePptx(deleted);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);
    const rereadShapeNames: (string | undefined)[] = [];
    for (const shape of reread.slides[0].shapes) {
      rereadShapeNames.push(shape.kind === "raw" ? undefined : shape.name);
    }

    expect(rereadShapeNames).toEqual(expect.arrayContaining(["Keep Picture", "Keep Shape"]));
    expect(reread.slides[0].shapes).toHaveLength(2);
    expect(slideXml).not.toContain("Delete Me");
    expect(slideXml).toContain("Keep Picture");
    expect(slideXml).toContain("Keep Shape");
    expect(slideXml).toContain("<p:timing>");
    expect(decoder.decode(getEntry(output, "docProps/custom.xml"))).toContain("preserve-me");
  });

  it("deletes a top-level picture and cleans its unshared relationship and media part", () => {
    const source = readPptx(buildMediaReplacementFixture());
    const image = source.slides[0].shapes.find((shape) => shape.name === "Replace Target");
    const deleted = deleteShape(source, requireHandle(image?.handle));
    const output = writePptx(deleted);
    const reread = readPptx(output);
    const archive = unzipSync(output);
    const slideRelationships = reread.packageGraph.relationships.find(
      (group) => group.sourcePartPath === "ppt/slides/slide1.xml",
    );

    expect(reread.slides[0].shapes).toEqual([]);
    expect(slideRelationships?.relationships.map((relationship) => relationship.id)).not.toContain(
      "rIdImage1",
    );
    expect(archive["ppt/media/image1.png"]).toBeUndefined();
    expect(slideRelationships?.relationships.map((relationship) => relationship.id)).toContain(
      "rIdImage2",
    );
    expect(archive["ppt/media/image2.png"]).toEqual(GREEN_PNG);
    expect(archive["docProps/custom.xml"]).toBeDefined();
  });

  it("removes the last picture media default only after the final extension user is deleted", () => {
    let authored = createPptx();
    for (const [index, bytes] of [RED_PNG, GREEN_PNG].entries()) {
      authored = addPicture(authored, authored.slides[0].handle!, {
        bytes,
        offsetX: asEmu(100 + index * 1200),
        offsetY: asEmu(100),
        width: asEmu(1000),
        height: asEmu(1000),
        name: `Picture ${index + 1}`,
      });
    }
    const persisted = readPptx(writePptx(authored));
    const first = persisted.slides[0].shapes.find((shape) => shape.name === "Picture 1");
    const afterFirst = deleteShape(persisted, requireHandle(first?.handle));
    const firstReread = readPptx(writePptx(afterFirst));

    expect(firstReread.packageGraph.media).toHaveLength(1);
    expect(
      firstReread.packageGraph.contentTypes.defaults.map((entry) => entry.extension),
    ).toContain("png");

    const second = afterFirst.slides[0].shapes.find((shape) => shape.name === "Picture 2");
    const reread = readPptx(writePptx(deleteShape(afterFirst, requireHandle(second?.handle))));

    expect(reread.packageGraph.media).toEqual([]);
    expect(reread.packageGraph.contentTypes.defaults.map((entry) => entry.extension)).not.toContain(
      "png",
    );
  });

  it("cleans the old shared image after the surviving picture has a pending copy-on-write replacement", () => {
    const source = readPptx(buildMediaReplacementFixture(true));
    const pictureA = source.slides[0].shapes.find((shape) => shape.name === "Replace Target");
    const pictureB = source.slides[0].shapes.find((shape) => shape.name === "Keep Shared");
    const replaced = replaceImageBytes(source, requireHandle(pictureB?.handle), BLUE_PNG);
    const deleted = deleteShape(replaced, requireHandle(pictureA?.handle));
    const output = writePptx(deleted);
    const reread = readPptx(output);
    const archive = unzipSync(output);
    const remainingPicture = reread.slides[0].shapes.find((shape) => shape.kind === "image");
    const relationshipIds = reread.packageGraph.relationships
      .find((group) => group.sourcePartPath === "ppt/slides/slide1.xml")
      ?.relationships.map((relationship) => relationship.id);

    expect(remainingPicture?.name).toBe("Keep Shared");
    expect(remainingPicture?.blipRelationshipId).toBe("rId3");
    expect(relationshipIds).not.toContain("rIdImage1");
    expect(relationshipIds).toContain("rId3");
    expect(archive["ppt/media/image1.png"]).toBeUndefined();
    expect(archive["ppt/media/image3.png"]).toEqual(BLUE_PNG);
  });

  it("cancels a newly added picture and its package resources before the first write", () => {
    let source = createPptx();
    source = addPicture(source, source.slides[0].handle!, {
      bytes: RED_PNG,
      offsetX: asEmu(100),
      offsetY: asEmu(100),
      width: asEmu(1000),
      height: asEmu(1000),
    });
    const picture = source.slides[0].shapes.find((shape) => shape.kind === "image");
    const deleted = deleteShape(source, requireHandle(picture?.handle));

    expect(deleted.edits).toEqual([]);
    expect(deleted.packageGraph.media).toEqual([]);
    expect(readPptx(writePptx(deleted)).slides[0].shapes).toEqual([]);
  });

  it("cancels a newly added table and its external hyperlink relationship", () => {
    let source = createPptx();
    source = addTable(source, source.slides[0].handle!, {
      offsetX: asEmu(100),
      offsetY: asEmu(100),
      width: asEmu(1000),
      height: asEmu(1000),
      columnWidths: [asEmu(1000)],
      rows: [
        {
          height: asEmu(1000),
          cells: [{ runs: [{ text: "Link", hyperlink: "https://example.com" }] }],
        },
      ],
    });
    const table = source.slides[0].shapes.find((shape) => shape.kind === "table");
    const deleted = deleteShape(source, requireHandle(table?.handle));
    const slideRelationships = deleted.packageGraph.relationships.find(
      (group) => group.sourcePartPath === deleted.slides[0].partPath,
    );

    expect(
      slideRelationships?.relationships.filter((relationship) =>
        relationship.type.endsWith("/hyperlink"),
      ),
    ).toEqual([]);
    expect(readPptx(writePptx(deleted)).slides[0].shapes).toEqual([]);
  });

  it("deletes a group created earlier in the same edit journal", () => {
    let source = createPptx();
    const slideHandle = source.slides[0].handle!;
    for (const offsetX of [100, 1200]) {
      source = addShape(source, slideHandle, {
        geometry: { kind: "preset", preset: "rect" },
        offsetX: asEmu(offsetX),
        offsetY: asEmu(100),
        width: asEmu(1000),
        height: asEmu(1000),
      });
    }
    source = groupShapes(
      source,
      source.slides[0].shapes.map((shape) => requireHandle(shape.handle)),
    );
    const group = source.slides[0].shapes.find((shape) => shape.kind === "group");
    const deleted = deleteShape(source, requireHandle(group?.handle));

    expect(readPptx(writePptx(deleted)).slides[0].shapes).toEqual([]);
    expect(deleted.edits?.map((edit) => edit.kind)).toEqual([
      "addShape",
      "addShape",
      "groupShapes",
      "deleteShape",
    ]);
  });

  it("keeps media referenced by another picture, image fill, or raw VML node", () => {
    for (const shared of [true, "fill", "vml"] as const) {
      const source = readPptx(buildMediaReplacementFixture(shared));
      const target = source.slides[0].shapes.find((shape) => shape.name === "Replace Target");
      const output = writePptx(deleteShape(source, requireHandle(target?.handle)));
      const reread = readPptx(output);

      expect(reread.packageGraph.media.map((media) => media.partPath)).toContain(
        "ppt/media/image1.png",
      );
      expect(
        reread.packageGraph.relationships
          .find((group) => group.sourcePartPath === "ppt/slides/slide1.xml")
          ?.relationships.map((relationship) => relationship.id),
      ).toContain("rIdImage1");
    }
  });

  it("removes a shared image relationship and media after sequentially deleting all pictures", () => {
    const source = readPptx(buildMediaReplacementFixture(true));
    const first = source.slides[0].shapes.find((shape) => shape.name === "Replace Target");
    const afterFirst = deleteShape(source, requireHandle(first?.handle));
    const second = afterFirst.slides[0].shapes.find((shape) => shape.name === "Keep Shared");
    const deleted = deleteShape(afterFirst, requireHandle(second?.handle));
    const reread = readPptx(writePptx(deleted));

    expect(reread.packageGraph.media.map((media) => media.partPath)).not.toContain(
      "ppt/media/image1.png",
    );
    expect(
      reread.packageGraph.relationships
        .find((group) => group.sourcePartPath === "ppt/slides/slide1.xml")
        ?.relationships.map((relationship) => relationship.id),
    ).not.toContain("rIdImage1");
  });

  it("cleans relationships inside every descendant of a pending nested group delete", () => {
    let source = readPptx(buildMediaReplacementFixture(true));
    source = groupShapes(
      source,
      source.slides[0].shapes.map((shape) => requireHandle(shape.handle)),
    );
    source = addShape(source, source.slides[0].handle!, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(100),
      offsetY: asEmu(100),
      width: asEmu(1000),
      height: asEmu(1000),
    });
    source = groupShapes(
      source,
      source.slides[0].shapes.map((shape) => requireHandle(shape.handle)),
    );
    const outerGroup = source.slides[0].shapes.find((shape) => shape.kind === "group");
    const deleted = deleteShape(source, requireHandle(outerGroup?.handle));
    const reread = readPptx(writePptx(deleted));

    expect(reread.slides[0].shapes).toEqual([]);
    expect(reread.packageGraph.media.map((media) => media.partPath)).not.toContain(
      "ppt/media/image1.png",
    );
  });

  it("deletes native table, chart, and group drawings while preserving sibling order", () => {
    const persisted = buildAuthoredDrawingDeleteSource();
    const targets = [
      persisted.slides[0].shapes.find(
        (shape) => shape.kind === "table" && shape.name === "Delete Table",
      ),
      persisted.slides[0].shapes.find(
        (shape) => shape.kind === "chart" && shape.name === "Delete Chart",
      ),
      persisted.slides[0].shapes.find((shape) => shape.kind === "group"),
    ];
    for (const target of targets) {
      const beforeNames = persisted.slides[0].shapes
        .filter((shape) => shape !== target)
        .map((shape) => (shape.kind === "raw" ? undefined : shape.name));
      const reread = readPptx(writePptx(deleteShape(persisted, requireHandle(target?.handle))));

      expect(
        reread.slides[0].shapes.map((shape) => (shape.kind === "raw" ? undefined : shape.name)),
      ).toEqual(beforeNames);
      expect(reread.slides[0].shapes.map((shape) => shape.nodeId)).toEqual(
        persisted.slides[0].shapes.filter((shape) => shape !== target).map((shape) => shape.nodeId),
      );
    }
  });

  it("recursively removes an unshared chart and embedded workbook but keeps an orphan", () => {
    const persisted = buildAuthoredDrawingDeleteSource();
    const chart = persisted.slides[0].shapes.find((shape) => shape.name === "Delete Chart");
    const chartEdit = persisted.edits?.find(
      (edit) => edit.kind === "addChart" && edit.shapeId === String(chart?.nodeId),
    );
    expect(chartEdit).toBeUndefined();
    const chartRelationship = persisted.packageGraph.relationships
      .find((group) => group.sourcePartPath === persisted.slides[0].partPath)
      ?.relationships.find((relationship) => relationship.id === chart?.handle?.relationshipId);
    const chartPath =
      chartRelationship === undefined
        ? undefined
        : resolveInternalRelationshipTarget(persisted.slides[0].partPath, chartRelationship);
    const workbookRelationship = persisted.packageGraph.relationships
      .find((group) => group.sourcePartPath === chartPath)
      ?.relationships.find((relationship) => relationship.type.endsWith("/package"));
    const workbookPath =
      chartPath === undefined || workbookRelationship === undefined
        ? undefined
        : resolveInternalRelationshipTarget(chartPath, workbookRelationship);
    if (chartPath === undefined || workbookPath === undefined) {
      throw new Error("authored chart cleanup paths were not found");
    }

    const output = writePptx(deleteShape(persisted, requireHandle(chart?.handle)));
    const archive = unzipSync(output);
    expect(archive[chartPath]).toBeUndefined();
    expect(archive[workbookPath]).toBeUndefined();
    expect(archive["docProps/custom.xml"]).toBeDefined();
    const contentTypes = decoder.decode(archive["[Content_Types].xml"]);
    expect(contentTypes).not.toContain(`PartName="/${chartPath}"`);
    expect(contentTypes).not.toContain(`PartName="/${workbookPath}"`);
  });

  it("keeps a shared chart and workbook when another frame retains the owner relationship", () => {
    const persisted = buildAuthoredDrawingDeleteSource();
    const archive = unzipSync(writePptx(persisted));
    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = decoder.decode(archive[slidePath]);
    const chartXml = /<p:graphicFrame>.*?Delete Chart.*?<\/p:graphicFrame>/.exec(slideXml)?.[0];
    if (chartXml === undefined) throw new Error("authored chart XML was not found");
    const sharedChartXml = chartXml
      .replace(/<p:cNvPr id="\d+" name="Delete Chart"/, '<p:cNvPr id="98" name="Delete Chart"')
      .replace("Delete Chart", "Keep Shared Chart");
    const source = readPptx(
      zipSync({
        ...archive,
        [slidePath]: encoder.encode(
          slideXml.replace("</p:spTree>", `${sharedChartXml}</p:spTree>`),
        ),
      }),
    );
    const target = source.slides[0].shapes.find(
      (shape) => shape.kind === "chart" && shape.name === "Delete Chart",
    );
    const output = writePptx(deleteShape(source, requireHandle(target?.handle)));
    const reread = readPptx(output);

    expect(
      reread.slides[0].shapes.map((shape) => (shape.kind === "raw" ? undefined : shape.name)),
    ).toContain("Keep Shared Chart");
    expect(reread.packageGraph.parts.some((part) => part.partPath.startsWith("ppt/charts/"))).toBe(
      true,
    );
    expect(
      reread.packageGraph.parts.some((part) => part.partPath.startsWith("ppt/embeddings/")),
    ).toBe(true);
  });

  it("removes a shared chart and workbook after sequentially deleting all frames", () => {
    const persisted = buildAuthoredDrawingDeleteSource();
    const archive = unzipSync(writePptx(persisted));
    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = decoder.decode(archive[slidePath]);
    const chartXml = /<p:graphicFrame>.*?Delete Chart.*?<\/p:graphicFrame>/.exec(slideXml)?.[0];
    if (chartXml === undefined) throw new Error("authored chart XML was not found");
    const sharedChartXml = chartXml
      .replace(/<p:cNvPr id="\d+" name="Delete Chart"/, '<p:cNvPr id="98" name="Delete Chart"')
      .replace("Delete Chart", "Keep Shared Chart");
    const source = readPptx(
      zipSync({
        ...archive,
        [slidePath]: encoder.encode(
          slideXml.replace("</p:spTree>", `${sharedChartXml}</p:spTree>`),
        ),
      }),
    );
    const first = source.slides[0].shapes.find((shape) => shape.name === "Delete Chart");
    const afterFirst = deleteShape(source, requireHandle(first?.handle));
    const second = afterFirst.slides[0].shapes.find((shape) => shape.name === "Keep Shared Chart");
    const reread = readPptx(writePptx(deleteShape(afterFirst, requireHandle(second?.handle))));

    expect(reread.packageGraph.parts.some((part) => part.partPath.startsWith("ppt/charts/"))).toBe(
      false,
    );
    expect(
      reread.packageGraph.parts.some((part) => part.partPath.startsWith("ppt/embeddings/")),
    ).toBe(false);
  });

  it("atomically rejects SmartArt, unknown graphicFrame, and AlternateContent targets", () => {
    const archive = unzipSync(buildShapeDeleteFixture());
    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = decoder.decode(getEntry(buildShapeDeleteFixture(), slidePath));
    const unsupported =
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="40" name="Unknown Frame"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="urn:unknown"><x:payload xmlns:x="urn:unknown"/></a:graphicData></a:graphic></p:graphicFrame>` +
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="41" name="SmartArt"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:dm="rIdDiagram"/></a:graphicData></a:graphic></p:graphicFrame>` +
      `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="p14"><p:sp><p:nvSpPr><p:cNvPr id="42" name="Alternate Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:prstGeom prst="rect"/></p:spPr></p:sp></mc:Choice></mc:AlternateContent>`;
    const source = readPptx(
      zipSync({
        ...archive,
        [slidePath]: encoder.encode(slideXml.replace("</p:spTree>", `${unsupported}</p:spTree>`)),
      }),
    );

    for (const [nodeId, alternate] of [
      ["40", false],
      ["41", false],
      ["42", true],
    ] as const) {
      const target = source.slides[0].shapes.find((shape) => shape.nodeId === nodeId);
      const before = structuredClone(source);
      expect(() => deleteShape(source, requireHandle(target?.handle))).toThrow(
        alternate
          ? /AlternateContent/
          : /only top-level sp, cxnSp, pic, native table\/chart graphicFrame, or grpSp/,
      );
      expect(source).toEqual(before);
    }
  });

  it("atomically rejects deleting a group whose descendant is referenced externally", () => {
    const persisted = buildAuthoredDrawingDeleteSource();
    const group = persisted.slides[0].shapes.find((shape) => shape.kind === "group");
    if (group?.kind !== "group") throw new Error("group fixture was not found");
    const descendantId = group.children[0]?.nodeId;
    const archive = unzipSync(writePptx(persisted));
    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = decoder.decode(archive[slidePath]);
    const connector = `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="99" name="External Connector"/><p:cNvCxnSpPr><a:stCxn id="${String(descendantId)}" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:prstGeom prst="straightConnector1"/></p:spPr></p:cxnSp>`;
    const source = readPptx(
      zipSync({
        ...archive,
        [slidePath]: encoder.encode(slideXml.replace("</p:spTree>", `${connector}</p:spTree>`)),
      }),
    );
    const sourceGroup = source.slides[0].shapes.find((shape) => shape.kind === "group");
    const before = structuredClone(source);

    expect(() => deleteShape(source, requireHandle(sourceGroup?.handle))).toThrow(
      /referenced by connector 'External Connector'/,
    );
    expect(source).toEqual(before);
  });

  it("rejects conflicting shape additions for the same shape id", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const withTextBox = addTextBox(source, source.slides[0].handle!, {
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(300),
      height: asEmu(400),
      text: "Added once",
    });
    const additions = withTextBox.edits?.filter((edit) => edit.kind === "addTextBox") ?? [];
    const conflicted = { ...withTextBox, edits: [...(withTextBox.edits ?? []), ...additions] };

    expect(() => writePptx(conflicted)).toThrow(/conflicting shape additions/);
  });

  it("rejects conflicting shape delete edits for the same handle", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const deleted = deleteShape(source, requireHandle(source.slides[0].shapes[0]?.handle));
    const deletes = deleted.edits?.filter((edit) => edit.kind === "deleteShape") ?? [];
    const conflicted = { ...deleted, edits: [...(deleted.edits ?? []), ...deletes] };

    expect(() => writePptx(conflicted)).toThrow(/conflicting shape delete edits/);
  });
});

describe("writePptx - shape fill and outline editing", () => {
  it("Writes solid shape fill and outline color or width edits.", () => {
    const source = readPptx(buildShapeStyleFixture());
    const shape = findShapeByName(source, "Styled");
    const edited = setShapeOutline(
      setShapeFill(source, shape.handle!, {
        kind: "solid",
        color: { kind: "srgb", hex: "00aa44" },
      }),
      shape.handle!,
      {
        width: asEmu(25400),
        fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } },
      },
    );
    const output = writePptx(edited);
    const reread = readPptx(output);
    const rereadShape = findShapeByName(reread, "Styled");
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(rereadShape.fill).toEqual({
      kind: "solid",
      color: { kind: "srgb", hex: "00AA44" },
    });
    expect(rereadShape.outline).toMatchObject({
      width: 25400,
      fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } },
      dashStyle: "dash",
    });
    expect(slideXml).toContain(`<a:srgbClr val="00AA44"`);
    expect(slideXml).toContain(`<a:ln w="25400">`);
    expect(slideXml).toContain(`<a:prstDash val="dash"`);
  });

  it("Writes noFill for shape fill and connector outline while preserving connector details.", () => {
    const source = readPptx(buildShapeStyleFixture());
    const shape = findShapeByName(source, "Styled");
    const connector = findConnectorByName(source, "Connector");
    const edited = setShapeOutline(
      setShapeFill(source, shape.handle!, { kind: "none" }),
      connector.handle!,
      { fill: { kind: "none" } },
    );
    const reread = readPptx(writePptx(edited));

    expect(findShapeByName(reread, "Styled").fill).toEqual({ kind: "none" });
    expect(findConnectorByName(reread, "Connector").outline).toMatchObject({
      width: 12700,
      fill: { kind: "none" },
      tailEnd: { type: "triangle", width: "med", length: "med" },
    });
  });

  it("Writes repeated direct helper shape style changes as compact final edits.", () => {
    const source = readPptx(buildShapeStyleFixture());
    const shape = findShapeByName(source, "Styled");
    const edited = setShapeOutline(
      setShapeFill(
        setShapeOutline(
          setShapeFill(source, shape.handle!, {
            kind: "solid",
            color: { kind: "srgb", hex: "00aa44" },
          }),
          shape.handle!,
          { fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } } },
        ),
        shape.handle!,
        { kind: "none" },
      ),
      shape.handle!,
      { width: asEmu(38100) },
    );
    const rereadShape = findShapeByName(readPptx(writePptx(edited)), "Styled");

    expect(edited.edits).toHaveLength(2);
    expect(rereadShape.fill).toEqual({ kind: "none" });
    expect(rereadShape.outline).toMatchObject({
      width: 38100,
      fill: { kind: "solid", color: { kind: "srgb", hex: "336699" } },
    });
  });

  it("Rejects conflicting shape fill and outline edit journals.", () => {
    const source = readPptx(buildShapeStyleFixture());
    const shape = findShapeByName(source, "Styled");
    const withConflictingFillEdits = {
      ...source,
      edits: [
        { kind: "updateShapeFill", handle: shape.handle!, fill: { kind: "none" } },
        {
          kind: "updateShapeFill",
          handle: shape.handle!,
          fill: { kind: "solid", color: { kind: "srgb", hex: "FFFFFF" } },
        },
      ],
    } satisfies typeof source;
    const withConflictingOutlineEdits = {
      ...source,
      edits: [
        { kind: "updateShapeOutline", handle: shape.handle!, outline: { width: asEmu(1) } },
        { kind: "updateShapeOutline", handle: shape.handle!, outline: { width: asEmu(2) } },
      ],
    } satisfies typeof source;

    expect(() => writePptx(withConflictingFillEdits)).toThrow(/conflicting shape fill edits/);
    expect(() => writePptx(withConflictingOutlineEdits)).toThrow(/conflicting shape outline edits/);
  });
});

describe("writePptx - shape xfrm edit", () => {
  it("edits existing master and nested layout shape properties by part-local identity", () => {
    let authored = createPptx();
    const masterHandle = authored.slideMasters[0]?.handle;
    const layoutHandle = authored.slideLayouts[0]?.handle;
    if (masterHandle === undefined || layoutHandle === undefined) {
      throw new Error("master or layout handle not found");
    }
    authored = addShape(authored, layoutHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(10),
      offsetY: asEmu(20),
      width: asEmu(300),
      height: asEmu(400),
    });
    authored = addShape(authored, layoutHandle, {
      geometry: { kind: "preset", preset: "ellipse" },
      offsetX: asEmu(500),
      offsetY: asEmu(600),
      width: asEmu(700),
      height: asEmu(800),
    });
    const layoutShapes = authored.slideLayouts[0]?.shapes ?? [];
    authored = groupShapes(authored, [
      requireHandle(layoutShapes.at(-2)?.handle),
      requireHandle(layoutShapes.at(-1)?.handle),
    ]);
    authored = addShape(authored, masterHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(1100),
      height: asEmu(1200),
    });
    const existing = readPptx(writePptx(authored));
    const layoutGroup = existing.slideLayouts[0]?.shapes.at(-1);
    const masterShape = existing.slideMasters[0]?.shapes.at(-1);
    if (layoutGroup?.kind !== "group" || masterShape?.handle === undefined) {
      throw new Error("existing group or master shape not found");
    }
    const nestedShape = layoutGroup.children[0];
    if (nestedShape?.handle === undefined) throw new Error("nested layout shape not found");

    let edited = updateShapeTransform(existing, nestedShape.handle, {
      offsetX: asEmu(111),
      offsetY: asEmu(222),
      width: asEmu(333),
      height: asEmu(444),
    });
    edited = setShapeFill(edited, nestedShape.handle, {
      kind: "solid",
      color: { kind: "srgb", hex: "12AB34" },
    });
    edited = setShapeOutline(edited, masterShape.handle, {
      width: asEmu(12700),
      fill: { kind: "solid", color: { kind: "srgb", hex: "56789A" } },
    });
    const output = writePptx(edited);
    const reread = readPptx(output);

    expect(findShapeNodeBySourceHandle(reread, nestedShape.handle)).toMatchObject({
      kind: "shape",
      transform: { offsetX: 111, offsetY: 222, width: 333, height: 444 },
      fill: { kind: "solid", color: { kind: "srgb", hex: "12AB34" } },
    });
    expect(findShapeNodeBySourceHandle(reread, masterShape.handle)).toMatchObject({
      outline: {
        width: 12700,
        fill: { kind: "solid", color: { kind: "srgb", hex: "56789A" } },
      },
    });
    expect(decoder.decode(getEntry(output, layoutHandle.partPath))).toContain(
      '<a:off x="111" y="222"',
    );
    expect(decoder.decode(getEntry(output, masterHandle.partPath))).toContain("56789A");
  });

  it("rejects duplicate ids and AlternateContent targets in layout parts", () => {
    const base = unzipSync(writePptx(createPptx()));
    const layoutPath = "ppt/slideLayouts/slideLayout1.xml";
    const originalLayout = base[layoutPath];
    if (originalLayout === undefined) throw new Error("layout part not found");
    const original = decoder.decode(originalLayout);
    const shapeXml = (id: string, name: string) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm>` +
      `<a:prstGeom prst="rect"/></p:spPr></p:sp>`;

    const duplicateFiles = { ...base };
    duplicateFiles[layoutPath] = encoder.encode(
      original.replace(
        "</p:spTree>",
        `${shapeXml("50", "First")}${shapeXml("50", "Second")}</p:spTree>`,
      ),
    );
    const duplicate = readPptx(zipSync(duplicateFiles));
    const duplicateHandle = duplicate.slideLayouts[0]?.shapes.at(-2)?.handle;
    if (duplicateHandle === undefined) throw new Error("duplicate layout handle not found");
    expect(() => setShapeFill(duplicate, duplicateHandle, { kind: "none" })).toThrow(
      /duplicate node id/,
    );

    const alternateFiles = { ...base };
    alternateFiles[layoutPath] = encoder.encode(
      original.replace(
        "</p:spTree>",
        `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
          `<mc:Fallback>${shapeXml("60", "Fallback")}</mc:Fallback></mc:AlternateContent>` +
          `</p:spTree>`,
      ),
    );
    const alternate = readPptx(zipSync(alternateFiles));
    const alternateHandle = alternate.slideLayouts[0]?.shapes.at(-1)?.handle;
    if (alternateHandle === undefined) throw new Error("AlternateContent handle not found");
    expect(() => setShapeFill(alternate, alternateHandle, { kind: "none" })).toThrow(
      /AlternateContent/,
    );
  });

  it("Apply offset and extent update to PptxSourceModel source and reflect in PPTX after write", () => {
    const source = readPptx(buildTextEditFixture());
    const shape = firstShape(source);
    const handle = shape.handle!;

    const edited = updateShapeTransform(source, handle, {
      offsetX: asEmu(1111),
      offsetY: asEmu(2222),
      width: asEmu(3333),
      height: asEmu(4444),
    });
    const reread = readPptx(writePptx(edited));
    const editedShape = findShapeNodeBySourceHandle(reread, handle);

    expect(firstShape(edited).transform).toMatchObject({
      offsetX: 1111,
      offsetY: 2222,
      width: 3333,
      height: 4444,
    });
    expect(editedShape?.transform).toMatchObject({
      offsetX: 1111,
      offsetY: 2222,
      width: 3333,
      height: 4444,
    });
  });

  it("Rejects conflicting shape transform edits for the same shape", () => {
    const source = readPptx(buildTextEditFixture());
    const handle = firstShape(source).handle!;
    const edited = updateShapeTransform(
      updateShapeTransform(source, handle, {
        offsetX: asEmu(1111),
        offsetY: asEmu(2222),
        width: asEmu(3333),
        height: asEmu(4444),
      }),
      handle,
      {
        offsetX: asEmu(5555),
        offsetY: asEmu(6666),
        width: asEmu(7777),
        height: asEmu(8888),
      },
    );

    expect(() => writePptx(edited)).toThrow(/conflicting shape transform edits/);
  });

  it("Preserves unrelated package material while replacing only dirty slide XML", () => {
    const input = buildTextEditFixture();
    const source = readPptx(input);
    const edited = updateShapeTransform(source, firstShape(source).handle!, {
      offsetX: asEmu(1111),
      offsetY: asEmu(2222),
      width: asEmu(3333),
      height: asEmu(4444),
    });
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml).toContain('<a:off x="1111" y="2222"');
    expect(slideXml).toContain('<a:ext cx="3333" cy="4444"');
    expect(getEntry(output, "docProps/custom.xml")).toEqual(getEntry(input, "docProps/custom.xml"));
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(
      getEntry(input, "ppt/media/image1.png"),
    );
  });

  it("Rejects shape handles that do not have xfrm", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="50" name="No xfrm"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp>`,
      ),
    );
    const shapeWithoutXfrm = firstShape(source);

    expect(() =>
      updateShapeTransform(source, shapeWithoutXfrm.handle!, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/does not reference a shape with xfrm/);
  });

  it("Rejects shape transform handles without node ids", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr name="No Id Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp>`,
      ),
    );
    const shape = firstShape(source);

    expect(shape.handle?.nodeId).toBeUndefined();
    expect(() =>
      updateShapeTransform(source, shape.handle!, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/requires a node id/);
  });

  it("Rejects nonexistent shape handles", () => {
    const source = readPptx(buildTextEditFixture());
    const handle = {
      ...firstShape(source).handle!,
      nodeId: asSourceNodeId("999"),
    };

    expect(() =>
      updateShapeTransform(source, handle, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/shape handle was not found/);
  });

  it("Patches a nested group child by stable node identity and replaces only the dirty slide", () => {
    const archive = unzipSync(
      buildTextEditFixtureFromSlide(
        `<p:grpSp>` +
          `<p:nvGrpSpPr><p:cNvPr id="30" name="Outer Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
          `<p:grpSpPr><a:xfrm rot="5400000" flipH="1"><a:off x="10" y="20"/><a:ext cx="300" cy="400"/><a:chOff x="0" y="0"/><a:chExt cx="300" cy="400"/></a:xfrm></p:grpSpPr>` +
          `<p:grpSp>` +
          `<p:nvGrpSpPr><p:cNvPr id="31" name="Inner Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
          `<p:grpSpPr><a:xfrm><a:off x="5" y="6"/><a:ext cx="100" cy="200"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="200"/></a:xfrm></p:grpSpPr>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="32" name="Nested Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
          `</p:sp>` +
          `</p:grpSp>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="33" name="Preserved Sibling"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="901" y="902"/><a:ext cx="903" cy="904"/></a:xfrm><a:prstGeom prst="ellipse"/></p:spPr>` +
          `<p:extLst><p:ext uri="preserve-sibling"/></p:extLst></p:sp>` +
          `</p:grpSp>`,
      ),
    );
    archive["ppt/slides/slide2.xml"] = xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
        `<p:cSld name="unrelated-slide"/></p:sld>`,
    );
    const input = zipSync(archive);
    const source = readPptx(input);
    const outer = source.slides[0].shapes[0];
    if (outer.kind !== "group") throw new Error("outer group not found");
    const inner = outer.children[0];
    if (inner.kind !== "group" || inner.handle === undefined) {
      throw new Error("inner group not found");
    }
    const child = inner.children[0];
    if (child.kind !== "shape" || child.handle === undefined) {
      throw new Error("nested child not found");
    }
    const stableHandle = { ...child.handle, orderingSlot: 999 };

    const edited = setShapeOutline(
      setShapeFill(
        updateShapeTransform(
          updateShapeTransform(
            source,
            { ...inner.handle, orderingSlot: 998 },
            {
              offsetX: asEmu(51),
              offsetY: asEmu(61),
              width: asEmu(101),
              height: asEmu(201),
            },
          ),
          stableHandle,
          {
            offsetX: asEmu(111),
            offsetY: asEmu(222),
            width: asEmu(333),
            height: asEmu(444),
          },
        ),
        stableHandle,
        { kind: "solid", color: { kind: "srgb", hex: "12ab34" } },
      ),
      stableHandle,
      {
        width: asEmu(12700),
        fill: { kind: "solid", color: { kind: "srgb", hex: "56789a" } },
      },
    );
    const output = writePptx(edited);
    const reread = readPptx(output);
    const editedChild = findShapeNodeBySourceHandle(reread, child.handle);
    const editedInner = findShapeNodeBySourceHandle(reread, inner.handle);
    const rereadOuter = reread.slides[0].shapes[0];
    if (rereadOuter.kind !== "group") throw new Error("reread outer group not found");

    expect(findShapeNodeBySourceHandle(source, stableHandle)).toBe(child);
    expect(() => deleteShape(source, stableHandle)).toThrow(
      /nested group shape deletion is not supported/,
    );
    expect(editedChild).toMatchObject({
      kind: "shape",
      transform: { offsetX: 111, offsetY: 222, width: 333, height: 444 },
      fill: { kind: "solid", color: { kind: "srgb", hex: "12AB34" } },
      outline: {
        width: 12700,
        fill: { kind: "solid", color: { kind: "srgb", hex: "56789A" } },
      },
    });
    expect(editedInner).toMatchObject({
      kind: "group",
      transform: { offsetX: 51, offsetY: 61, width: 101, height: 201 },
    });
    expect(rereadOuter).toMatchObject({
      transform: {
        offsetX: 10,
        offsetY: 20,
        width: 300,
        height: 400,
        rotation: 5400000,
        flipHorizontal: true,
      },
    });
    expect(
      rereadOuter.children.map((node) => (node.kind === "raw" ? undefined : node.name)),
    ).toEqual(["Inner Group", "Preserved Sibling"]);
    expect(rereadOuter.children[1]).toMatchObject({
      kind: "shape",
      transform: { offsetX: 901, offsetY: 902, width: 903, height: 904 },
    });
    expect(decoder.decode(getEntry(output, "ppt/slides/slide1.xml"))).toContain(
      'uri="preserve-sibling"',
    );
    expect(getEntry(output, "ppt/slides/slide2.xml")).toEqual(
      getEntry(input, "ppt/slides/slide2.xml"),
    );
    expect(getEntry(output, "docProps/custom.xml")).toEqual(getEntry(input, "docProps/custom.xml"));
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(
      getEntry(input, "ppt/media/image1.png"),
    );
  });

  it("Preserves nested group XML and child order across no-edit write and reread", () => {
    const input = buildTextEditFixtureFromSlide(
      `<p:grpSp>` +
        `<p:nvGrpSpPr><p:cNvPr id="30" name="Outer Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm rot="5400000" flipH="1"><a:off x="10" y="20"/><a:ext cx="600" cy="400"/>` +
        `<a:chOff x="50" y="60"/><a:chExt cx="300" cy="100"/></a:xfrm></p:grpSpPr>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="31" name="First Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="51" y="61"/><a:ext cx="20" cy="30"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:sp>` +
        `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="32" name="Second Child"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
        `<p:spPr><a:xfrm><a:off x="71" y="81"/><a:ext cx="40" cy="50"/></a:xfrm><a:prstGeom prst="straightConnector1"/></p:spPr></p:cxnSp>` +
        `<p:grpSp>` +
        `<p:nvGrpSpPr><p:cNvPr id="33" name="Third Nested Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm flipV="1"><a:off x="91" y="101"/><a:ext cx="120" cy="80"/>` +
        `<a:chOff x="9" y="10"/><a:chExt cx="60" cy="20"/></a:xfrm></p:grpSpPr>` +
        `</p:grpSp>` +
        `</p:grpSp>`,
    );
    const originalSlideXml = decoder.decode(getEntry(input, "ppt/slides/slide1.xml"));

    const output = writePptx(readPptx(input));
    const writtenSlideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);
    const outer = reread.slides[0].shapes[0];
    if (outer?.kind !== "group") throw new Error("outer group not found");
    const nested = outer.children[2];
    if (nested?.kind !== "group") throw new Error("nested group not found");

    expect(writtenSlideXml).toBe(originalSlideXml);
    expect(outer.children.map((child) => (child.kind === "raw" ? undefined : child.name))).toEqual([
      "First Child",
      "Second Child",
      "Third Nested Group",
    ]);
    expect(outer).toMatchObject({
      transform: {
        offsetX: 10,
        offsetY: 20,
        width: 600,
        height: 400,
        rotation: 5400000,
        flipHorizontal: true,
      },
      childTransform: { offsetX: 50, offsetY: 60, width: 300, height: 100 },
    });
    expect(nested.childTransform).toEqual({
      offsetX: 9,
      offsetY: 10,
      width: 60,
      height: 20,
    });
  });

  it("Rejects AlternateContent fallback shape handles for this writer slice", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
          `<mc:Fallback>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="40" name="Fallback"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
          `</p:sp>` +
          `</mc:Fallback>` +
          `</mc:AlternateContent>`,
      ),
    );
    const shape = firstShape(source);

    expect(() =>
      updateShapeTransform(source, shape.handle!, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/AlternateContent/);
  });

  it("Rejects duplicate ids, AlternateContent descendants, and id-less nested targets", () => {
    const duplicate = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="50" name="Root duplicate"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr></p:sp>` +
          `<p:grpSp>` +
          `<p:nvGrpSpPr><p:cNvPr id="60" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
          `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/><a:chOff x="0" y="0"/><a:chExt cx="10" cy="10"/></a:xfrm></p:grpSpPr>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="50" name="Nested duplicate"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr></p:sp>` +
          `</p:grpSp>`,
      ),
    );
    const duplicateHandle = duplicate.slides[0].shapes[0]?.handle;
    if (duplicateHandle === undefined) throw new Error("duplicate handle not found");
    expect(() => findShapeNodeBySourceHandle(duplicate, duplicateHandle)).toThrow(
      /duplicate node id/,
    );
    expect(() =>
      updateShapeTransform(duplicate, duplicateHandle, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/duplicate node id/);

    const alternateContent = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:grpSp>` +
          `<p:nvGrpSpPr><p:cNvPr id="70" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
          `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/><a:chOff x="0" y="0"/><a:chExt cx="10" cy="10"/></a:xfrm></p:grpSpPr>` +
          `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
          `<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="71" name="Fallback child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr></p:sp></mc:Fallback>` +
          `</mc:AlternateContent></p:grpSp>`,
      ),
    );
    const alternateGroup = alternateContent.slides[0].shapes[0];
    if (alternateGroup.kind !== "group") throw new Error("alternate group not found");
    const alternateChild = alternateGroup.children[0];
    if (alternateChild?.handle === undefined) throw new Error("alternate child not found");
    expect(() => setShapeFill(alternateContent, alternateChild.handle!, { kind: "none" })).toThrow(
      /AlternateContent/,
    );

    const mixedAlternateContent = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:grpSp>` +
          `<p:nvGrpSpPr><p:cNvPr id="75" name="Mixed Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
          `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/><a:chOff x="0" y="0"/><a:chExt cx="10" cy="10"/></a:xfrm></p:grpSpPr>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="76" name="Regular sibling"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr></p:sp>` +
          `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
          `<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="77" name="Fallback sibling"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="5" y="6"/><a:ext cx="7" cy="8"/></a:xfrm></p:spPr></p:sp></mc:Fallback>` +
          `</mc:AlternateContent></p:grpSp>`,
      ),
    );
    const mixedGroup = mixedAlternateContent.slides[0].shapes[0];
    if (mixedGroup.kind !== "group") throw new Error("mixed group not found");
    const regularSibling = mixedGroup.children[0];
    if (regularSibling?.handle === undefined) throw new Error("regular sibling not found");
    const mixedEdited = setShapeFill(mixedAlternateContent, regularSibling.handle, {
      kind: "solid",
      color: { kind: "srgb", hex: "ABCDEF" },
    });
    expect(
      findShapeNodeBySourceHandle(readPptx(writePptx(mixedEdited)), regularSibling.handle),
    ).toMatchObject({
      kind: "shape",
      fill: { kind: "solid", color: { kind: "srgb", hex: "ABCDEF" } },
    });

    const idless = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:grpSp>` +
          `<p:nvGrpSpPr><p:cNvPr id="80" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
          `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/><a:chOff x="0" y="0"/><a:chExt cx="10" cy="10"/></a:xfrm></p:grpSpPr>` +
          `<p:sp><p:nvSpPr><p:cNvPr name="Idless child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr></p:sp>` +
          `</p:grpSp>`,
      ),
    );
    const idlessGroup = idless.slides[0].shapes[0];
    if (idlessGroup.kind !== "group") throw new Error("idless group not found");
    const idlessChild = idlessGroup.children[0];
    if (idlessChild?.handle === undefined) throw new Error("idless child not found");
    expect(() =>
      updateShapeTransform(idless, idlessChild.handle!, {
        offsetX: asEmu(1),
        offsetY: asEmu(2),
        width: asEmu(3),
        height: asEmu(4),
      }),
    ).toThrow(/requires a node id/);
  });
});
