import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";

// Import via the actual public surface (`@pptx-glimpse/document`).
import {
  addEmptySlideFromLayout,
  asPartPath,
  deleteSlide,
  duplicateSlide,
  moveSlide,
  readPptx,
  replaceTextRunPlainText,
  type SourceConnector,
  writePptx,
} from "../index.js";
import {
  buildConnectedShapeFixture,
  buildRoundTripFixture,
  buildSlideTopologyFixture,
  buildSlideTopologyFixtureWithRelationshipOverrides,
  buildTextEditFixture,
  buildUnreferencedLayoutFixture,
  decoder,
  firstRun,
  getEntry,
} from "./write-pptx.test-helpers.js";

describe("writePptx - slide topology edits", () => {
  it("adds an empty slide that references a layout unused by existing slides", () => {
    const source = readPptx(buildUnreferencedLayoutFixture());
    const targetLayout = source.slideLayouts.find(
      (layout) => layout.partPath === "ppt/slideLayouts/slideLayout3.xml",
    );
    const edited = addEmptySlideFromLayout(source, {
      layoutPartPath: asPartPath("ppt/slideLayouts/slideLayout3.xml"),
    });
    const output = writePptx(edited);
    const entries = unzipSync(output);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));
    const presentationRels = decoder.decode(getEntry(output, "ppt/_rels/presentation.xml.rels"));
    const newSlideXml = decoder.decode(getEntry(output, "ppt/slides/slide2.xml"));
    const newSlideRels = decoder.decode(getEntry(output, "ppt/slides/_rels/slide2.xml.rels"));

    expect(targetLayout).toBeDefined();
    expect(edited.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(reread.slides[1]?.layoutPartPath).toBe("ppt/slideLayouts/slideLayout3.xml");
    expect(presentationXml).toContain(`<p:sldId id="257" r:id="rId3"/>`);
    expect(presentationRels).toContain(
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>`,
    );
    expect(newSlideRels).toContain(
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout3.xml"/>`,
    );
    expect(newSlideXml).toContain("<p:cSld><p:spTree>");
    expect(newSlideXml).not.toContain("<p:sp>");
    expect(entries["ppt/slides/slide2.xml"]).toBeDefined();
    expect(entries["ppt/slides/_rels/slide2.xml.rels"]).toBeDefined();
    expect(reread.packageGraph.contentTypes.overrides).toEqual(
      expect.arrayContaining([
        {
          partName: "ppt/slides/slide2.xml",
          contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        },
      ]),
    );
  });

  it("duplicates a slide immediately after the source and preserves raw invisible slide material", () => {
    const source = readPptx(buildSlideTopologyFixture());
    const edited = duplicateSlide(source, source.slides[0].handle!);
    const output = writePptx(edited);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));
    const presentationRels = decoder.decode(getEntry(output, "ppt/_rels/presentation.xml.rels"));
    const duplicateSlideXml = decoder.decode(getEntry(output, "ppt/slides/slide3.xml"));
    const duplicateSlideRels = decoder.decode(getEntry(output, "ppt/slides/_rels/slide3.xml.rels"));

    expect(reread.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide3.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(
      reread.slides.map((slide) => slide.shapes[0]?.kind === "shape" && slide.shapes[0].name),
    ).toEqual(["Invisible Source", "Invisible Source", "Second"]);
    expect(presentationXml).toContain(`<p:sldId id="301" r:id="rId10"/>`);
    expect(presentationXml.indexOf(`r:id="rIdSlide1"`)).toBeLessThan(
      presentationXml.indexOf(`r:id="rId10"`),
    );
    expect(presentationXml.indexOf(`r:id="rId10"`)).toBeLessThan(
      presentationXml.indexOf(`r:id="rIdSlide2"`),
    );
    expect(presentationRels).toContain(
      `<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>`,
    );
    expect(duplicateSlideXml).toContain("<p:timing>");
    expect(duplicateSlideRels).toContain(`Id="rIdComments"`);
    expect(duplicateSlideRels).toContain(`Target="../comments/comment1.xml"`);
    expect(duplicateSlideRels).toContain(`Id="rIdNotes"`);
    expect(duplicateSlideRels).toContain(`Target="../notesSlides/notesSlide2.xml"`);
    expect(decoder.decode(getEntry(output, "ppt/notesSlides/notesSlide2.xml"))).toContain(
      "<p:notes",
    );
    expect(
      decoder.decode(getEntry(output, "ppt/notesSlides/_rels/notesSlide2.xml.rels")),
    ).toContain(`notesMaster`);
    expect(
      decoder.decode(getEntry(output, "ppt/notesSlides/_rels/notesSlide2.xml.rels")),
    ).toContain(`Target="../slides/slide3.xml"`);
    expect(reread.packageGraph.contentTypes.overrides).toEqual(
      expect.arrayContaining([
        {
          partName: "ppt/slides/slide3.xml",
          contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        },
        {
          partName: "ppt/notesSlides/notesSlide2.xml",
          contentType:
            "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
        },
      ]),
    );
  });

  it("preserves connector shape IDs and connection sites when duplicating a slide", () => {
    const source = readPptx(buildConnectedShapeFixture());
    const output = writePptx(duplicateSlide(source, source.slides[0].handle!));
    const reread = readPptx(output);
    const duplicatedConnector = reread.slides[1]?.shapes.find(
      (shape): shape is SourceConnector => shape.kind === "connector",
    );

    expect(duplicatedConnector).toMatchObject({
      nodeId: "31",
      name: "Connected Shapes",
      connection: {
        start: { shapeId: "10", connectionSiteIndex: 1 },
        end: { shapeId: "30", connectionSiteIndex: 3 },
      },
      handle: { partPath: "ppt/slides/slide2.xml", nodeId: "31" },
    });
    expect(reread.slides[1]?.shapes).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          nodeId: "10",
          handle: { partPath: "ppt/slides/slide2.xml", nodeId: "10", orderingSlot: 0 },
        }),
        expect.objectContaining({
          nodeId: "30",
          handle: { partPath: "ppt/slides/slide2.xml", nodeId: "30", orderingSlot: 2 },
        }),
      ]),
    );
  });

  it("moves an existing slide by reordering presentation slide ids only", () => {
    const source = readPptx(buildSlideTopologyFixture());
    const edited = moveSlide(source, source.slides[0].handle!, { toIndex: 1 });
    const output = writePptx(edited);
    const entries = unzipSync(output);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));
    const presentationRels = decoder.decode(getEntry(output, "ppt/_rels/presentation.xml.rels"));

    expect(reread.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide2.xml",
      "ppt/slides/slide1.xml",
    ]);
    expect(
      reread.slides.map((slide) => slide.shapes[0]?.kind === "shape" && slide.shapes[0].name),
    ).toEqual(["Second", "Invisible Source"]);
    expect(presentationXml.indexOf(`r:id="rIdSlide2"`)).toBeLessThan(
      presentationXml.indexOf(`r:id="rIdSlide1"`),
    );
    expect(presentationRels).toContain(`Id="rIdSlide1"`);
    expect(presentationRels).toContain(`Id="rIdSlide2"`);
    expect(entries["ppt/slides/slide1.xml"]).toBeDefined();
    expect(entries["ppt/slides/slide2.xml"]).toBeDefined();
    expect(entries["ppt/notesSlides/notesSlide1.xml"]).toBeDefined();
  });

  it("deletes a slide and its notes part while keeping remaining slide order and orphan cleanup out of scope", () => {
    const source = readPptx(buildSlideTopologyFixture());
    const edited = deleteSlide(source, source.slides[0].handle!);
    const output = writePptx(edited);
    const entries = unzipSync(output);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));
    const presentationRels = decoder.decode(getEntry(output, "ppt/_rels/presentation.xml.rels"));

    expect(reread.presentation.slidePartPaths).toEqual(["ppt/slides/slide2.xml"]);
    expect(reread.slides[0]?.shapes[0]?.kind === "shape" && reread.slides[0].shapes[0].name).toBe(
      "Second",
    );
    expect(presentationXml).not.toContain(`r:id="rIdSlide1"`);
    expect(presentationRels).not.toContain(`Id="rIdSlide1"`);
    expect(entries["ppt/slides/slide1.xml"]).toBeUndefined();
    expect(entries["ppt/slides/_rels/slide1.xml.rels"]).toBeUndefined();
    expect(entries["ppt/notesSlides/notesSlide1.xml"]).toBeUndefined();
    expect(entries["ppt/notesSlides/_rels/notesSlide1.xml.rels"]).toBeUndefined();
    expect(entries["ppt/comments/comment1.xml"]).toBeDefined();
    expect(
      reread.packageGraph.contentTypes.overrides.some(
        (override) => override.partName === "ppt/notesSlides/notesSlide1.xml",
      ),
    ).toBe(false);
  });

  it("keeps relationship content type overrides consistent when no rels default exists", () => {
    const source = readPptx(buildSlideTopologyFixtureWithRelationshipOverrides());
    const duplicated = readPptx(writePptx(duplicateSlide(source, source.slides[0].handle!)));
    const deleted = readPptx(writePptx(deleteSlide(source, source.slides[0].handle!)));
    const duplicatedOverrides = duplicated.packageGraph.contentTypes.overrides;
    const deletedOverridePartNames = new Set(
      deleted.packageGraph.contentTypes.overrides.map((override) => override.partName),
    );

    expect(
      duplicated.packageGraph.contentTypes.defaults.some((entry) => entry.extension === "rels"),
    ).toBe(false);
    expect(duplicatedOverrides).toEqual(
      expect.arrayContaining([
        {
          partName: "ppt/slides/_rels/slide3.xml.rels",
          contentType: "application/vnd.openxmlformats-package.relationships+xml",
        },
        {
          partName: "ppt/notesSlides/_rels/notesSlide2.xml.rels",
          contentType: "application/vnd.openxmlformats-package.relationships+xml",
        },
      ]),
    );
    expect(deletedOverridePartNames.has("ppt/slides/slide1.xml")).toBe(false);
    expect(deletedOverridePartNames.has("ppt/slides/_rels/slide1.xml.rels")).toBe(false);
    expect(deletedOverridePartNames.has("ppt/notesSlides/notesSlide1.xml")).toBe(false);
    expect(deletedOverridePartNames.has("ppt/notesSlides/_rels/notesSlide1.xml.rels")).toBe(false);
  });

  it("rejects deleting the last slide and duplicating a dirty slide", () => {
    const singleSlide = readPptx(buildTextEditFixture());
    expect(() => deleteSlide(singleSlide, singleSlide.slides[0].handle!)).toThrow(/last slide/);

    const dirty = replaceTextRunPlainText(singleSlide, firstRun(singleSlide).handle!, "Dirty");
    expect(() => duplicateSlide(dirty, dirty.slides[0].handle!)).toThrow(
      /pending dirty part edits/,
    );
  });

  it("does not reuse a deleted slide part name within the same edit journal", () => {
    const source = readPptx(buildRoundTripFixture());
    const deletedSecond = deleteSlide(source, source.slides[1].handle!);
    const duplicatedFirst = duplicateSlide(deletedSecond, deletedSecond.slides[0].handle!);
    const deletedDuplicate = deleteSlide(duplicatedFirst, duplicatedFirst.slides[1].handle!);
    const output = writePptx(deletedDuplicate);
    const reread = readPptx(output);
    const presentationXml = decoder.decode(getEntry(output, "ppt/presentation.xml"));

    expect(duplicatedFirst.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide3.xml",
    ]);
    expect(reread.presentation.slidePartPaths).toEqual(["ppt/slides/slide1.xml"]);
    expect(presentationXml).not.toContain(`r:id="rIdSlide2"`);
    expect(presentationXml).not.toContain(`r:id="rId3"`);
  });

  it("assigns new slide numeric ids at edit time and the writer only applies them", () => {
    const source = readPptx(buildSlideTopologyFixture());
    const duplicatedOnce = duplicateSlide(source, source.slides[0].handle!);
    const duplicatedTwice = duplicateSlide(duplicatedOnce, duplicatedOnce.slides[0].handle!);
    const newSlideNumericIds = (duplicatedTwice.edits ?? []).flatMap((edit) =>
      edit.kind === "duplicateSlide" ? [edit.newSlideNumericId] : [],
    );
    const presentationXml = decoder.decode(
      getEntry(writePptx(duplicatedTwice), "ppt/presentation.xml"),
    );

    expect(newSlideNumericIds).toEqual([301, 302]);
    expect(presentationXml).toContain(`<p:sldId id="301"`);
    expect(presentationXml).toContain(`<p:sldId id="302"`);
  });
});
