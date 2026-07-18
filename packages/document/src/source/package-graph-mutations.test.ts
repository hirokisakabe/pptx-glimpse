import { describe, expect, it } from "vitest";

import { asPartPath, asRelationshipId } from "./handles.js";
import type { PackageGraph } from "./package-graph.js";
import {
  addMediaPartRelationship,
  addPackagePart,
  addPartRelationship,
  nextNumberedName,
  nextNumberedPartPath,
  nextRelationshipId,
  removePackageParts,
  removePartRelationship,
} from "./package-graph-mutations.js";

const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const SLIDE_PATH = asPartPath("ppt/slides/slide1.xml");
const SLIDE_RELS_PATH = asPartPath("ppt/slides/_rels/slide1.xml.rels");

function buildGraph(overrides?: Partial<PackageGraph>): PackageGraph {
  return {
    contentTypes: {
      defaults: [{ extension: "rels", contentType: RELS_CONTENT_TYPE }],
      overrides: [
        { partName: asPartPath("ppt/slides/slide1.xml"), contentType: SLIDE_CONTENT_TYPE },
      ],
    },
    parts: [{ partPath: asPartPath("ppt/slides/slide1.xml"), contentType: SLIDE_CONTENT_TYPE }],
    relationships: [
      {
        sourcePartPath: asPartPath("ppt/presentation.xml"),
        relationships: [
          { id: asRelationshipId("rId1"), type: SLIDE_REL_TYPE, target: "slides/slide1.xml" },
        ],
      },
      {
        sourcePartPath: asPartPath("ppt/slides/slide1.xml"),
        relationships: [],
      },
    ],
    media: [],
    rawParts: [
      {
        kind: "binary",
        partPath: asPartPath("ppt/slides/slide1.xml"),
        contentType: SLIDE_CONTENT_TYPE,
        bytes: new Uint8Array([1]),
      },
    ],
    ...overrides,
  };
}

describe("addPackagePart", () => {
  it("Add a part with its .rels part to all four package graph lists", () => {
    const graph = buildGraph();
    const bytes = new Uint8Array([2]);
    const next = addPackagePart(graph, {
      partPath: asPartPath("ppt/slides/slide2.xml"),
      contentType: SLIDE_CONTENT_TYPE,
      bytes,
      relationships: {
        sourcePartPath: asPartPath("ppt/slides/slide2.xml"),
        relationships: [],
      },
    });

    expect(next.parts).toEqual([
      ...graph.parts,
      { partPath: "ppt/slides/slide2.xml", contentType: SLIDE_CONTENT_TYPE },
      { partPath: "ppt/slides/_rels/slide2.xml.rels", contentType: RELS_CONTENT_TYPE },
    ]);
    expect(next.contentTypes.overrides).toEqual([
      ...graph.contentTypes.overrides,
      { partName: "ppt/slides/slide2.xml", contentType: SLIDE_CONTENT_TYPE },
    ]);
    expect(next.relationships).toEqual([
      ...graph.relationships,
      { sourcePartPath: "ppt/slides/slide2.xml", relationships: [] },
    ]);
    expect(next.rawParts).toEqual([
      ...(graph.rawParts ?? []),
      {
        kind: "binary",
        partPath: "ppt/slides/slide2.xml",
        contentType: SLIDE_CONTENT_TYPE,
        bytes,
      },
    ]);
    expect(graph.parts).toHaveLength(1);
  });

  it("Do not register a .rels part when the new part owns no relationships", () => {
    const next = addPackagePart(buildGraph(), {
      partPath: asPartPath("ppt/slides/slide2.xml"),
      contentType: SLIDE_CONTENT_TYPE,
      bytes: new Uint8Array([2]),
    });

    expect(next.parts.map((part) => part.partPath)).not.toContain(
      "ppt/slides/_rels/slide2.xml.rels",
    );
    expect(next.relationships).toHaveLength(2);
  });

  it("Add a content type override for the .rels part only when no rels default exists", () => {
    const graph = buildGraph({
      contentTypes: {
        defaults: [],
        overrides: [],
      },
    });
    const next = addPackagePart(graph, {
      partPath: asPartPath("ppt/slides/slide2.xml"),
      contentType: SLIDE_CONTENT_TYPE,
      bytes: new Uint8Array([2]),
      relationships: {
        sourcePartPath: asPartPath("ppt/slides/slide2.xml"),
        relationships: [],
      },
    });

    expect(next.contentTypes.overrides).toEqual([
      { partName: "ppt/slides/slide2.xml", contentType: SLIDE_CONTENT_TYPE },
      { partName: "ppt/slides/_rels/slide2.xml.rels", contentType: RELS_CONTENT_TYPE },
    ]);
  });
});

describe("addMediaPartRelationship", () => {
  it.each([
    { hasRelationshipGroup: false, hasRelationshipPart: false, hasRelsDefault: false },
    { hasRelationshipGroup: false, hasRelationshipPart: false, hasRelsDefault: true },
    { hasRelationshipGroup: false, hasRelationshipPart: true, hasRelsDefault: false },
    { hasRelationshipGroup: false, hasRelationshipPart: true, hasRelsDefault: true },
    { hasRelationshipGroup: true, hasRelationshipPart: false, hasRelsDefault: false },
    { hasRelationshipGroup: true, hasRelationshipPart: false, hasRelsDefault: true },
    { hasRelationshipGroup: true, hasRelationshipPart: true, hasRelsDefault: false },
    { hasRelationshipGroup: true, hasRelationshipPart: true, hasRelsDefault: true },
  ])(
    "registers media consistently with relationship group=$hasRelationshipGroup, .rels part=$hasRelationshipPart, rels default=$hasRelsDefault",
    ({ hasRelationshipGroup, hasRelationshipPart, hasRelsDefault }) => {
      const base = buildGraph();
      const graph = buildGraph({
        contentTypes: {
          defaults: hasRelsDefault ? [{ extension: "rels", contentType: RELS_CONTENT_TYPE }] : [],
          overrides: [
            ...base.contentTypes.overrides,
            ...(hasRelationshipPart && !hasRelsDefault
              ? [{ partName: SLIDE_RELS_PATH, contentType: RELS_CONTENT_TYPE }]
              : []),
          ],
        },
        parts: [
          ...base.parts,
          ...(hasRelationshipPart
            ? [{ partPath: SLIDE_RELS_PATH, contentType: RELS_CONTENT_TYPE }]
            : []),
        ],
        relationships: hasRelationshipGroup
          ? base.relationships
          : base.relationships.filter(
              (relationships) => relationships.sourcePartPath !== SLIDE_PATH,
            ),
      });
      const bytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
      const media = {
        partPath: asPartPath("ppt/media/image1.png"),
        contentType: "image/png",
        bytes,
      };
      const relationship = {
        id: asRelationshipId("rId2"),
        type: IMAGE_REL_TYPE,
        target: "../media/image1.png",
      };

      const next = addMediaPartRelationship(graph, {
        ownerPartPath: SLIDE_PATH,
        media,
        extension: "png",
        relationship,
        contentTypeDefaultConflictError: (existingContentType) =>
          new Error(`unexpected existing content type: ${existingContentType}`),
      });

      expect(next.parts.filter((part) => part.partPath === media.partPath)).toEqual([
        { partPath: media.partPath, contentType: media.contentType },
      ]);
      expect(next.parts.filter((part) => part.partPath === SLIDE_RELS_PATH)).toHaveLength(1);
      expect(next.contentTypes.defaults).toContainEqual({
        extension: "png",
        contentType: "image/png",
      });
      expect(
        next.contentTypes.overrides.filter((override) => override.partName === SLIDE_RELS_PATH),
      ).toHaveLength(hasRelsDefault ? 0 : 1);
      expect(
        next.relationships.filter((relationships) => relationships.sourcePartPath === SLIDE_PATH),
      ).toEqual([{ sourcePartPath: SLIDE_PATH, relationships: [relationship] }]);
      expect(next.media).toEqual([media]);
      expect(next.rawParts).toBe(graph.rawParts);
      expect(next.rawParts?.some((part) => part.partPath === media.partPath)).toBe(false);
      expect(graph.media).toEqual([]);
    },
  );

  it("uses the caller-provided error when an extension has a conflicting content type", () => {
    const graph = buildGraph({
      contentTypes: {
        defaults: [
          { extension: "rels", contentType: RELS_CONTENT_TYPE },
          { extension: "png", contentType: "application/not-png" },
        ],
        overrides: [],
      },
    });

    expect(() =>
      addMediaPartRelationship(graph, {
        ownerPartPath: SLIDE_PATH,
        media: {
          partPath: asPartPath("ppt/media/image1.png"),
          contentType: "image/png",
          bytes: new Uint8Array([1]),
        },
        extension: "png",
        relationship: {
          id: asRelationshipId("rId2"),
          type: IMAGE_REL_TYPE,
          target: "../media/image1.png",
        },
        contentTypeDefaultConflictError: (existingContentType) =>
          new Error(`operation-specific: ${existingContentType}`),
      }),
    ).toThrow("operation-specific: application/not-png");
  });
});

describe("removePackageParts", () => {
  it("Remove the parts and their .rels parts from all four package graph lists", () => {
    const graph = buildGraph({
      parts: [
        { partPath: asPartPath("ppt/slides/slide1.xml"), contentType: SLIDE_CONTENT_TYPE },
        {
          partPath: asPartPath("ppt/slides/_rels/slide1.xml.rels"),
          contentType: RELS_CONTENT_TYPE,
        },
        { partPath: asPartPath("ppt/presentation.xml"), contentType: "presentation" },
      ],
      contentTypes: {
        defaults: [],
        overrides: [
          { partName: asPartPath("ppt/slides/slide1.xml"), contentType: SLIDE_CONTENT_TYPE },
          {
            partName: asPartPath("ppt/slides/_rels/slide1.xml.rels"),
            contentType: RELS_CONTENT_TYPE,
          },
        ],
      },
      rawParts: [
        {
          kind: "binary",
          partPath: asPartPath("ppt/slides/slide1.xml"),
          contentType: SLIDE_CONTENT_TYPE,
          bytes: new Uint8Array([1]),
        },
        {
          kind: "binary",
          partPath: asPartPath("ppt/slides/_rels/slide1.xml.rels"),
          contentType: RELS_CONTENT_TYPE,
          bytes: new Uint8Array([2]),
        },
      ],
    });
    const next = removePackageParts(graph, [asPartPath("ppt/slides/slide1.xml")]);

    expect(next.parts.map((part) => part.partPath)).toEqual(["ppt/presentation.xml"]);
    expect(next.contentTypes.overrides).toEqual([]);
    expect(next.relationships.map((relationships) => relationships.sourcePartPath)).toEqual([
      "ppt/presentation.xml",
    ]);
    expect(next.rawParts).toEqual([]);
  });

  it("Keep rawParts undefined when the graph preserves no raw package material", () => {
    const graph = buildGraph({ rawParts: undefined });

    expect(removePackageParts(graph, [asPartPath("ppt/slides/slide1.xml")]).rawParts).toBe(
      undefined,
    );
    expect(
      addPackagePart(graph, {
        partPath: asPartPath("ppt/slides/slide2.xml"),
        contentType: SLIDE_CONTENT_TYPE,
        bytes: new Uint8Array([2]),
      }).rawParts,
    ).toHaveLength(1);
  });
});

describe("addPartRelationship / removePartRelationship", () => {
  it("Add and remove a relationship only on the owning part", () => {
    const graph = buildGraph();
    const added = addPartRelationship(graph, asPartPath("ppt/presentation.xml"), {
      id: asRelationshipId("rId2"),
      type: SLIDE_REL_TYPE,
      target: "slides/slide2.xml",
    });

    expect(added.relationships[0].relationships.map((relationship) => relationship.id)).toEqual([
      "rId1",
      "rId2",
    ]);
    expect(added.relationships[1].relationships).toEqual([]);

    const removed = removePartRelationship(
      added,
      asPartPath("ppt/presentation.xml"),
      asRelationshipId("rId1"),
    );
    expect(removed.relationships[0].relationships.map((relationship) => relationship.id)).toEqual([
      "rId2",
    ]);
  });

  it("Return the graph unchanged when no relationships entry owns the source part", () => {
    const graph = buildGraph();
    const next = addPartRelationship(graph, asPartPath("ppt/unknown.xml"), {
      id: asRelationshipId("rId9"),
      type: SLIDE_REL_TYPE,
      target: "slides/slide9.xml",
    });

    expect(next.relationships).toEqual(graph.relationships);
  });
});

describe("numbering helpers", () => {
  it("Continue relationship ids after the maximum trailing number", () => {
    expect(
      nextRelationshipId([
        { id: asRelationshipId("rId1"), type: SLIDE_REL_TYPE, target: "slides/slide1.xml" },
        { id: asRelationshipId("rIdSlide9"), type: SLIDE_REL_TYPE, target: "slides/slide9.xml" },
        { id: asRelationshipId("rIdNotes"), type: SLIDE_REL_TYPE, target: "notes" },
      ]),
    ).toBe("rId10");
    expect(nextRelationshipId([])).toBe("rId1");
  });

  it("Allocate the next numbered part path across parts, overrides, rawParts and reserved paths", () => {
    const graph = buildGraph({
      contentTypes: {
        defaults: [],
        overrides: [
          { partName: asPartPath("ppt/slides/slide5.xml"), contentType: SLIDE_CONTENT_TYPE },
        ],
      },
    });

    expect(nextNumberedPartPath(graph, [], "ppt/slides/slide", ".xml")).toBe(
      "ppt/slides/slide6.xml",
    );
    expect(nextNumberedPartPath(graph, ["ppt/slides/slide7.xml"], "ppt/slides/slide", ".xml")).toBe(
      "ppt/slides/slide8.xml",
    );
    expect(nextNumberedPartPath(graph, [], "ppt/notesSlides/notesSlide", ".xml")).toBe(
      "ppt/notesSlides/notesSlide1.xml",
    );

    const withRawOnlyPart = buildGraph({
      rawParts: [
        {
          kind: "binary",
          partPath: asPartPath("ppt/slides/slide9.xml"),
          contentType: SLIDE_CONTENT_TYPE,
          bytes: new Uint8Array([1]),
        },
      ],
    });
    expect(nextNumberedPartPath(withRawOnlyPart, [], "ppt/slides/slide", ".xml")).toBe(
      "ppt/slides/slide10.xml",
    );
  });

  it("Ignore names that do not match the pattern when computing the maximum", () => {
    const used = new Set(["item2abc", "item3"]);
    expect(nextNumberedName(used, /^item(\d+)$/, (index) => `item${index}`)).toBe("item4");
  });
});
