import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";

// Import via the actual public surface (`@pptx-glimpse/document`).
import { readPptx } from "../index.js";

const encoder = new TextEncoder();

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

/**
 * Synthetic PPTX for precise validation of acceptance conditions. 2 slides (ordered by rId
 * intentionally arranged in reverse), 1 media, unsupported part (docProps/custom.xml),
 * external relationships, including relative targets containing `../`.
 */
function buildSyntheticPptx(): Uint8Array {
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst>` +
        // The order of sldIdLst (slide1 -> slide2) is the truth of slide order.
        `<p:sldId id="256" r:id="rIdSlide1"/>` +
        `<p:sldId id="257" r:id="rIdSlide2"/>` +
        `</p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`),
    "ppt/slides/slide2.xml": xml(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`),
    "ppt/slides/_rels/slide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        // Relative targets and external relationships that include `../`.
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>` +
        `</Relationships>`,
    ),
    "ppt/media/image1.png": new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]),
    "docProps/custom.xml": xml(`<Properties><custom/></Properties>`),
  };
  return zipSync(files);
}

function buildMasterLayoutOrderPptx(options: {
  readonly includeIdLists: boolean;
  readonly includeUnresolvedIds?: boolean;
}): Uint8Array {
  const masterIdList = options.includeIdLists
    ? `<p:sldMasterIdLst>` +
      `<p:sldMasterId id="2147483648" r:id="rIdMaster2"/>` +
      `<p:sldMasterId id="2147483649" r:id="rIdMissingMaster"/>`.repeat(
        options.includeUnresolvedIds ? 1 : 0,
      ) +
      `<p:sldMasterId id="2147483650" r:id="rIdMaster1"/>` +
      `</p:sldMasterIdLst>`
    : "";
  const layoutIdList = options.includeIdLists
    ? `<p:sldLayoutIdLst>` +
      `<p:sldLayoutId id="2147483649" r:id="rIdLayout2"/>` +
      `<p:sldLayoutId id="2147483650" r:id="rIdMissingLayout"/>`.repeat(
        options.includeUnresolvedIds ? 1 : 0,
      ) +
      `<p:sldLayoutId id="2147483651" r:id="rIdLayout1"/>` +
      `</p:sldLayoutIdLst>`
    : "";

  return zipSync({
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
        `<Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        masterIdList +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        // Relationship order intentionally differs from p:sldMasterIdLst order.
        `<Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
        `<Relationship Id="rIdMaster2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideMasters/slideMaster1.xml": xml(
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        layoutIdList +
        `</p:sldMaster>`,
    ),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        // Relationship order intentionally differs from p:sldLayoutIdLst order.
        `<Relationship Id="rIdLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
        `<Relationship Id="rIdLayout2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideMasters/slideMaster2.xml": xml(
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sldMaster>`,
    ),
    "ppt/slideLayouts/slideLayout1.xml": xml(
      `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
    ),
    "ppt/slideLayouts/slideLayout2.xml": xml(
      `<p:sldLayout show="0" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
    ),
  });
}

describe("readPptx source model parsing", () => {
  const source = readPptx(buildSyntheticPptx());

  it("Returns a PptxSourceModel source containing presentation metadata", () => {
    expect(source.presentation.partPath).toBe("ppt/presentation.xml");
    expect(source.presentation.handle?.partPath).toBe("ppt/presentation.xml");
    // slide is read typed according to presentation order (cSld is empty, so shapes is empty).
    expect(source.slides.map((slide) => slide.partPath)).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(source.slides.every((slide) => slide.shapes.length === 0)).toBe(true);
    // This composite fixture has no slideLayout relationship, so the chain cannot be followed.
    expect(source.slideLayouts).toEqual([]);
    expect(source.slideMasters).toEqual([]);
    expect(source.themes).toEqual([]);
    expect(source.diagnostics).toContainEqual(
      expect.objectContaining({ code: "slide-layout-unresolved" }),
    );
  });

  it("Can get slide count / slide order / slide size", () => {
    expect(source.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(source.presentation.slideSize).toEqual({ width: 9144000, height: 5143500 });
  });

  it("Can maintain relationship IDs / targets / target modes", () => {
    const slideRels = source.packageGraph.relationships.find(
      (rel) => rel.sourcePartPath === "ppt/slides/slide1.xml",
    );
    expect(slideRels?.relationships).toEqual([
      {
        id: "rId1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        target: "../media/image1.png",
      },
      {
        id: "rId2",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target: "https://example.com/",
        targetMode: "External",
      },
    ]);

    // Package root rels retains sourcePartPath as "".
    const rootRels = source.packageGraph.relationships.find((rel) => rel.sourcePartPath === "");
    expect(rootRels?.relationships[0]?.id).toBe("rId1");
  });

  it("Can maintain content type defaults / overrides", () => {
    expect(source.packageGraph.contentTypes.defaults).toContainEqual({
      extension: "png",
      contentType: "image/png",
    });
    // The override PartName is normalized to PartPath by removing the leading slash.
    expect(source.packageGraph.contentTypes.overrides).toContainEqual({
      partName: "ppt/slides/slide1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    });
  });

  it("Can hold media bytes and part paths", () => {
    expect(source.packageGraph.media).toEqual([
      {
        partPath: "ppt/media/image1.png",
        contentType: "image/png",
        bytes: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]),
      },
    ]);
  });

  it("Can retain unsupported package parts as raw material", () => {
    const rawPaths = source.packageGraph.rawParts?.map((part) => part.partPath) ?? [];
    // Unsupported parts/slides/presentations that are not converted to typed are retained as raw.
    expect(rawPaths).toContain("docProps/custom.xml");
    expect(rawPaths).toContain("ppt/slides/slide1.xml");
    expect(rawPaths).toContain("ppt/presentation.xml");
    // Content types / rels / media are managed separately as structural data / media and are not included in raw.
    expect(rawPaths).not.toContain("[Content_Types].xml");
    expect(rawPaths).not.toContain("ppt/media/image1.png");
    expect(rawPaths.some((path) => path.endsWith(".rels"))).toBe(false);

    const customPart = source.packageGraph.rawParts?.find(
      (part) => part.partPath === "docProps/custom.xml",
    );
    expect(customPart?.kind).toBe("binary");
    expect(customPart?.contentType).toBe("application/xml");
  });

  it("part manifest lists all parts with content type", () => {
    const partMap = new Map(
      source.packageGraph.parts.map((part) => [part.partPath, part.contentType]),
    );
    expect(partMap.get("ppt/media/image1.png")).toBe("image/png");
    expect(partMap.get("ppt/slides/slide1.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    // rels part resolves content type with Default extension.
    expect(partMap.get("ppt/_rels/presentation.xml.rels")).toBe(
      "application/vnd.openxmlformats-package.relationships+xml",
    );
    // Do not include [Content_Types].xml itself in the part manifest.
    expect(partMap.has("[Content_Types].xml")).toBe(false);
  });

  it("Throws an error if presentation part is missing", () => {
    const bogus = zipSync({
      "[Content_Types].xml": xml(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
      ),
    });
    expect(() => readPptx(bogus)).toThrow(/presentation part not found/);
  });

  it("Throws an error if officeDocument relationship points to something other than presentation", () => {
    const bogus = zipSync({
      "[Content_Types].xml": xml(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `</Types>`,
      ),
      "_rels/.rels": xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          // officeDocument incorrectly points to another XML part.
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="docProps/app.xml"/>` +
          `</Relationships>`,
      ),
      "docProps/app.xml": xml(`<Properties/>`),
    });
    expect(() => readPptx(bogus)).toThrow(/not a presentation part/);
  });

  it("If p:sldId points to a relationship other than slide, exclude it and leave diagnostic", () => {
    const bogus = zipSync({
      "[Content_Types].xml": xml(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `</Types>`,
      ),
      "_rels/.rels": xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
          `</Relationships>`,
      ),
      "ppt/presentation.xml": xml(
        `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
          `<p:sldIdLst><p:sldId id="256" r:id="rIdBogus"/></p:sldIdLst>` +
          `</p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": xml(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          // A relationship that points to notesMaster instead of slide.
          `<Relationship Id="rIdBogus" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>` +
          `</Relationships>`,
      ),
    });
    const result = readPptx(bogus);
    expect(result.presentation.slidePartPaths).toEqual([]);
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({ code: "slide-relationship-invalid" }),
    );
  });
});

describe("readPptx master and layout authoring order", () => {
  it("uses master and layout id-list order instead of relationship order", () => {
    const source = readPptx(buildMasterLayoutOrderPptx({ includeIdLists: true }));

    expect(source.presentation.slideMasterPartPaths).toEqual([
      "ppt/slideMasters/slideMaster2.xml",
      "ppt/slideMasters/slideMaster1.xml",
    ]);
    expect(source.slideMasters.map((master) => master.partPath)).toEqual(
      source.presentation.slideMasterPartPaths,
    );
    expect(
      source.slideMasters.find((master) => master.partPath === "ppt/slideMasters/slideMaster1.xml")
        ?.layoutPartPaths,
    ).toEqual(["ppt/slideLayouts/slideLayout2.xml", "ppt/slideLayouts/slideLayout1.xml"]);
    expect(
      source.slideLayouts.find((layout) => layout.partPath === "ppt/slideLayouts/slideLayout2.xml")
        ?.show,
    ).toBe(false);
  });

  it("reports unresolved relationships from master and layout id lists", () => {
    const source = readPptx(
      buildMasterLayoutOrderPptx({ includeIdLists: true, includeUnresolvedIds: true }),
    );

    expect(
      source.diagnostics.find(
        (diagnostic) => diagnostic.code === "slide-master-relationship-unresolved",
      )?.handle?.relationshipId,
    ).toBe("rIdMissingMaster");
    expect(
      source.diagnostics.find(
        (diagnostic) => diagnostic.code === "slide-layout-relationship-unresolved",
      )?.handle?.relationshipId,
    ).toBe("rIdMissingLayout");
  });

  it("falls back to relationship order when id lists are absent", () => {
    const source = readPptx(buildMasterLayoutOrderPptx({ includeIdLists: false }));

    expect(source.presentation.slideMasterPartPaths).toEqual([
      "ppt/slideMasters/slideMaster1.xml",
      "ppt/slideMasters/slideMaster2.xml",
    ]);
    expect(source.slideMasters[0]?.layoutPartPaths).toEqual([
      "ppt/slideLayouts/slideLayout1.xml",
      "ppt/slideLayouts/slideLayout2.xml",
    ]);
  });
});

describe("readPptx - real fixture smoke test", () => {
  const fixturePath = fileURLToPath(
    new URL("../../../../shared-fixtures/real-basic-theme.pptx", import.meta.url),
  );
  const source = readPptx(readFileSync(fixturePath));

  it("Can read slide order / size / media from real-basic-theme", () => {
    expect(source.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
    expect(source.presentation.slideSize?.width).toBeGreaterThan(0);
    expect(source.presentation.slideSize?.height).toBeGreaterThan(0);

    const image = source.packageGraph.media.find(
      (part) => part.partPath === "ppt/media/image1.png",
    );
    expect(image?.contentType).toBe("image/png");
    expect(image && image.bytes.length).toBeGreaterThan(0);

    // All slide parts are kept as raw for round-trip.
    const rawPaths = source.packageGraph.rawParts?.map((part) => part.partPath) ?? [];
    expect(rawPaths).toContain("ppt/slides/slide1.xml");
    expect(rawPaths).toContain("ppt/slides/slide2.xml");
  });
});
