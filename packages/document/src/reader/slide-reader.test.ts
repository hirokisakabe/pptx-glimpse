import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import type {
  SourceChart,
  SourceConnector,
  SourceGroup,
  SourceImage,
  SourceShape,
  SourceSmartArt,
  SourceTable,
} from "../index.js";
// Import via the actual public surface (`@pptx-glimpse/document`).
import { readPptx } from "../index.js";
import { unsafeFixtureAssertion } from "../unsafe-type-assertion.js";

const encoder = new TextEncoder();

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

function fixture(name: string): Uint8Array {
  return readFileSync(
    fileURLToPath(new URL(`../../../../shared-fixtures/${name}`, import.meta.url)),
  );
}

describe("readPptx - typed slide reading (real fixtures)", () => {
  const product = readPptx(fixture("real-product-page.pptx"));
  const basic = readPptx(fixture("real-basic-theme.pptx"));

  it("Follow the chain of slide -> layout -> master -> theme to typed", () => {
    const [slide] = product.slides;
    expect(slide.partPath).toBe("ppt/slides/slide1.xml");
    expect(slide.layoutPartPath).toBe("ppt/slideLayouts/slideLayout1.xml");

    const layout = product.slideLayouts.find((l) => l.partPath === slide.layoutPartPath);
    expect(layout?.masterPartPath).toBe("ppt/slideMasters/slideMaster1.xml");

    const master = product.slideMasters.find((m) => m.partPath === layout?.masterPartPath);
    expect(master?.themePartPath).toBe("ppt/theme/theme1.xml");
    expect(master?.layoutPartPaths).toContain("ppt/slideLayouts/slideLayout1.xml");

    expect(product.themes.map((t) => t.partPath)).toContain(master?.themePartPath);
  });

  it("Read simple p:sp as source node with transform / geometry / fill", () => {
    const shape = firstShape(product, "Text 0");
    expect(shape.kind).toBe("shape");
    expect(shape.nodeId).toBe("2");
    expect(shape.name).toBe("Text 0");
    expect(shape.transform).toEqual({
      offsetX: 5238750,
      offsetY: 457200,
      width: 1714500,
      height: 304800,
    });
    expect(shape.geometry).toEqual({ preset: "roundRect" });
    expect(shape.fill).toEqual({ kind: "solid", color: { kind: "srgb", hex: "DBEAFE" } });
  });

  it("noFill shape is read as fill kind=none", () => {
    const shape = firstShape(product, "Text 1");
    expect(shape.fill).toEqual({ kind: "none" });
  });

  it("Read plain text paragraph / run and basic run properties", () => {
    const shape = firstShape(product, "Text 0");
    const body = shape.textBody;
    expect(body?.properties?.anchor).toBe("middle");
    expect(body?.paragraphs).toHaveLength(1);

    const [paragraph] = body!.paragraphs;
    expect(paragraph.properties?.align).toBe("center");
    expect(paragraph.runs).toHaveLength(1);

    const [run] = paragraph.runs;
    expect(run.kind).toBe("textRun");
    expect(run.text).toBe("NEW RELEASE");
    expect(run.properties).toMatchObject({
      bold: true,
      fontSize: 10.5,
      typeface: "Noto Sans JP",
      color: { kind: "srgb", hex: "2563EB" },
    });
  });

  it("Read text body with multiple paragraphs and preserve leading whitespace", () => {
    const shape = firstShape(product, "Text 2");
    expect(shape.textBody?.paragraphs).toHaveLength(2);
    const secondParagraphText = shape.textBody!.paragraphs[1].runs[0].text;
    expect(secondParagraphText.startsWith("      From onboarding")).toBe(true);
  });

  it("Preserve type/index metadata of placeholder", () => {
    const title = firstShape(basic, "Google Shape;86;p13");
    expect(title.placeholder).toEqual({ type: "ctrTitle" });

    const subtitle = firstShape(basic, "Google Shape;87;p13");
    expect(subtitle.placeholder).toEqual({ type: "subTitle", index: 1 });
  });

  it("Embedded raster p:pic can be read with relationship reference and resolved to media", () => {
    const slide2 = basic.slides.find((s) => s.partPath === "ppt/slides/slide2.xml");
    const image = slide2!.shapes.find((s): s is SourceImage => s.kind === "image");
    expect(image?.nodeId).toBe("98");
    expect(image?.blipRelationshipId).toBe("rId3");
    expect(image?.transform).toEqual({
      offsetX: 2895500,
      offsetY: 2215600,
      width: 1201300,
      height: 1201300,
    });

    // Resolve blip relationship -> target -> media part bytes via package graph.
    const slideRels = basic.packageGraph.relationships.find(
      (rel) => rel.sourcePartPath === "ppt/slides/slide2.xml",
    );
    const blipRel = slideRels?.relationships.find((rel) => rel.id === image?.blipRelationshipId);
    expect(blipRel?.target).toBe("../media/image1.png");
    const media = basic.packageGraph.media.find((m) => m.partPath === "ppt/media/image1.png");
    expect(media && media.bytes.length).toBeGreaterThan(0);
  });

  it("Keep table graphicFrame as typed table node", () => {
    const slide2 = basic.slides.find((s) => s.partPath === "ppt/slides/slide2.xml");
    const table = slide2!.shapes.find((s): s is SourceTable => s.kind === "table");
    expect(table?.table.columns.length).toBeGreaterThan(0);
    expect(table?.table.rows.length).toBeGreaterThan(0);
  });

  it("Preserve typed shape effects and unsupported child elements", () => {
    const shadowed = firstShape(product, "Shape 3");
    expect(shadowed.effects?.outerShadow?.blurRadius).toBeGreaterThan(0);
    expect(shadowed.effects?.outerShadow?.color).toBeDefined();
    const names = shadowed.rawSidecars?.map((sidecar) => sidecar.node.name) ?? [];
    expect(names).not.toContain("a:effectLst");

    // Run properties `a:ea` / `a:cs` are retained as typed source properties.
    const run = firstShape(product, "Text 0").textBody!.paragraphs[0].runs[0];
    expect(run.properties).toMatchObject({
      typefaceEa: "Noto Sans JP",
      typefaceCs: "Noto Sans JP",
    });
    const runSidecarNames = run.rawSidecars?.map((sidecar) => sidecar.node.name) ?? [];
    expect(runSidecarNames).not.toContain("a:ea");
    expect(runSidecarNames).not.toContain("a:cs");

    // The image default `a:stretch` is interpreted as typed, and unsupported blip children are retained.
    const slide2 = basic.slides.find((s) => s.partPath === "ppt/slides/slide2.xml");
    const image = slide2!.shapes.find((s): s is SourceImage => s.kind === "image");
    const imageSidecarNames = image?.rawSidecars?.map((sidecar) => sidecar.node.name) ?? [];
    expect(image?.stretch).toBeUndefined();
    expect(imageSidecarNames).toContain("a:alphaModFix");
  });

  it("Read master's clrMap and theme's color/font scheme", () => {
    const master = product.slideMasters[0];
    expect(master.colorMap?.mapping.bg1).toBe("lt1");
    expect(master.colorMap?.mapping.tx1).toBe("dk1");

    const theme = product.themes[0];
    expect(theme.colorScheme?.colors.accent1).toBeDefined();
    expect(theme.fontScheme).toBeDefined();
  });
});

/**
 * Real fixtures such as color conversion / line / rotation / gradient fill / unsupported attributes, etc.
 * Synthetic PPTX for definitive verification of misaligned structures.
 */
function buildSyntheticPptx(slideSpTree: string): Uint8Array {
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree>${slideSpTree}</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  };
  return zipSync(files);
}

function buildSyntheticPptxWithLayout(slideLayoutAttrs: string): Uint8Array {
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
        `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/presentation.xml": xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slides/slide1.xml": xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sld>`,
    ),
    "ppt/slides/_rels/slide1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideLayouts/slideLayout1.xml": xml(
      `<p:sldLayout${slideLayoutAttrs} xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sldLayout>`,
    ),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
        `</Relationships>`,
    ),
    "ppt/slideMasters/slideMaster1.xml": xml(
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree/></p:cSld>` +
        `</p:sldMaster>`,
    ),
  };
  return zipSync(files);
}

describe("readPptx - typed shape detail (synthetic)", () => {
  it.each([
    ['show="0"', ` show="0"`, false, false],
    ['show="false"', ` show="false"`, false, false],
    ["missing show", "", undefined, true],
  ])("Read p:sldLayout@show for %s", (_label, slideLayoutAttrs, authored, effective) => {
    const source = readPptx(buildSyntheticPptxWithLayout(slideLayoutAttrs));
    const layout = source.slideLayouts[0];

    expect(layout?.show).toBe(authored);
    expect(layout?.show ?? true).toBe(effective);
  });

  it("Read scheme color + lumMod transform / outline width+color / rotation+flip", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Themed"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr>` +
          `<a:xfrm rot="5400000" flipH="1"><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>` +
          `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>` +
          `<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/><a:alpha val="50000"/></a:schemeClr></a:solidFill>` +
          `<a:ln w="12700"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:ln>` +
          `</p:spPr></p:sp>`,
      ),
    );
    const shape = unsafeFixtureAssertion<SourceShape>(source.slides[0].shapes[0]);
    expect(shape.transform).toEqual({
      offsetX: 100,
      offsetY: 200,
      width: 300,
      height: 400,
      rotation: 5400000,
      flipHorizontal: true,
    });
    expect(shape.geometry).toEqual({ preset: "ellipse" });
    expect(shape.fill).toEqual({
      kind: "solid",
      color: {
        kind: "scheme",
        scheme: "accent1",
        transforms: [
          { kind: "lumMod", value: 75000 },
          { kind: "alpha", value: 50000 },
        ],
      },
    });
    expect(shape.outline).toEqual({
      width: 12700,
      fill: { kind: "solid", color: { kind: "srgb", hex: "FF0000" } },
    });
  });

  it("Read bodyPr autofit / wrap / vert and ea/cs run fonts as typed source property", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:sp><p:nvSpPr><p:cNvPr id="14" name="Text props"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr>` +
          `<p:txBody>` +
          `<a:bodyPr wrap="none" vert="eaVert" numCol="2"><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>` +
          `<a:p><a:r><a:rPr sz="1800"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:rPr><a:t>文</a:t></a:r></a:p>` +
          `</p:txBody>` +
          `</p:sp>`,
      ),
    );

    const shape = unsafeFixtureAssertion<SourceShape>(source.slides[0].shapes[0]);
    expect(shape.textBody?.properties).toMatchObject({
      wrap: "none",
      vert: "eaVert",
      numCol: 2,
      autoFit: "normAutofit",
      fontScale: 0.625,
      lnSpcReduction: 0.2,
    });
    expect(shape.textBody?.paragraphs[0].runs[0].properties).toMatchObject({
      typeface: "+mn-lt",
      typefaceEa: "+mn-ea",
      typefaceCs: "+mn-cs",
    });
  });

  it("interleaved bullet Preserve br / fld run when splitting pPr", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:sp><p:nvSpPr><p:cNvPr id="17" name="Interleaved text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:p>` +
          `<a:pPr><a:buChar char="&#x2022;"/></a:pPr>` +
          `<a:r><a:t>first</a:t></a:r>` +
          `<a:br/>` +
          `<a:r><a:t>after break</a:t></a:r>` +
          `<a:pPr><a:buChar char="&#x25E6;"/></a:pPr>` +
          `<a:fld id="{00000000-0000-0000-0000-000000000000}" type="slidenum"><a:t>field</a:t></a:fld>` +
          `</a:p>` +
          `<a:p><a:r><a:t>tail</a:t></a:r></a:p>` +
          `</p:txBody>` +
          `</p:sp>`,
      ),
    );

    const shape = unsafeFixtureAssertion<SourceShape>(source.slides[0].shapes[0]);
    const paragraphs = shape.textBody!.paragraphs;
    expect(paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text))).toEqual([
      ["first", "\n", "after break"],
      ["field"],
      ["tail"],
    ]);
    expect(new Set(paragraphs.map((paragraph) => paragraph.handle.nodeId)).size).toBe(3);
  });

  it("Reading spAutoFit as typed source body property", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:sp><p:nvSpPr><p:cNvPr id="15" name="Sp autofit"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr>` +
          `<p:txBody><a:bodyPr><a:spAutoFit/></a:bodyPr><a:p><a:r><a:t>grow</a:t></a:r></a:p></p:txBody>` +
          `</p:sp>`,
      ),
    );

    const shape = unsafeFixtureAssertion<SourceShape>(source.slides[0].shapes[0]);
    expect(shape.textBody?.properties).toMatchObject({ autoFit: "spAutofit" });
  });

  it("Read noAutofit as typed source body property", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:sp><p:nvSpPr><p:cNvPr id="16" name="No autofit"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr>` +
          `<p:txBody><a:bodyPr><a:noAutofit/></a:bodyPr><a:p><a:r><a:t>fixed</a:t></a:r></a:p></p:txBody>` +
          `</p:sp>`,
      ),
    );

    const shape = unsafeFixtureAssertion<SourceShape>(source.slides[0].shapes[0]);
    expect(shape.textBody?.properties).toMatchObject({
      autoFit: "noAutofit",
      fontScale: 1,
      lnSpcReduction: 0,
    });
  });

  it("Keep gradient fill as typed source fill and custom geometry as raw", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:sp><p:nvSpPr><p:cNvPr id="11" name="Grad"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr>` +
          `<a:custGeom><a:pathLst/></a:custGeom>` +
          `<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs></a:gsLst></a:gradFill>` +
          `</p:spPr></p:sp>`,
      ),
    );
    const shape = unsafeFixtureAssertion<SourceShape>(source.slides[0].shapes[0]);
    expect(shape.fill).toMatchObject({
      kind: "gradient",
      gradientType: "linear",
      stops: [{ position: 0, color: { kind: "srgb", hex: "000000" } }],
    });
    // custGeom is kept as raw sidecar.
    const names = shape.rawSidecars?.map((sidecar) => sidecar.node.name) ?? [];
    expect(names).toContain("a:custGeom");
  });

  it("Read direct effectLst as typed source effects", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:sp><p:nvSpPr><p:cNvPr id="12" name="Ext"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:effectLst>` +
          `<a:outerShdw blurRad="40000" dist="20000" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw>` +
          `<a:innerShdw blurRad="50000" dist="30000" dir="10800000"><a:srgbClr val="111111"/></a:innerShdw>` +
          `<a:glow rad="1000"><a:schemeClr val="accent1"/></a:glow>` +
          `<a:softEdge rad="3000"/>` +
          `<a:reflection blurRad="7000"/>` +
          `</a:effectLst></p:spPr></p:sp>`,
      ),
    );
    const shape = unsafeFixtureAssertion<SourceShape>(source.slides[0].shapes[0]);
    expect(shape.effects).toMatchObject({
      outerShadow: {
        blurRadius: 40000,
        distance: 20000,
        direction: 5400000,
        color: {
          kind: "srgb",
          hex: "000000",
          transforms: [{ kind: "alpha", value: 40000 }],
        },
        alignment: "ctr",
        rotateWithShape: false,
      },
      innerShadow: {
        blurRadius: 50000,
        distance: 30000,
        direction: 10800000,
        color: { kind: "srgb", hex: "111111" },
      },
      glow: { radius: 1000, color: { kind: "scheme", scheme: "accent1" } },
      softEdge: { radius: 3000 },
    });
    expect(shape.rawSidecars?.map((sidecar) => sidecar.node.name) ?? []).not.toContain(
      "a:effectLst",
    );
    expect(shape.rawSidecars?.map((sidecar) => sidecar.node.name) ?? []).toContain("a:reflection");
  });

  it("unknown rectangle alignment token falls back to default value per source field", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:sp><p:nvSpPr><p:cNvPr id="13" name="Shadow"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:effectLst>` +
          `<a:outerShdw algn="invalid"><a:srgbClr val="000000"/></a:outerShdw>` +
          `</a:effectLst></p:spPr></p:sp>` +
          `<p:pic><p:nvPicPr><p:cNvPr id="14" name="Pic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
          `<p:blipFill><a:blip r:embed="rIdImage"/><a:tile algn="invalid"/></p:blipFill>` +
          `<p:spPr/></p:pic>`,
      ),
    );

    const shape = unsafeFixtureAssertion<SourceShape>(source.slides[0].shapes[0]);
    const image = unsafeFixtureAssertion<SourceImage>(source.slides[0].shapes[1]);

    expect(shape.effects?.outerShadow?.alignment).toBe("b");
    expect(image.tile?.align).toBe("tl");
  });

  it("Read image shape effects and blip effects as typed source effects", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:pic><p:nvPicPr><p:cNvPr id="13" name="Pic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
          `<p:blipFill><a:blip r:embed="rIdImage">` +
          `<a:grayscl/><a:biLevel thresh="25000"/><a:blur rad="5000" grow="0"/>` +
          `<a:lum bright="10000" contrast="-20000"/>` +
          `<a:duotone><a:schemeClr val="accent2"/><a:srgbClr val="FFFFFF"/></a:duotone>` +
          `<a:clrChange><a:clrFrom><a:prstClr val="black"/></a:clrFrom><a:clrTo><a:schemeClr val="accent2"/></a:clrTo></a:clrChange>` +
          `<a:alphaModFix/>` +
          `</a:blip></p:blipFill>` +
          `<p:spPr><a:effectLst><a:softEdge rad="1000"/></a:effectLst></p:spPr>` +
          `</p:pic>`,
      ),
    );
    const image = unsafeFixtureAssertion<SourceImage>(source.slides[0].shapes[0]);
    expect(image.effects?.softEdge).toEqual({ radius: 1000 });
    expect(image.blipEffects).toMatchObject({
      grayscale: true,
      biLevel: { threshold: 0.25 },
      blur: { radius: 5000, grow: false },
      lum: { brightness: 0.1, contrast: -0.2 },
      duotone: {
        color1: { kind: "scheme", scheme: "accent2" },
        color2: { kind: "srgb", hex: "FFFFFF" },
      },
      clrChange: {
        from: { kind: "srgb", hex: "000000" },
        to: { kind: "scheme", scheme: "accent2" },
      },
    });
    expect(image.rawSidecars?.map((sidecar) => sidecar.node.name) ?? []).toContain("a:alphaModFix");
  });

  it("Read graphicFrame table as typed source node", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:graphicFrame>` +
          `<p:nvGraphicFramePr><p:cNvPr id="20" name="Sales table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
          `<p:xfrm><a:off x="1000" y="2000"/><a:ext cx="3000" cy="4000"/></p:xfrm>` +
          `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
          `<a:tbl>` +
          `<a:tblPr><a:tableStyleId>{style}</a:tableStyleId></a:tblPr>` +
          `<a:tblGrid><a:gridCol w="111"/><a:gridCol w="222"/></a:tblGrid>` +
          `<a:tr h="333">` +
          `<a:tc gridSpan="2" rowSpan="2">` +
          `<a:txBody><a:bodyPr/><a:p>` +
          `<a:r><a:rPr sz="1200"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>A</a:t></a:r>` +
          `<a:br/>` +
          `<a:r><a:t>B</a:t></a:r>` +
          `<a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:t>C</a:t></a:fld>` +
          `</a:p></a:txBody>` +
          `<a:tcPr><a:lnL w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnL><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:tcPr>` +
          `</a:tc>` +
          `<a:tc hMerge="1" vMerge="1">` +
          `<a:txBody><a:bodyPr/><a:p><a:r><a:t></a:t></a:r></a:p></a:txBody>` +
          `<a:tcPr/>` +
          `</a:tc>` +
          `</a:tr>` +
          `</a:tbl>` +
          `</a:graphicData></a:graphic>` +
          `</p:graphicFrame>`,
      ),
    );

    const table = unsafeFixtureAssertion<SourceTable>(source.slides[0].shapes[0]);
    expect(table.kind).toBe("table");
    expect(table.nodeId).toBe("20");
    expect(table.name).toBe("Sales table");
    expect(table.transform).toEqual({
      offsetX: 1000,
      offsetY: 2000,
      width: 3000,
      height: 4000,
    });
    expect(table.table.tableStyleId).toBe("{style}");
    expect(table.table.columns).toEqual([{ width: 111 }, { width: 222 }]);
    expect(table.table.rows[0].height).toBe(333);
    expect(table.table.rows[0].cells[0]).toMatchObject({
      gridSpan: 2,
      rowSpan: 2,
      hMerge: false,
      vMerge: false,
      fill: { kind: "solid", color: { kind: "srgb", hex: "00FF00" } },
    });
    expect(table.table.rows[0].cells[1]).toMatchObject({
      gridSpan: 1,
      rowSpan: 1,
      hMerge: true,
      vMerge: true,
    });
    expect(table.table.rows[0].cells[0].textBody?.paragraphs[0].runs[0]).toMatchObject({
      text: "A",
      properties: {
        fontSize: 12,
        color: { kind: "scheme", scheme: "accent1" },
      },
    });
    expect(
      table.table.rows[0].cells[0].textBody?.paragraphs[0].runs.map((run) => run.text),
    ).toEqual(["A", "\n", "B", "C"]);
    expect(table.table.rows[0].cells[0].borders?.left).toEqual({
      width: 12700,
      fill: { kind: "solid", color: { kind: "srgb", hex: "FF0000" } },
    });
  });

  it("Read SmartArt in graphicFrame chart and AlternateContent as typed source node", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:graphicFrame>` +
          `<p:nvGraphicFramePr><p:cNvPr id="30" name="Revenue chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
          `<p:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></p:xfrm>` +
          `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
          `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/>` +
          `</a:graphicData></a:graphic>` +
          `</p:graphicFrame>` +
          `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
          `<mc:Choice Requires="dgm">` +
          `<p:graphicFrame>` +
          `<p:nvGraphicFramePr><p:cNvPr id="31" name="Process"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
          `<p:xfrm><a:off x="500" y="600"/><a:ext cx="700" cy="800"/></p:xfrm>` +
          `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
          `<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdDiagramData" r:lo="rIdLayout" r:qs="rIdQuickStyle" r:cs="rIdColorStyle"/>` +
          `</a:graphicData></a:graphic>` +
          `</p:graphicFrame>` +
          `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="32" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr></p:cxnSp>` +
          `</mc:Choice>` +
          `</mc:AlternateContent>`,
      ),
    );

    const [chart, smartArt, connector] = unsafeFixtureAssertion<
      [SourceChart, SourceSmartArt, SourceConnector]
    >(source.slides[0].shapes);
    expect(chart).toMatchObject({
      kind: "chart",
      nodeId: "30",
      name: "Revenue chart",
      chartRelationshipId: "rIdChart",
      transform: { offsetX: 100, offsetY: 200, width: 300, height: 400 },
    });
    expect(smartArt).toMatchObject({
      kind: "smartArt",
      nodeId: "31",
      name: "Process",
      dataRelationshipId: "rIdDiagramData",
      transform: { offsetX: 500, offsetY: 600, width: 700, height: 800 },
    });
    expect(chart.rawSidecars?.map((sidecar) => sidecar.node.name)).toContain("c:chart");
    expect(smartArt.rawSidecars?.map((sidecar) => sidecar.node.name)).toEqual(
      expect.arrayContaining(["dgm:relIds", "mc:AlternateContent"]),
    );
    expect(connector).toMatchObject({ kind: "connector", nodeId: "32", name: "Connector" });
  });

  it("Read connector branch of AlternateContent as typed source node", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
          `<mc:Choice Requires="cxn">` +
          `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="40" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr></p:cxnSp>` +
          `</mc:Choice>` +
          `</mc:AlternateContent>`,
      ),
    );

    expect(source.slides[0].shapes).toHaveLength(1);
    expect(source.slides[0].shapes[0]).toMatchObject({
      kind: "connector",
      nodeId: "40",
      name: "Connector",
    });
    const [connector] = source.slides[0].shapes;
    expect(connector.kind).toBe("connector");
    if (connector.kind !== "connector") throw new Error("Expected connector");
    expect(connector.rawSidecars?.map((sidecar) => sidecar.node.name)).toContain(
      "mc:AlternateContent",
    );
  });

  it("Read group / connector / custom geometry as typed source node and preserve heterogeneous tag order", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:cxnSp>` +
          `<p:nvCxnSpPr><p:cNvPr id="50" name="Connector first"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
          `<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm>` +
          `<a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst></a:prstGeom>` +
          `<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/><a:tailEnd type="triangle" w="med" len="med"/></a:ln>` +
          `</p:spPr>` +
          `</p:cxnSp>` +
          `<p:grpSp>` +
          `<p:nvGrpSpPr><p:cNvPr id="51" name="Group second"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
          `<p:grpSpPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/><a:chOff x="5" y="6"/><a:chExt cx="30" cy="40"/></a:xfrm></p:grpSpPr>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="52" name="Group child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="30" cy="40"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>` +
          `</p:grpSp>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="53" name="Custom third"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>` +
          `<a:custGeom><a:pathLst><a:path w="1000" h="1000"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="w" y="h"/></a:lnTo></a:path></a:pathLst></a:custGeom>` +
          `</p:spPr></p:sp>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="54" name="Ordered custom"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>` +
          `<a:custGeom><a:pathLst><a:path w="1000" h="1000">` +
          `<a:moveTo><a:pt x="0" y="0"/></a:moveTo>` +
          `<a:lnTo><a:pt x="100" y="0"/></a:lnTo>` +
          `<a:quadBezTo><a:pt x="200" y="0"/><a:pt x="200" y="100"/></a:quadBezTo>` +
          `<a:lnTo><a:pt x="0" y="100"/></a:lnTo>` +
          `<a:close/>` +
          `</a:path></a:pathLst></a:custGeom>` +
          `</p:spPr></p:sp>`,
      ),
    );

    expect(source.slides[0].shapes.map((shape) => shape.kind)).toEqual([
      "connector",
      "group",
      "shape",
      "shape",
    ]);
    const [connector, group, custom, orderedCustom] = unsafeFixtureAssertion<
      [SourceConnector, SourceGroup, SourceShape, SourceShape]
    >(source.slides[0].shapes);
    expect(connector).toMatchObject({
      name: "Connector first",
      geometry: { preset: "bentConnector3", adjustValues: { adj1: 50000 } },
      outline: { width: 12700, dashStyle: "dash", tailEnd: { type: "triangle" } },
    });
    expect(group).toMatchObject({
      name: "Group second",
      transform: { offsetX: 10, offsetY: 20, width: 300, height: 400 },
      childTransform: { offsetX: 5, offsetY: 6, width: 30, height: 40 },
    });
    expect(group.children.map((child) => child.kind)).toEqual(["shape"]);
    expect(custom.geometry).toMatchObject({
      kind: "custom",
      paths: [{ width: 1000, height: 1000, commands: "M 0 0 L 1000 1000" }],
    });
    expect(orderedCustom.geometry).toMatchObject({
      kind: "custom",
      paths: [{ width: 1000, height: 1000, commands: "M 0 0 L 100 0 Q 200 0, 200 100 L 0 100 Z" }],
    });
  });

  it("Read Strict OOXML chart graphicData URI as chart source node", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:graphicFrame>` +
          `<p:nvGraphicFramePr><p:cNvPr id="41" name="Strict chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
          `<p:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></p:xfrm>` +
          `<a:graphic><a:graphicData uri="http://purl.oclc.org/ooxml/drawingml/chart">` +
          `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdStrictChart"/>` +
          `</a:graphicData></a:graphic>` +
          `</p:graphicFrame>`,
      ),
    );

    expect(source.slides[0].shapes[0]).toMatchObject({
      kind: "chart",
      nodeId: "41",
      name: "Strict chart",
      chartRelationshipId: "rIdStrictChart",
    });
  });

  it("If AlternateContent Choice is only raw, read supported Fallback branch", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
          `<mc:Choice Requires="ext"><p:contentPart><p:nvContentPartPr><p:cNvPr id="41" name="Unsupported"/></p:nvContentPartPr></p:contentPart></mc:Choice>` +
          `<mc:Fallback>` +
          `<p:sp><p:nvSpPr><p:cNvPr id="42" name="Fallback shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
          `</p:sp>` +
          `</mc:Fallback>` +
          `</mc:AlternateContent>`,
      ),
    );

    expect(source.slides[0].shapes).toHaveLength(1);
    expect(source.slides[0].shapes[0]).toMatchObject({
      kind: "shape",
      nodeId: "42",
      name: "Fallback shape",
      rawSidecars: [{ node: { name: "mc:AlternateContent" } }],
    });
  });
});

function firstShape(source: ReturnType<typeof readPptx>, name: string): SourceShape {
  for (const slide of source.slides) {
    const match = slide.shapes.find(
      (shape): shape is SourceShape => shape.kind === "shape" && shape.name === name,
    );
    if (match) return match;
  }
  throw new Error(`shape '${name}' not found`);
}
