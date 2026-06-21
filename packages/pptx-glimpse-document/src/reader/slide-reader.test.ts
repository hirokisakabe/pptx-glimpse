import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import type { SourceImage, SourceRawShapeNode, SourceShape } from "../experimental.js";
// 実際の公開面 (`@pptx-glimpse/document/experimental`) 経由で import する。
import { readPptx } from "../experimental.js";

const encoder = new TextEncoder();

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

function fixture(name: string): Uint8Array {
  return readFileSync(
    fileURLToPath(new URL(`../../../../shared-fixtures/${name}`, import.meta.url)),
  );
}

describe("readPptx — typed slide reading (real fixtures)", () => {
  const product = readPptx(fixture("real-product-page.pptx"));
  const basic = readPptx(fixture("real-basic-theme.pptx"));

  it("slide → layout → master → theme の chain を typed に辿る", () => {
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

  it("simple p:sp を transform / geometry / fill 付きの source node として読む", () => {
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

  it("noFill の shape は fill kind=none として読む", () => {
    const shape = firstShape(product, "Text 1");
    expect(shape.fill).toEqual({ kind: "none" });
  });

  it("plain text の paragraph / run と basic run properties を読む", () => {
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

  it("複数 paragraph を持つ text body を読み、先頭空白を保持する", () => {
    const shape = firstShape(product, "Text 2");
    expect(shape.textBody?.paragraphs).toHaveLength(2);
    const secondParagraphText = shape.textBody!.paragraphs[1].runs[0].text;
    expect(secondParagraphText.startsWith("      From onboarding")).toBe(true);
  });

  it("placeholder の type / index metadata を保持する", () => {
    const title = firstShape(basic, "Google Shape;86;p13");
    expect(title.placeholder).toEqual({ type: "ctrTitle" });

    const subtitle = firstShape(basic, "Google Shape;87;p13");
    expect(subtitle.placeholder).toEqual({ type: "subTitle", index: 1 });
  });

  it("embedded raster p:pic を relationship 参照付きで読み、media へ解決できる", () => {
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

    // blip relationship → target → media part bytes を package graph 経由で解決する。
    const slideRels = basic.packageGraph.relationships.find(
      (rel) => rel.sourcePartPath === "ppt/slides/slide2.xml",
    );
    const blipRel = slideRels?.relationships.find((rel) => rel.id === image?.blipRelationshipId);
    expect(blipRel?.target).toBe("../media/image1.png");
    const media = basic.packageGraph.media.find((m) => m.partPath === "ppt/media/image1.png");
    expect(media && media.bytes.length).toBeGreaterThan(0);
  });

  it("未対応の shape tree node (graphicFrame) を raw shape node として保持する", () => {
    const slide2 = basic.slides.find((s) => s.partPath === "ppt/slides/slide2.xml");
    const raw = slide2!.shapes.find((s): s is SourceRawShapeNode => s.kind === "raw");
    expect(raw?.raw.node.name).toBe("p:graphicFrame");
  });

  it("typed shape 内の未対応子要素を raw sidecar として保持する", () => {
    // `a:effectLst` (outerShdw) は typed 化しないため sidecar として残す。
    const shadowed = firstShape(product, "Shape 3");
    const names = shadowed.rawSidecars?.map((sidecar) => sidecar.node.name) ?? [];
    expect(names).toContain("a:effectLst");

    // run property の `a:ea` / `a:cs` も sidecar として保持する。
    const run = firstShape(product, "Text 0").textBody!.paragraphs[0].runs[0];
    const runSidecarNames = run.rawSidecars?.map((sidecar) => sidecar.node.name) ?? [];
    expect(runSidecarNames).toContain("a:ea");
    expect(runSidecarNames).toContain("a:cs");

    // 画像の `a:stretch` / blip 配下の `a:alphaModFix` も保持する。
    const slide2 = basic.slides.find((s) => s.partPath === "ppt/slides/slide2.xml");
    const image = slide2!.shapes.find((s): s is SourceImage => s.kind === "image");
    const imageSidecarNames = image?.rawSidecars?.map((sidecar) => sidecar.node.name) ?? [];
    expect(imageSidecarNames).toContain("a:stretch");
    expect(imageSidecarNames).toContain("a:alphaModFix");
  });

  it("master の clrMap と theme の color / font scheme を読む", () => {
    const master = product.slideMasters[0];
    expect(master.colorMap?.mapping.bg1).toBe("lt1");
    expect(master.colorMap?.mapping.tx1).toBe("dk1");

    const theme = product.themes[0];
    expect(theme.colorScheme?.colors.accent1).toBeDefined();
    expect(theme.fontScheme).toBeDefined();
  });
});

/**
 * 色変換 / 線 / 回転 / gradient fill / 未対応属性など、real fixture では
 * 揃わない構造を決定的に検証するための合成 PPTX。
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

describe("readPptx — typed shape detail (synthetic)", () => {
  it("scheme color + lumMod transform / outline width+color / rotation+flip を読む", () => {
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
    const shape = source.slides[0].shapes[0] as SourceShape;
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

  it("gradient fill と custom geometry を raw として保持する", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:sp><p:nvSpPr><p:cNvPr id="11" name="Grad"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr>` +
          `<a:custGeom><a:pathLst/></a:custGeom>` +
          `<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs></a:gsLst></a:gradFill>` +
          `</p:spPr></p:sp>`,
      ),
    );
    const shape = source.slides[0].shapes[0] as SourceShape;
    // gradient fill は typed 化せず raw fill として保持する。
    expect(shape.fill?.kind).toBe("raw");
    // custGeom は raw sidecar として保持する。
    const names = shape.rawSidecars?.map((sidecar) => sidecar.node.name) ?? [];
    expect(names).toContain("a:custGeom");
  });

  it("raw sidecar が未対応要素の属性・子要素を保持する", () => {
    const source = readPptx(
      buildSyntheticPptx(
        `<p:sp><p:nvSpPr><p:cNvPr id="12" name="Ext"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:effectLst><a:outerShdw blurRad="40000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst></p:spPr></p:sp>`,
      ),
    );
    const shape = source.slides[0].shapes[0] as SourceShape;
    const effect = shape.rawSidecars?.find((sidecar) => sidecar.node.name === "a:effectLst");
    const outerShdw = effect?.node.children?.[0];
    expect(outerShdw?.name).toBe("a:outerShdw");
    expect(outerShdw?.attributes?.blurRad).toBe("40000");
    expect(outerShdw?.children?.[0].name).toBe("a:srgbClr");
    expect(outerShdw?.children?.[0].attributes?.val).toBe("000000");
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
