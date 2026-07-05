import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import * as document from "@pptx-glimpse/document";
import { clearFontCache, getWarningEntries, warn } from "@pptx-glimpse/renderer";
import JSZip from "jszip";
import { afterEach, beforeAll, describe, expect, it, vi } from "vitest";

import {
  convertPptxToPng as convertPptxToPngBase,
  convertPptxToSvg as convertPptxToSvgBase,
  renderPptxSourceModelToSvg as renderPptxSourceModelToSvgBase,
} from "./converter.js";
import * as adapterModule from "./pptx-computed-view-renderer-adapter.js";

const convertPptxToSvg: typeof convertPptxToSvgBase = (input, options) =>
  convertPptxToSvgBase(input, { skipSystemFonts: true, ...options });

const convertPptxToPng: typeof convertPptxToPngBase = (input, options) =>
  convertPptxToPngBase(input, { skipSystemFonts: true, ...options });

const renderPptxSourceModelToSvg: typeof renderPptxSourceModelToSvgBase = (source, options) =>
  renderPptxSourceModelToSvgBase(source, { skipSystemFonts: true, ...options });

const SELECTED_SHARED_FIXTURES = ["real-basic-theme.pptx", "real-product-page.pptx"] as const;

const CONVERTER_TEST_SCOPE = [
  "readPptx source model",
  "createComputedView cascade projection",
  "PptxSourceModel computed-view to current renderer model adapter",
  "existing SVG renderer",
  "existing SVG to PNG conversion",
] as const;

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

const presentationXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500" type="screen16x9"/>
</p:presentation>`;

const presentationRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`;

const slide1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Rectangle 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="274638"/>
            <a:ext cx="3048000" cy="1143000"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
          <a:solidFill>
            <a:srgbClr val="4472C4"/>
          </a:solidFill>
          <a:ln w="25400">
            <a:solidFill>
              <a:srgbClr val="2F528F"/>
            </a:solidFill>
          </a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="2400" b="1">
                <a:solidFill>
                  <a:srgbClr val="FFFFFF"/>
                </a:solidFill>
                <a:latin typeface="Arial"/>
              </a:rPr>
              <a:t>Hello World</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Ellipse 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="4572000" y="274638"/>
            <a:ext cx="2286000" cy="2286000"/>
          </a:xfrm>
          <a:prstGeom prst="ellipse">
            <a:avLst/>
          </a:prstGeom>
          <a:solidFill>
            <a:srgbClr val="ED7D31"/>
          </a:solidFill>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="RoundRect 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="2743200"/>
            <a:ext cx="3048000" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="roundRect">
            <a:avLst>
              <a:gd name="adj" fmla="val 16667"/>
            </a:avLst>
          </a:prstGeom>
          <a:solidFill>
            <a:schemeClr val="accent1"/>
          </a:solidFill>
          <a:ln w="12700">
            <a:solidFill>
              <a:schemeClr val="accent1">
                <a:lumMod val="75000"/>
              </a:schemeClr>
            </a:solidFill>
          </a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="1800">
                <a:solidFill>
                  <a:srgbClr val="FFFFFF"/>
                </a:solidFill>
              </a:rPr>
              <a:t>Rounded Rectangle</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

const slide1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

const slideMaster1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>`;

const slideMaster1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

const slideLayout1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;

const slideLayout1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

const theme1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`;

async function createTestPptx(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes);
  zip.file("_rels/.rels", rootRels);
  zip.file("ppt/presentation.xml", presentationXml);
  zip.file("ppt/_rels/presentation.xml.rels", presentationRels);
  zip.file("ppt/slides/slide1.xml", slide1Xml);
  zip.file("ppt/slides/_rels/slide1.xml.rels", slide1Rels);
  zip.file("ppt/slideMasters/slideMaster1.xml", slideMaster1);
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMaster1Rels);
  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayout1);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayout1Rels);
  zip.file("ppt/theme/theme1.xml", theme1);
  return zip.generateAsync({ type: "nodebuffer" });
}

let testPptx: Buffer;

beforeAll(async () => {
  // Discards the font cache left by other tests that ran earlier within the same worker.
  clearFontCache();
  testPptx = await createTestPptx();
});

describe("public conversion orchestration", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("documents the intentionally focused dogfood scope", () => {
    expect(CONVERTER_TEST_SCOPE).toMatchInlineSnapshot(`
      [
        "readPptx source model",
        "createComputedView cascade projection",
        "PptxSourceModel computed-view to current renderer model adapter",
        "existing SVG renderer",
        "existing SVG to PNG conversion",
      ]
    `);
  });

  it.each(SELECTED_SHARED_FIXTURES)(
    "renders selected slides from %s to SVG through the public converter",
    async (fixtureName) => {
      const { slides: result } = await convertPptxToSvg(readSharedFixture(fixtureName), {
        slides: [1],
        textOutput: "text",
      });

      expect(result.map((slide) => slide.slideNumber)).toEqual([1]);
      expect(result[0]?.svg).toContain("<svg");
      expect(result[0]?.svg).toMatch(/viewBox="0 0 \d+ \d+"/);
      expect(result[0]?.svg).toContain("</svg>");
    },
  );

  it("connects the public SVG conversion to PNG conversion", async () => {
    const { slides: result } = await convertPptxToPng(readSharedFixture("real-basic-theme.pptx"), {
      slides: [1],
      width: 240,
    });

    expect(result.map((slide) => slide.slideNumber)).toEqual([1]);
    const pngSlide = result[0];
    expect(pngSlide).toMatchObject({ width: 240 });
    if (pngSlide === undefined) {
      throw new Error("Expected one PNG slide from the public converter");
    }
    expect([...pngSlide.png.subarray(0, 4)]).toEqual([0x89, 0x50, 0x4e, 0x47]);
  });

  it("keeps Node Buffer input working while returning Uint8Array PNG bytes", async () => {
    const input = Buffer.from(readSharedFixture("real-basic-theme.pptx"));
    const { slides: result } = await convertPptxToPng(input, {
      slides: [1],
      width: 240,
    });

    const pngSlide = result[0];
    if (pngSlide === undefined) {
      throw new Error("Expected one PNG slide from the public converter");
    }

    expect(pngSlide.png).toBeInstanceOf(Uint8Array);
    expect(Buffer.isBuffer(pngSlide.png)).toBe(false);
    expect([...pngSlide.png.subarray(0, 4)]).toEqual([0x89, 0x50, 0x4e, 0x47]);
  });

  it("uses the source model reader and renderer adapter for SVG conversion", async () => {
    const readPptxSpy = vi.spyOn(document, "readPptx");
    const adapterSpy = vi.spyOn(adapterModule, "adaptComputedViewToRendererModel");
    const input = readSharedFixture("real-basic-theme.pptx");
    const { slides: publicDefault } = await convertPptxToSvg(input, {
      slides: [1],
      textOutput: "text",
    });

    expect(publicDefault.map((slide) => slide.slideNumber)).toEqual([1]);
    expect(publicDefault[0]?.svg).toContain("<svg");
    expect(readPptxSpy).toHaveBeenCalledOnce();
    expect(adapterSpy).toHaveBeenCalledOnce();
  });

  it("renders from a PptxSourceModel repeatedly without re-reading PPTX bytes", async () => {
    const readPptxSpy = vi.spyOn(document, "readPptx");
    const adapterSpy = vi.spyOn(adapterModule, "adaptComputedViewToRendererModel");
    const source = document.readPptx(readSharedFixture("real-basic-theme.pptx"));

    const first = await renderPptxSourceModelToSvg(source, { slides: [1] });
    const second = await renderPptxSourceModelToSvg(source, { slides: [1] });

    expect(first.slides[0]?.svg).toContain("<svg");
    expect(second.slides[0]?.svg).toContain("<svg");
    expect(readPptxSpy).toHaveBeenCalledOnce();
    expect(adapterSpy).toHaveBeenCalledTimes(2);
  });

  it("keeps concurrent source-model renders isolated by font options", async () => {
    const source = document.readPptx(testPptx);

    const [alpha, beta] = await Promise.all([
      renderPptxSourceModelToSvg(source, {
        textOutput: "text",
        fontMapping: { Arial: "MappedAlpha" },
      }),
      renderPptxSourceModelToSvg(source, {
        textOutput: "text",
        fontMapping: { Arial: "MappedBeta" },
      }),
    ]);

    expect(alpha.slides[0]?.svg).toContain("MappedAlpha");
    expect(alpha.slides[0]?.svg).not.toContain("MappedBeta");
    expect(beta.slides[0]?.svg).toContain("MappedBeta");
    expect(beta.slides[0]?.svg).not.toContain("MappedAlpha");
  });

  it("uses the source model reader and renderer adapter for PNG conversion", async () => {
    const readPptxSpy = vi.spyOn(document, "readPptx");
    const adapterSpy = vi.spyOn(adapterModule, "adaptComputedViewToRendererModel");
    const input = readSharedFixture("real-basic-theme.pptx");
    const { slides: publicDefault } = await convertPptxToPng(input, {
      slides: [1],
      width: 240,
    });

    expect(publicDefault.map((slide) => slide.slideNumber)).toEqual([1]);
    expect(publicDefault[0]).toMatchObject({ width: 240 });
    expect(readPptxSpy).toHaveBeenCalledOnce();
    expect(adapterSpy).toHaveBeenCalledOnce();
  });
});

describe("convertPptxToSvg", () => {
  it("returns a conversion report with slides, diagnostics, and support coverage", async () => {
    const report = await convertPptxToSvg(testPptx);

    expect(Array.isArray(report)).toBe(false);
    expect(report.slides).toHaveLength(1);
    expect(report.diagnostics).toEqual(expect.any(Array));
    expect(typeof report.supportCoverage.overall.inputElements).toBe("number");
    expect(typeof report.supportCoverage.overall.outputElements).toBe("number");
    expect(typeof report.supportCoverage.overall.skippedElements).toBe("number");
    expect(typeof report.supportCoverage.overall.unresolvedElements).toBe("number");
    expect(typeof report.supportCoverage.overall.fallbackElements).toBe("number");
    expect(typeof report.supportCoverage.overall.warnings).toBe("number");
    expect(report.supportCoverage.slides[0]).toMatchObject({
      slideNumber: 1,
      inputElements: 3,
      outputElements: 3,
    });
  });

  it("integrates document reader diagnostics into the conversion report", async () => {
    const report = await convertPptxToSvg(await createPptxWithInvalidSlideRelationship());

    expect(report.slides).toHaveLength(0);
    expect(report.diagnostics).toContainEqual(
      expect.objectContaining({
        source: "document",
        severity: "warning",
        code: "slide-relationship-invalid",
        sourcePartPath: "ppt/presentation.xml",
      }),
    );
  });

  it("integrates renderer adapter diagnostics into diagnostics and support coverage", async () => {
    const report = await convertPptxToSvg(await createPptxWithRawGraphicFrame());

    expect(report.slides).toHaveLength(1);
    expect(report.diagnostics).toContainEqual(
      expect.objectContaining({
        source: "renderer-adapter",
        severity: "warning",
        code: "pptx-computed-view-adapter.raw-element-skipped",
        slideNumber: 1,
      }),
    );
    expect(report.supportCoverage.slides[0]).toMatchObject({
      slideNumber: 1,
      inputElements: 4,
      outputElements: 3,
      skippedElements: 1,
      warnings: 1,
    });
  });

  it("integrates computed view diagnostics into the conversion report", async () => {
    const report = await convertPptxToSvg(await createPptxWithSmartArtMissingShapeTree());

    expect(report.diagnostics).toContainEqual(
      expect.objectContaining({
        source: "computed-view",
        severity: "warning",
        code: "diagram-drawing-shape-tree-missing",
        slideNumber: 1,
        sourcePartPath: "ppt/diagrams/drawing1.xml",
      }),
    );
    expect(report.supportCoverage.slides[0]).toMatchObject({
      skippedElements: 0,
      unresolvedElements: 1,
    });
  });

  it("converts a PPTX file to SVG", async () => {
    const { slides } = await convertPptxToSvg(testPptx);

    expect(slides).toHaveLength(1);
    expect(slides[0].slideNumber).toBe(1);
    expect(slides[0].svg).toContain("<svg");
    expect(slides[0].svg).toContain("</svg>");
  });

  it("renders basic shapes", async () => {
    const { slides } = await convertPptxToSvg(testPptx);
    const svg = slides[0].svg;

    // Should contain rect (blue rectangle)
    expect(svg).toContain("<rect");
    // Should contain ellipse (orange)
    expect(svg).toContain("<ellipse");
    // Should contain text (as <text> or converted to <path> by opentype.js)
    expect(svg).toMatch(/<text|<path/);
  });

  it("has correct viewBox dimensions for 16:9", async () => {
    const { slides } = await convertPptxToSvg(testPptx);
    const svg = slides[0].svg;

    // 9144000 EMU = 960px, 5143500 EMU ≈ 540px
    expect(svg).toContain('viewBox="0 0 960');
  });

  it("applies fill colors", async () => {
    const { slides } = await convertPptxToSvg(testPptx);
    const svg = slides[0].svg;

    // Blue rectangle fill
    expect(svg.toLowerCase()).toContain("#4472c4");
    // Orange ellipse fill
    expect(svg.toLowerCase()).toContain("#ed7d31");
  });

  it("supports slide number filtering", async () => {
    const { slides } = await convertPptxToSvg(testPptx, { slides: [1] });

    expect(slides).toHaveLength(1);
    expect(slides[0].slideNumber).toBe(1);
  });

  it("returns empty for non-existent slide numbers", async () => {
    const { slides } = await convertPptxToSvg(testPptx, { slides: [99] });

    expect(slides).toHaveLength(0);
  });
});

function readSharedFixture(name: (typeof SELECTED_SHARED_FIXTURES)[number]): Buffer {
  return readFileSync(fileURLToPath(new URL(`../../../shared-fixtures/${name}`, import.meta.url)));
}

async function createPptxWithInvalidSlideRelationship(): Promise<Buffer> {
  const invalidPresentationXml = presentationXml.replace(
    '<p:sldId id="256" r:id="rId2"/>',
    '<p:sldId id="256" r:id="rIdBogus"/>',
  );
  const invalidPresentationRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdBogus" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>
</Relationships>`;

  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes);
  zip.file("_rels/.rels", rootRels);
  zip.file("ppt/presentation.xml", invalidPresentationXml);
  zip.file("ppt/_rels/presentation.xml.rels", invalidPresentationRels);
  zip.file("ppt/slideMasters/slideMaster1.xml", slideMaster1);
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMaster1Rels);
  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayout1);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayout1Rels);
  zip.file("ppt/theme/theme1.xml", theme1);
  return zip.generateAsync({ type: "nodebuffer" });
}

async function createPptxWithRawGraphicFrame(): Promise<Buffer> {
  const rawGraphicFrame = `
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="5" name="Unsupported Graphic"/>
          <p:cNvGraphicFramePr/>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="914400" cy="914400"/>
        </p:xfrm>
        <a:graphic>
          <a:graphicData uri="urn:pptx-glimpse:test:unsupported"/>
        </a:graphic>
      </p:graphicFrame>`;
  const slideWithRawGraphicFrame = slide1Xml.replace(
    "</p:spTree>",
    `${rawGraphicFrame}
    </p:spTree>`,
  );

  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes);
  zip.file("_rels/.rels", rootRels);
  zip.file("ppt/presentation.xml", presentationXml);
  zip.file("ppt/_rels/presentation.xml.rels", presentationRels);
  zip.file("ppt/slides/slide1.xml", slideWithRawGraphicFrame);
  zip.file("ppt/slides/_rels/slide1.xml.rels", slide1Rels);
  zip.file("ppt/slideMasters/slideMaster1.xml", slideMaster1);
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMaster1Rels);
  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayout1);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayout1Rels);
  zip.file("ppt/theme/theme1.xml", theme1);
  return zip.generateAsync({ type: "nodebuffer" });
}

async function createPptxWithSmartArtMissingShapeTree(): Promise<Buffer> {
  const smartArtFrame = `
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="5" name="SmartArt"/>
          <p:cNvGraphicFramePr/>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="914400" cy="914400"/>
        </p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
            <dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdDiagramData"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>`;
  const slideWithSmartArt = slide1Xml.replace(
    "</p:spTree>",
    `${smartArtFrame}
    </p:spTree>`,
  );
  const slideRelsWithSmartArt = slide1Rels.replace(
    "</Relationships>",
    `  <Relationship Id="rIdDiagramData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
</Relationships>`,
  );
  const diagramDataRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDiagramDrawing" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="drawing1.xml"/>
</Relationships>`;

  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes);
  zip.file("_rels/.rels", rootRels);
  zip.file("ppt/presentation.xml", presentationXml);
  zip.file("ppt/_rels/presentation.xml.rels", presentationRels);
  zip.file("ppt/slides/slide1.xml", slideWithSmartArt);
  zip.file("ppt/slides/_rels/slide1.xml.rels", slideRelsWithSmartArt);
  zip.file("ppt/slideMasters/slideMaster1.xml", slideMaster1);
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMaster1Rels);
  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayout1);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayout1Rels);
  zip.file("ppt/theme/theme1.xml", theme1);
  zip.file("ppt/diagrams/data1.xml", `<dgm:dataModel/>`);
  zip.file("ppt/diagrams/_rels/data1.xml.rels", diagramDataRels);
  zip.file(
    "ppt/diagrams/drawing1.xml",
    `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"/>`,
  );
  return zip.generateAsync({ type: "nodebuffer" });
}

describe("convertPptxToPng", () => {
  it("returns a conversion report with PNG slides and SVG-path diagnostics", async () => {
    const report = await convertPptxToPng(testPptx);

    expect(Array.isArray(report)).toBe(false);
    expect(report.slides).toHaveLength(1);
    expect(report.slides[0]).toMatchObject({ slideNumber: 1 });
    expect(report.diagnostics).toEqual(expect.any(Array));
    expect(report.supportCoverage.slides[0]).toMatchObject({
      slideNumber: 1,
      inputElements: 3,
      outputElements: 3,
    });
  });

  it("converts a PPTX file to PNG", async () => {
    const { slides } = await convertPptxToPng(testPptx);

    expect(slides).toHaveLength(1);
    expect(slides[0].slideNumber).toBe(1);
    expect(slides[0].png).toBeInstanceOf(Uint8Array);
    expect(slides[0].png.length).toBeGreaterThan(0);
    // PNG magic bytes
    expect(slides[0].png[0]).toBe(0x89);
    expect(slides[0].png[1]).toBe(0x50); // P
    expect(slides[0].png[2]).toBe(0x4e); // N
    expect(slides[0].png[3]).toBe(0x47); // G
  });

  it("respects width option", async () => {
    const { slides } = await convertPptxToPng(testPptx, { width: 480 });

    expect(slides[0].width).toBe(480);
    expect(slides[0].height).toBeGreaterThan(0);
  });
});

describe("master placeholder text filtering", () => {
  async function createPptxWithMasterPlaceholder(): Promise<Buffer> {
    const masterWithPlaceholder = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title Placeholder"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="274638"/>
            <a:ext cx="8229600" cy="1143000"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="4400"/>
              <a:t>MASTER_TITLE_TEMPLATE</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Body Placeholder"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="1600200"/>
            <a:ext cx="8229600" cy="3200400"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="2400"/>
              <a:t>MASTER_BODY_TEMPLATE</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="Decorative Logo"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="914400" cy="457200"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="1200"/>
              <a:t>MASTER_DECORATIVE_TEXT</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>`;

    const simpleSlide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="TextBox 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="2743200"/>
            <a:ext cx="3048000" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="1800"/>
              <a:t>SLIDE_ACTUAL_TEXT</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

    const zip = new JSZip();
    zip.file("[Content_Types].xml", contentTypes);
    zip.file("_rels/.rels", rootRels);
    zip.file("ppt/presentation.xml", presentationXml);
    zip.file("ppt/_rels/presentation.xml.rels", presentationRels);
    zip.file("ppt/slides/slide1.xml", simpleSlide);
    zip.file("ppt/slides/_rels/slide1.xml.rels", slide1Rels);
    zip.file("ppt/slideMasters/slideMaster1.xml", masterWithPlaceholder);
    zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMaster1Rels);
    zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayout1);
    zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayout1Rels);
    zip.file("ppt/theme/theme1.xml", theme1);
    return zip.generateAsync({ type: "nodebuffer" });
  }

  it("does not render master placeholder shapes on actual slides", async () => {
    const pptx = await createPptxWithMasterPlaceholder();
    const { slides } = await convertPptxToSvg(pptx);
    const svg = slides[0].svg;

    // Master has 3 shapes: title placeholder (id=2), body placeholder (id=3), decorative (id=4)
    // Only decorative (non-placeholder) should be rendered (red fill #FF0000)
    // Placeholder shapes should be filtered out

    // Decorative shape's red fill should appear
    expect(svg.toLowerCase()).toContain("#ff0000");

    // Count shape groups - should have decorative master shape + slide shape = 2 groups
    const groupCount = (svg.match(/<g\b[^>]*transform="translate/g) || []).length;
    expect(groupCount).toBe(2);
  });
});

describe("layout placeholder text filtering", () => {
  async function createPptxWithLayoutPlaceholder(): Promise<Buffer> {
    const layoutWithPlaceholders = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="685800" y="1143000"/>
            <a:ext cx="7772400" cy="1470025"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="4400"/>
              <a:t>LAYOUT_TITLE_TEMPLATE</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Subtitle 2"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="1371600" y="2743200"/>
            <a:ext cx="6400800" cy="1752600"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="2800"/>
              <a:t>LAYOUT_SUBTITLE_TEMPLATE</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="Date Placeholder 3"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="4800600"/>
            <a:ext cx="2133600" cy="365125"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="1000"/>
              <a:t>LAYOUT_DATE_TEMPLATE</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="5" name="Footer Placeholder 4"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="3124200" y="4800600"/>
            <a:ext cx="2895600" cy="365125"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="1000"/>
              <a:t>LAYOUT_FOOTER_TEMPLATE</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="6" name="Decorative Shape"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="914400" cy="457200"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="1200"/>
              <a:t>LAYOUT_DECORATIVE_TEXT</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;

    const simpleSlide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="TextBox 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="457200"/>
            <a:ext cx="3048000" cy="914400"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" sz="1800"/>
              <a:t>SLIDE_ACTUAL_TEXT</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

    const zip = new JSZip();
    zip.file("[Content_Types].xml", contentTypes);
    zip.file("_rels/.rels", rootRels);
    zip.file("ppt/presentation.xml", presentationXml);
    zip.file("ppt/_rels/presentation.xml.rels", presentationRels);
    zip.file("ppt/slides/slide1.xml", simpleSlide);
    zip.file("ppt/slides/_rels/slide1.xml.rels", slide1Rels);
    zip.file("ppt/slideMasters/slideMaster1.xml", slideMaster1);
    zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMaster1Rels);
    zip.file("ppt/slideLayouts/slideLayout1.xml", layoutWithPlaceholders);
    zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayout1Rels);
    zip.file("ppt/theme/theme1.xml", theme1);
    return zip.generateAsync({ type: "nodebuffer" });
  }

  it("does not render layout placeholder shapes on actual slides", async () => {
    const pptx = await createPptxWithLayoutPlaceholder();
    const { slides } = await convertPptxToSvg(pptx);
    const svg = slides[0].svg;

    // Layout has 5 shapes: ctrTitle, subTitle, dt, ftr (all placeholders), decorative (non-placeholder)
    // Only decorative (non-placeholder) should be rendered (blue fill #0000FF)
    // Placeholder shapes should be filtered out

    // Decorative shape's blue fill should appear
    expect(svg.toLowerCase()).toContain("#0000ff");

    // Count shape groups - should have decorative layout shape + slide shape = 2 groups
    const groupCount = (svg.match(/<g\b[^>]*transform="translate/g) || []).length;
    expect(groupCount).toBe(2);
  });
});

describe("slide placeholder text filtering", () => {
  async function createPptxWithSlidePlaceholders(slideXml: string): Promise<Buffer> {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", contentTypes);
    zip.file("_rels/.rels", rootRels);
    zip.file("ppt/presentation.xml", presentationXml);
    zip.file("ppt/_rels/presentation.xml.rels", presentationRels);
    zip.file("ppt/slides/slide1.xml", slideXml);
    zip.file("ppt/slides/_rels/slide1.xml.rels", slide1Rels);
    zip.file("ppt/slideMasters/slideMaster1.xml", slideMaster1);
    zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMaster1Rels);
    zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayout1);
    zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayout1Rels);
    zip.file("ppt/theme/theme1.xml", theme1);
    return zip.generateAsync({ type: "nodebuffer" });
  }

  // Slide-level placeholder shapes with no run text. PowerPoint hides these
  // (the "Click to add title" prompt is layout-side and never copied here);
  // we expect the renderer to drop them too.
  const emptyPlaceholderSlide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Empty Title Placeholder"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="274638"/>
            <a:ext cx="8229600" cy="1143000"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:endParaRPr lang="en-US"/></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Empty Body Placeholder"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="1600200"/>
            <a:ext cx="8229600" cy="3200400"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p/>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="Decorative Shape"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="914400" cy="457200"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill>
        </p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

  it("does not render empty placeholder shapes on the slide itself", async () => {
    const pptx = await createPptxWithSlidePlaceholders(emptyPlaceholderSlide);
    const { slides } = await convertPptxToSvg(pptx);
    const svg = slides[0].svg;

    // Decorative (non-placeholder) shape with magenta fill should remain.
    expect(svg.toLowerCase()).toContain("#ff00ff");

    // Empty placeholders' green fill must NOT be rendered.
    expect(svg.toLowerCase()).not.toContain("#00ff00");

    // Only the decorative shape should produce a translate group.
    const groupCount = (svg.match(/<g\b[^>]*transform="translate/g) || []).length;
    expect(groupCount).toBe(1);
  });

  it("keeps slide placeholders that contain run text", async () => {
    const filledSlide = emptyPlaceholderSlide.replace(
      `<a:p><a:endParaRPr lang="en-US"/></a:p>`,
      `<a:p><a:r><a:rPr lang="en-US" sz="3600"/><a:t>FILLED_TITLE</a:t></a:r></a:p>`,
    );

    const pptx = await createPptxWithSlidePlaceholders(filledSlide);
    const { slides } = await convertPptxToSvg(pptx);
    const svg = slides[0].svg;

    // Filled title placeholder is kept (green fill renders).
    expect(svg.toLowerCase()).toContain("#00ff00");
    // Decorative shape still rendered.
    expect(svg.toLowerCase()).toContain("#ff00ff");

    // Filled title + still-empty body (filtered) + decorative = 2 groups.
    const groupCount = (svg.match(/<g\b[^>]*transform="translate/g) || []).length;
    expect(groupCount).toBe(2);
  });
});

describe("presentation.noSlides warning", () => {
  async function createEmptyPptx(): Promise<Buffer> {
    const zip = new JSZip();

    const emptyContentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;

    const emptyPresentationXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldSz cx="9144000" cy="5143500" type="screen16x9"/>
</p:presentation>`;

    const emptyPresentationRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`;

    zip.file("[Content_Types].xml", emptyContentTypes);
    zip.file("_rels/.rels", rootRels);
    zip.file("ppt/presentation.xml", emptyPresentationXml);
    zip.file("ppt/_rels/presentation.xml.rels", emptyPresentationRels);
    zip.file("ppt/slideMasters/slideMaster1.xml", slideMaster1);
    zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMaster1Rels);
    zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayout1);
    zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayout1Rels);
    zip.file("ppt/theme/theme1.xml", theme1);
    return zip.generateAsync({ type: "nodebuffer" });
  }

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("returns empty array for PPTX with no slides", async () => {
    const pptx = await createEmptyPptx();
    const report = await convertPptxToSvg(pptx);

    expect(report.slides).toHaveLength(0);
    expect(report.supportCoverage.overall).toMatchObject({
      inputElements: 0,
      outputElements: 0,
    });
  });

  it('collects renderer warnings in diagnostics even when logLevel is "off"', async () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const pptx = await createEmptyPptx();
    const report = await convertPptxToSvg(pptx, { logLevel: "off" });

    expect(warnSpy).not.toHaveBeenCalled();
    expect(report.diagnostics).toContainEqual(
      expect.objectContaining({
        source: "renderer",
        severity: "warning",
        code: "renderer.presentation.noSlides",
      }),
    );
    expect(report.supportCoverage.overall.warnings).toBe(1);
  });

  it('restores the warning logger when logLevel is "off" and conversion throws', async () => {
    await expect(convertPptxToSvg(Buffer.from("not a pptx"), { logLevel: "off" })).rejects.toThrow(
      /zip|pptx|invalid/i,
    );

    warn("test.afterFailure", "should not be collected");
    expect(getWarningEntries()).toHaveLength(0);
  });

  it("emits presentation.noSlides warning for empty PPTX", async () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const pptx = await createEmptyPptx();
    await convertPptxToSvg(pptx, { logLevel: "warn" });

    const calls = warnSpy.mock.calls.map((c) => String(c[0]));
    expect(calls.some((msg) => msg.includes("presentation.noSlides"))).toBe(true);
  });

  it("does not emit presentation.noSlides when slides exist but filter matches none", async () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    await convertPptxToSvg(testPptx, { slides: [99], logLevel: "warn" });

    const calls = warnSpy.mock.calls.map((c) => String(c[0]));
    expect(calls.some((msg) => msg.includes("presentation.noSlides"))).toBe(false);
  });
});
