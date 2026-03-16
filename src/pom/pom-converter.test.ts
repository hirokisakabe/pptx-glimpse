import JSZip from "jszip";
import { beforeAll, describe, expect, it } from "vitest";

import { convertPptxToPom } from "./index.js";

// --- Minimal PPTX XML templates ---

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

const slideMaster1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
</p:sldMaster>`;

const slideMaster1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

const slideLayout1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
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
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`;

const slide1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

async function createPptx(slideXml: string): Promise<Buffer> {
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

function makeSlideXml(spTree: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      ${spTree}
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

describe("convertPptxToPom", () => {
  describe("basic structure", () => {
    let pptx: Buffer;
    beforeAll(async () => {
      pptx = await createPptx(
        makeSlideXml(`
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Rect 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="457200" y="274638"/><a:ext cx="3048000" cy="1143000"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
            <a:ln w="25400"><a:solidFill><a:srgbClr val="2F528F"/></a:solidFill></a:ln>
          </p:spPr>
          <p:txBody>
            <a:bodyPr anchor="ctr"/>
            <a:lstStyle/>
            <a:p>
              <a:pPr algn="ctr"/>
              <a:r>
                <a:rPr lang="en-US" sz="2400" b="1">
                  <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                </a:rPr>
                <a:t>Hello World</a:t>
              </a:r>
            </a:p>
          </p:txBody>
        </p:sp>`),
      );
    });

    it("returns one PomSlide per slide", () => {
      const result = convertPptxToPom(pptx);
      expect(result).toHaveLength(1);
      expect(result[0].slideNumber).toBe(1);
    });

    it("produces XML with Layer root element and slide dimensions", () => {
      const result = convertPptxToPom(pptx);
      const xml = result[0].xml;
      expect(xml).toMatch(/^<Layer/);
      expect(xml).toContain('w="1280"');
      expect(xml).toContain('h="720"');
    });

    it("converts a shape with fill and text to pom XML", () => {
      const result = convertPptxToPom(pptx);
      const xml = result[0].xml;
      expect(xml).toContain("<Shape");
      expect(xml).toContain('shapeType="rect"');
      expect(xml).toContain('fill.color="4472C4"');
      expect(xml).toContain('bold="true"');
      expect(xml).toContain('color="FFFFFF"');
      expect(xml).toContain('alignText="center"');
      expect(xml).toContain(">Hello World</Shape>");
    });

    it("positions elements with x and y attributes", () => {
      const result = convertPptxToPom(pptx);
      const xml = result[0].xml;
      // 457200 / 9144000 * 1280 = 64
      expect(xml).toContain('x="64"');
      // 274638 / 5143500 * 720 ≈ 38.44
      expect(xml).toContain('y="38.44"');
    });

    it("supports slide number filtering", () => {
      const result = convertPptxToPom(pptx, { slides: [99] });
      expect(result).toHaveLength(0);
    });
  });

  describe("text-only shapes", () => {
    let pptx: Buffer;
    beforeAll(async () => {
      pptx = await createPptx(
        makeSlideXml(`
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </p:spPr>
          <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:r>
                <a:rPr lang="en-US" sz="1800" i="1">
                  <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
                </a:rPr>
                <a:t>Plain text box</a:t>
              </a:r>
            </a:p>
          </p:txBody>
        </p:sp>`),
      );
    });

    it("converts a text-only shape to Text element in XML", () => {
      const result = convertPptxToPom(pptx);
      const xml = result[0].xml;
      expect(xml).toContain("<Text");
      expect(xml).toContain('italic="true"');
      expect(xml).toContain('color="333333"');
      expect(xml).toContain('fontPx="24"'); // 18pt * 4/3
      expect(xml).toContain(">Plain text box</Text>");
    });
  });

  describe("ellipse shape", () => {
    let pptx: Buffer;
    beforeAll(async () => {
      pptx = await createPptx(
        makeSlideXml(`
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Ellipse 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="4572000" y="274638"/><a:ext cx="2286000" cy="2286000"/></a:xfrm>
            <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>
          </p:spPr>
        </p:sp>`),
      );
    });

    it("converts ellipse with correct shapeType", () => {
      const result = convertPptxToPom(pptx);
      const xml = result[0].xml;
      expect(xml).toContain("<Shape");
      expect(xml).toContain('shapeType="ellipse"');
      expect(xml).toContain('fill.color="ED7D31"');
    });
  });

  describe("table", () => {
    let pptx: Buffer;
    beforeAll(async () => {
      pptx = await createPptx(
        makeSlideXml(`
        <p:graphicFrame>
          <p:nvGraphicFramePr>
            <p:cNvPr id="4" name="Table 1"/>
            <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
            <p:nvPr/>
          </p:nvGraphicFramePr>
          <p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="7315200" cy="1828800"/></p:xfrm>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
              <a:tbl>
                <a:tblGrid>
                  <a:gridCol w="3657600"/>
                  <a:gridCol w="3657600"/>
                </a:tblGrid>
                <a:tr h="914400">
                  <a:tc>
                    <a:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>Header 1</a:t></a:r></a:p>
                    </a:txBody>
                    <a:tcPr>
                      <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
                    </a:tcPr>
                  </a:tc>
                  <a:tc>
                    <a:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>Header 2</a:t></a:r></a:p>
                    </a:txBody>
                    <a:tcPr>
                      <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
                    </a:tcPr>
                  </a:tc>
                </a:tr>
                <a:tr h="914400">
                  <a:tc>
                    <a:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr lang="en-US" sz="1400"/><a:t>Cell A</a:t></a:r></a:p>
                    </a:txBody>
                    <a:tcPr/>
                  </a:tc>
                  <a:tc>
                    <a:txBody>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr lang="en-US" sz="1400"/><a:t>Cell B</a:t></a:r></a:p>
                    </a:txBody>
                    <a:tcPr/>
                  </a:tc>
                </a:tr>
              </a:tbl>
            </a:graphicData>
          </a:graphic>
        </p:graphicFrame>`),
      );
    });

    it("converts a table to pom Table XML", () => {
      const result = convertPptxToPom(pptx);
      const xml = result[0].xml;
      expect(xml).toContain("<Table");
      expect(xml).toContain("<TableColumn");
      expect(xml).toContain("<TableRow");
      expect(xml).toContain("<TableCell");
      expect(xml).toContain('bold="true"');
      expect(xml).toContain('backgroundColor="4472C4"');
      expect(xml).toContain(">Header 1</TableCell>");
      expect(xml).toContain(">Cell A</TableCell>");
      expect(xml).toContain(">Cell B</TableCell>");
    });
  });

  describe("connector", () => {
    let pptx: Buffer;
    beforeAll(async () => {
      pptx = await createPptx(
        makeSlideXml(`
        <p:cxnSp>
          <p:nvCxnSpPr>
            <p:cNvPr id="5" name="Connector 1"/>
            <p:cNvCxnSpPr/>
            <p:nvPr/>
          </p:nvCxnSpPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="914400" y="914400"/>
              <a:ext cx="4572000" cy="0"/>
            </a:xfrm>
            <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
            <a:ln w="25400">
              <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
              <a:tailEnd type="triangle"/>
            </a:ln>
          </p:spPr>
        </p:cxnSp>`),
      );
    });

    it("converts a connector to pom Line XML", () => {
      const result = convertPptxToPom(pptx);
      const xml = result[0].xml;
      expect(xml).toContain("<Line");
      expect(xml).toContain('color="FF0000"');
      expect(xml).toContain('endArrow.type="triangle"');
    });
  });

  describe("solid background", () => {
    let pptx: Buffer;
    beforeAll(async () => {
      pptx = await createPptx(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="1A1A2E"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:sld>`);
    });

    it("sets backgroundColor on the Layer element", () => {
      const result = convertPptxToPom(pptx);
      const xml = result[0].xml;
      expect(xml).toContain('backgroundColor="1A1A2E"');
    });
  });
});
