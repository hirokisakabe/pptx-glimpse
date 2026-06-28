import JSZip from "jszip";
import { beforeAll, describe, expect, it } from "vitest";

import { collectUsedFonts } from "./font-collector.js";

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
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Shape 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="274638"/>
            <a:ext cx="3048000" cy="1143000"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US">
                <a:latin typeface="Arial"/>
                <a:ea typeface="MS PGothic"/>
                <a:cs typeface="Arial"/>
              </a:rPr>
              <a:t>Hello</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Shape 2"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="4572000" y="274638"/>
            <a:ext cx="3048000" cy="1143000"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="ja-JP">
                <a:latin typeface="Times New Roman"/>
                <a:ea typeface="Yu Gothic"/>
              </a:rPr>
              <a:t>World</a:t>
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
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface="MS PGothic"/>
        <a:cs typeface="Times New Roman"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface="MS PGothic"/>
        <a:cs typeface="Arial"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`;

const slideWithThemeAliases = slide1Xml
  .replace(
    `<a:p>
            <a:r>
              <a:rPr lang="en-US">
                <a:latin typeface="Arial"/>
                <a:ea typeface="MS PGothic"/>
                <a:cs typeface="Arial"/>
              </a:rPr>`,
    `<a:p>
            <a:pPr><a:buFont typeface="+mj-lt"/></a:pPr>
            <a:r>
              <a:rPr lang="en-US">
                <a:latin typeface="+mn-lt"/>
                <a:ea typeface="+mn-ea"/>
                <a:cs typeface="+mn-cs"/>
              </a:rPr>`,
  )
  .replace(
    `<a:rPr lang="ja-JP">
                <a:latin typeface="Times New Roman"/>
                <a:ea typeface="Yu Gothic"/>
              </a:rPr>`,
    `<a:rPr lang="ja-JP">
                <a:latin typeface="+mj-lt"/>
                <a:ea typeface="+mj-ea"/>
                <a:cs typeface="+mj-cs"/>
              </a:rPr>`,
  );

const themeWithoutRegionalFonts = theme1
  .replaceAll('\n        <a:ea typeface="MS PGothic"/>', "")
  .replace('\n        <a:cs typeface="Times New Roman"/>', "")
  .replace('\n        <a:cs typeface="Arial"/>', "");

const contentTypesWithSecondSlide = contentTypes.replace(
  "</Types>",
  `  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`,
);

const presentationWithTwoSlidesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId r:id="rId1"/>
    <p:sldMasterId r:id="rId4"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="257" r:id="rId5"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500" type="screen16x9"/>
</p:presentation>`;

const presentationWithTwoSlidesRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>`;

const slide2Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
</Relationships>`;

const slideMaster2Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme2.xml"/>
</Relationships>`;

const slideLayout2Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster2.xml"/>
</Relationships>`;

const theme2 = theme1
  .replace("Calibri Light", "Contoso Heading")
  .replace("Calibri", "Contoso Body")
  .replaceAll("MS PGothic", "MS Mincho")
  .replace("Times New Roman", "Contoso Complex")
  .replace("Arial", "Contoso Minor Complex");

async function createTestPptx({
  contentTypesXml = contentTypes,
  presentationXmlOverride = presentationXml,
  presentationRelsXml = presentationRels,
  slideXml = slide1Xml,
  themeXml = theme1,
  extraFiles = {},
}: {
  contentTypesXml?: string;
  presentationXmlOverride?: string;
  presentationRelsXml?: string;
  slideXml?: string;
  themeXml?: string;
  extraFiles?: Record<string, string>;
} = {}): Promise<Buffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", rootRels);
  zip.file("ppt/presentation.xml", presentationXmlOverride);
  zip.file("ppt/_rels/presentation.xml.rels", presentationRelsXml);
  zip.file("ppt/slides/slide1.xml", slideXml);
  zip.file("ppt/slides/_rels/slide1.xml.rels", slide1Rels);
  zip.file("ppt/slideMasters/slideMaster1.xml", slideMaster1);
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMaster1Rels);
  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayout1);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayout1Rels);
  zip.file("ppt/theme/theme1.xml", themeXml);
  for (const [path, content] of Object.entries(extraFiles)) {
    zip.file(path, content);
  }
  return zip.generateAsync({ type: "nodebuffer" });
}

let testPptx: Buffer;

beforeAll(async () => {
  testPptx = await createTestPptx();
});

describe("collectUsedFonts", () => {
  it("Return theme font information", () => {
    const result = collectUsedFonts(testPptx);

    expect(result.theme.majorFont).toBe("Calibri Light");
    expect(result.theme.minorFont).toBe("Calibri");
    expect(result.theme.majorFontEa).toBe("MS PGothic");
    expect(result.theme.minorFontEa).toBe("MS PGothic");
    expect(result.theme.majorFontCs).toBe("Times New Roman");
    expect(result.theme.minorFontCs).toBe("Arial");
  });

  it("Collect font names for text runs", () => {
    const result = collectUsedFonts(testPptx);

    // font specified in text run
    expect(result.fonts).toContain("Arial");
    expect(result.fonts).toContain("MS PGothic");
    expect(result.fonts).toContain("Times New Roman");
    expect(result.fonts).toContain("Yu Gothic");
  });

  it("Theme font names are also included in the fonts list.", () => {
    const result = collectUsedFonts(testPptx);

    expect(result.fonts).toContain("Calibri Light");
    expect(result.fonts).toContain("Calibri");
  });

  it("Resolve and collect theme font aliases", async () => {
    const pptx = await createTestPptx({ slideXml: slideWithThemeAliases });
    const result = collectUsedFonts(pptx);

    expect(result.fonts).toContain("Calibri Light");
    expect(result.fonts).toContain("Calibri");
    expect(result.fonts).toContain("MS PGothic");
    expect(result.fonts).toContain("Times New Roman");
    expect(result.fonts).toContain("Arial");
    expect(result.fonts).not.toContain("+mj-lt");
    expect(result.fonts).not.toContain("+mn-lt");
    expect(result.fonts).not.toContain("+mj-ea");
    expect(result.fonts).not.toContain("+mn-ea");
    expect(result.fonts).not.toContain("+mj-cs");
    expect(result.fonts).not.toContain("+mn-cs");
  });

  it("Do not leak unresolvable theme regional aliases to the fonts list", async () => {
    const pptx = await createTestPptx({
      slideXml: slideWithThemeAliases,
      themeXml: themeWithoutRegionalFonts,
    });
    const result = collectUsedFonts(pptx);

    expect(result.theme.majorFontEa).toBeNull();
    expect(result.theme.minorFontEa).toBeNull();
    expect(result.theme.majorFontCs).toBeNull();
    expect(result.theme.minorFontCs).toBeNull();
    expect(result.fonts).toContain("Calibri Light");
    expect(result.fonts).toContain("Calibri");
    expect(result.fonts).not.toContain("+mj-ea");
    expect(result.fonts).not.toContain("+mn-ea");
    expect(result.fonts).not.toContain("+mj-cs");
    expect(result.fonts).not.toContain("+mn-cs");
  });

  it("Resolve per-slide theme font alias with each slide's theme", async () => {
    const pptx = await createTestPptx({
      contentTypesXml: contentTypesWithSecondSlide,
      presentationXmlOverride: presentationWithTwoSlidesXml,
      presentationRelsXml: presentationWithTwoSlidesRels,
      extraFiles: {
        "ppt/slides/slide2.xml": slideWithThemeAliases,
        "ppt/slides/_rels/slide2.xml.rels": slide2Rels,
        "ppt/slideMasters/slideMaster2.xml": slideMaster1,
        "ppt/slideMasters/_rels/slideMaster2.xml.rels": slideMaster2Rels,
        "ppt/slideLayouts/slideLayout2.xml": slideLayout1,
        "ppt/slideLayouts/_rels/slideLayout2.xml.rels": slideLayout2Rels,
        "ppt/theme/theme2.xml": theme2,
      },
    });
    const result = collectUsedFonts(pptx);

    expect(result.fonts).toContain("Contoso Heading");
    expect(result.fonts).toContain("Contoso Body");
    expect(result.fonts).toContain("MS Mincho");
    expect(result.fonts).toContain("Contoso Complex");
    expect(result.fonts).toContain("Contoso Minor Complex");
    expect(result.fonts).not.toContain("+mj-lt");
    expect(result.fonts).not.toContain("+mn-lt");
    expect(result.fonts).not.toContain("+mj-ea");
    expect(result.fonts).not.toContain("+mn-ea");
    expect(result.fonts).not.toContain("+mj-cs");
    expect(result.fonts).not.toContain("+mn-cs");
  });

  it("Font names are sorted without duplicates", () => {
    const result = collectUsedFonts(testPptx);

    // Make sure there are no duplicates
    const unique = [...new Set(result.fonts)];
    expect(result.fonts).toEqual(unique);

    // Make sure it's sorted
    const sorted = [...result.fonts].sort();
    expect(result.fonts).toEqual(sorted);
  });
});
