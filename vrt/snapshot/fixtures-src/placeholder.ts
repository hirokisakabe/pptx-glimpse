import {
  buildPptx,
  NS,
  outlineXml,
  REL_TYPES,
  savePptx,
  shapeXml,
  slideRelsXml,
  solidFillXml,
  textBodyXmlHelper,
  wrapSlideXml,
} from "../fixture-builder.js";

async function createPlaceholderOverlapFixture(): Promise<void> {
  // Custom slide master with decorative shapes (red rect + blue ellipse)
  const masterWithShapes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      ${shapeXml(2, "MasterRect", {
        preset: "rect",
        x: 200000,
        y: 200000,
        cx: 4000000,
        cy: 2500000,
        fillXml: solidFillXml("CC3333"),
        textBodyXml: textBodyXmlHelper("Master Rect", {
          fontSize: 16,
          bold: true,
          color: "FFFFFF",
        }),
      })}
      ${shapeXml(3, "MasterEllipse", {
        preset: "ellipse",
        x: 4800000,
        y: 200000,
        cx: 4000000,
        cy: 2500000,
        fillXml: solidFillXml("3366CC"),
        textBodyXml: textBodyXmlHelper("Master Ellipse", {
          fontSize: 16,
          bold: true,
          color: "FFFFFF",
        }),
      })}
      ${shapeXml(4, "MasterFooter", {
        preset: "rect",
        x: 200000,
        y: 4200000,
        cx: 8600000,
        cy: 600000,
        fillXml: solidFillXml("333333"),
        textBodyXml: textBodyXmlHelper("Master Footer Bar", { fontSize: 12, color: "AAAAAA" }),
      })}
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`;

  // Slide 1: Master shapes visible behind slide shapes (default showMasterSp=true)
  const slideShapes1 = [
    // Semi-transparent shape overlapping master rect
    shapeXml(2, "SlideOverlap1", {
      preset: "roundRect",
      x: 1500000,
      y: 800000,
      cx: 3500000,
      cy: 2000000,
      fillXml: `<a:solidFill><a:srgbClr val="70AD47"><a:alpha val="70000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(25400, "2E7D32"),
      textBodyXml: textBodyXmlHelper("Slide Shape 1", {
        fontSize: 14,
        bold: true,
        color: "FFFFFF",
      }),
    }),
    // Shape overlapping master ellipse
    shapeXml(3, "SlideOverlap2", {
      preset: "diamond",
      x: 5000000,
      y: 1000000,
      cx: 3000000,
      cy: 2500000,
      fillXml: `<a:solidFill><a:srgbClr val="FFC000"><a:alpha val="70000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(25400, "CC9900"),
      textBodyXml: textBodyXmlHelper("Slide Shape 2", {
        fontSize: 14,
        bold: true,
        color: "333333",
      }),
    }),
  ].join("\n");

  const slide1 = wrapSlideXml(slideShapes1);
  const rels1 = slideRelsXml();

  // Slide 2: showMasterSp="0" - master shapes should be hidden
  const slideShapes2 = [
    shapeXml(2, "SlideOnly", {
      preset: "rect",
      x: 1000000,
      y: 1000000,
      cx: 7000000,
      cy: 3000000,
      fillXml: solidFillXml("4472C4"),
      outlineXml: outlineXml(25400, "2F5496"),
      textBodyXml: textBodyXmlHelper("Only slide shapes (showMasterSp=0)", {
        fontSize: 16,
        bold: true,
        color: "FFFFFF",
      }),
    }),
  ].join("\n");

  const slide2Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" showMasterSp="0">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
      ${slideShapes2}
    </p:spTree>
  </p:cSld>
</p:sld>`;
  const rels2 = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels: rels1 },
      { xml: slide2Xml, rels: rels2 },
    ],
    slideMasterXml: masterWithShapes,
  });
  savePptx(buffer, "placeholder-overlap.pptx");
}

// --- Paragraph Spacing ---

async function createPlaceholderInheritanceExtendedFixture(): Promise<void> {
  // slideMaster with txStyles: titleStyle(36pt), bodyStyle lvl1-5, otherStyle(14pt)
  const customSlideMaster = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle>
      <a:lvl1pPr><a:defRPr sz="3600"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:lvl1pPr>
    </p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr><a:defRPr sz="2400"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:lvl1pPr>
      <a:lvl2pPr><a:defRPr sz="2200"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:lvl2pPr>
      <a:lvl3pPr><a:defRPr sz="2000"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:lvl3pPr>
      <a:lvl4pPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:lvl4pPr>
      <a:lvl5pPr><a:defRPr sz="1600"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:lvl5pPr>
    </p:bodyStyle>
    <p:otherStyle>
      <a:lvl1pPr><a:defRPr sz="1400"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:lvl1pPr>
    </p:otherStyle>
  </p:txStyles>
</p:sldMaster>`;

  const slideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_TYPES.slideLayout}" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

  // Slide 1: body placeholder with lvl 0-4 (5 levels)
  const bodyLevels = `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Body Levels"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="4594860"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="t"/>
    <a:p><a:pPr lvl="0"/><a:r><a:t>Level 1 (24pt)</a:t></a:r></a:p>
    <a:p><a:pPr lvl="1"/><a:r><a:t>Level 2 (22pt)</a:t></a:r></a:p>
    <a:p><a:pPr lvl="2"/><a:r><a:t>Level 3 (20pt)</a:t></a:r></a:p>
    <a:p><a:pPr lvl="3"/><a:r><a:t>Level 4 (18pt)</a:t></a:r></a:p>
    <a:p><a:pPr lvl="4"/><a:r><a:t>Level 5 (16pt)</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`;

  const slide1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="333333"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      ${bodyLevels}
    </p:spTree>
  </p:cSld>
</p:sld>`;

  // Slide 2: ctrTitle and subTitle placeholders
  const ctrTitle = `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Center Title"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="7772400" cy="1470025"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="b"/>
    <a:p><a:r><a:t>ctrTitle (36pt from titleStyle)</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`;

  const subTitle = `<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="Subtitle"/><p:cNvSpPr/><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="685800" y="2743200"/><a:ext cx="7772400" cy="1143000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="t"/>
    <a:p><a:r><a:t>subTitle (24pt from bodyStyle)</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`;

  const slide2Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="333333"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      ${ctrTitle}
      ${subTitle}
    </p:spTree>
  </p:cSld>
</p:sld>`;

  const buffer = await buildPptx({
    slides: [
      { xml: slide1Xml, rels: slideRels },
      { xml: slide2Xml, rels: slideRels },
    ],
    slideMasterXml: customSlideMaster,
  });
  savePptx(buffer, "placeholder-inheritance-extended.pptx");
}

// --- Table with table style border (no inline borders) ---

async function createPlaceholderGeometryInheritanceFixture(): Promise<void> {
  // Define placeholder position and size in slide layout
  const layoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" type="obj">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="685800" y="1600200"/>
            <a:ext cx="7772400" cy="1371600"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="b"/>
          <a:lstStyle/>
          <a:p><a:endParaRPr lang="en-US"/></a:p>
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
            <a:off x="1371600" y="3200400"/>
            <a:ext cx="6400800" cy="914400"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:endParaRPr lang="en-US"/></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;

  // Slide placeholder shapes have empty spPr (inherited from layout)
  const slide1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:rPr lang="en-US" sz="4400" dirty="0"/><a:t>Inherited Title</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Subtitle 2"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:rPr lang="en-US" sz="2400" dirty="0"/><a:t>Inherited Subtitle</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [{ xml: slide1, rels }],
    slideLayoutXml: layoutXml,
  });
  savePptx(buffer, "placeholder-geometry-inheritance.pptx");
}

// --- Empty slide placeholders ---
// In the template PPTX, the placeholder where the user did not enter any text is
// left empty on the slide. PowerPoint hides it completely.
// Confirm that pptx-glimpse has the same behavior.

async function createPlaceholderEmptyOnSlideFixture(): Promise<void> {
  // The layout side is a normal placeholder definition with geometry (inherited on the slide side).
  const layoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" type="obj">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="685800" y="457200"/><a:ext cx="7772400" cy="1143000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p><a:endParaRPr lang="en-US"/></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Body 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="685800" y="1828800"/><a:ext cx="7772400" cy="3200400"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="t"/>
          <a:lstStyle/>
          <a:p><a:endParaRPr lang="en-US"/></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;

  // Slide 1: Title has text entered, body remains empty (= should be hidden).
  // In addition, place one decorative non-placeholder shape and make sure it is drawn.
  const slide1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:solidFill><a:srgbClr val="C8E6C9"/></a:solidFill>
          <a:ln w="25400"><a:solidFill><a:srgbClr val="2E7D32"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:r><a:rPr lang="en-US" sz="3200" b="1"><a:solidFill><a:srgbClr val="1B5E20"/></a:solidFill></a:rPr><a:t>Filled Title</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Body 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:solidFill><a:srgbClr val="FFCDD2"/></a:solidFill>
          <a:ln w="25400"><a:solidFill><a:srgbClr val="C62828"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:endParaRPr lang="en-US"/></a:p>
        </p:txBody>
      </p:sp>
      ${shapeXml(4, "Decorative", {
        preset: "ellipse",
        x: 6858000,
        y: 4114800,
        cx: 1828800,
        cy: 914400,
        fillXml: solidFillXml("FFC107"),
        outlineXml: outlineXml(12700, "FF6F00"),
      })}
    </p:spTree>
  </p:cSld>
</p:sld>`;

  const rels = slideRelsXml();
  const buffer = await buildPptx({
    slides: [{ xml: slide1Xml, rels }],
    slideLayoutXml: layoutXml,
  });
  savePptx(buffer, "placeholder-empty-on-slide.pptx");
}

export const placeholderFixtureCreators: FixtureCreatorMap = {
  "placeholder-overlap.pptx": createPlaceholderOverlapFixture,
  "placeholder-inheritance-extended.pptx": createPlaceholderInheritanceExtendedFixture,
  "placeholder-geometry-inheritance.pptx": createPlaceholderGeometryInheritanceFixture,
  "placeholder-empty-on-slide.pptx": createPlaceholderEmptyOnSlideFixture,
};
