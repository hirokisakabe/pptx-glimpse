import {
  buildPptx,
  gradientFillXml,
  gridPosition,
  NS,
  outlineXml,
  REL_TYPES,
  savePptx,
  shapeXml,
  SLIDE_H,
  SLIDE_W,
  slideRelsXml,
  solidFillXml,
  textBodyXmlHelper,
  wrapSlideXml,
} from "../fixture-builder.js";

async function createCompositeFixture(): Promise<void> {
  // Slide 1: Non-rectangular shapes with text and different anchors
  let id = 2;
  const shapes1: string[] = [];

  const shapeTextTests = [
    { preset: "ellipse", label: "Ellipse Top", anchor: "t", fill: "4472C4" },
    { preset: "ellipse", label: "Ellipse Center", anchor: "ctr", fill: "5B9BD5" },
    { preset: "ellipse", label: "Ellipse Bottom", anchor: "b", fill: "2E75B6" },
    { preset: "diamond", label: "Diamond\nText", anchor: "ctr", fill: "ED7D31" },
    { preset: "triangle", label: "Tri", anchor: "ctr", fill: "70AD47" },
    {
      preset: "roundRect",
      label: "Multi-line\nRounded Rect\nwith Text",
      anchor: "ctr",
      fill: "FFC000",
    },
  ];
  shapeTextTests.forEach((t, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const pos = gridPosition(col, row, 3, 2);
    shapes1.push(
      shapeXml(id++, t.label, {
        preset: t.preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(t.fill),
        outlineXml: outlineXml(12700, "333333"),
        textBodyXml: textBodyXmlHelper(t.label, {
          fontSize: 14,
          color: "FFFFFF",
          anchor: t.anchor,
          align: "ctr",
        }),
      }),
    );
  });
  const slide1 = wrapSlideXml(shapes1.join("\n"));

  // Slide 2: Shape + Fill + Text + Outline combinations
  id = 2;
  const shapes2: string[] = [];

  // Gradient fill + text
  const pos2a = gridPosition(0, 0, 3, 2);
  shapes2.push(
    shapeXml(id++, "grad-text", {
      preset: "roundRect",
      x: pos2a.x,
      y: pos2a.y,
      cx: pos2a.w,
      cy: pos2a.h,
      fillXml: gradientFillXml(
        [
          { pos: 0, color: "1A1A2E" },
          { pos: 100000, color: "E94560" },
        ],
        5400000,
      ),
      outlineXml: outlineXml(12700, "333333"),
      textBodyXml: textBodyXmlHelper("Gradient Fill", {
        fontSize: 18,
        bold: true,
        color: "FFFFFF",
      }),
    }),
  );

  // Thick outline + text
  const pos2b = gridPosition(1, 0, 3, 2);
  shapes2.push(
    shapeXml(id++, "thick-outline-text", {
      preset: "rect",
      x: pos2b.x,
      y: pos2b.y,
      cx: pos2b.w,
      cy: pos2b.h,
      fillXml: solidFillXml("F0F0F0"),
      outlineXml: outlineXml(76200, "4472C4"),
      textBodyXml: textBodyXmlHelper("Thick Outline", {
        fontSize: 16,
        color: "333333",
      }),
    }),
  );

  // Semi-transparent fill + text
  const pos2c = gridPosition(2, 0, 3, 2);
  shapes2.push(
    shapeXml(id++, "semi-transparent-text", {
      preset: "roundRect",
      x: pos2c.x,
      y: pos2c.y,
      cx: pos2c.w,
      cy: pos2c.h,
      fillXml: `<a:solidFill><a:srgbClr val="4472C4"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(25400, "333333"),
      textBodyXml: textBodyXmlHelper("50% Alpha Fill", {
        fontSize: 16,
        bold: true,
        color: "000000",
      }),
    }),
  );

  // Gradient fill ellipse + italic text
  const pos2d = gridPosition(0, 1, 3, 2);
  shapes2.push(
    shapeXml(id++, "grad-ellipse-text", {
      preset: "ellipse",
      x: pos2d.x,
      y: pos2d.y,
      cx: pos2d.w,
      cy: pos2d.h,
      fillXml: gradientFillXml(
        [
          { pos: 0, color: "FF6384" },
          { pos: 50000, color: "FFCE56" },
          { pos: 100000, color: "36A2EB" },
        ],
        2700000,
      ),
      outlineXml: outlineXml(19050, "666666"),
      textBodyXml: textBodyXmlHelper("Rainbow Ellipse", {
        fontSize: 14,
        italic: true,
        color: "FFFFFF",
      }),
    }),
  );

  // Dashed outline + bold text
  const pos2e = gridPosition(1, 1, 3, 2);
  shapes2.push(
    shapeXml(id++, "dashed-bold", {
      preset: "diamond",
      x: pos2e.x,
      y: pos2e.y,
      cx: pos2e.w,
      cy: pos2e.h,
      fillXml: solidFillXml("E7E6E6"),
      outlineXml: outlineXml(38100, "ED7D31", "dash"),
      textBodyXml: textBodyXmlHelper("Dashed", {
        fontSize: 14,
        bold: true,
        color: "ED7D31",
      }),
    }),
  );

  // Solid fill hexagon + underline text
  const pos2f = gridPosition(2, 1, 3, 2);
  shapes2.push(
    shapeXml(id++, "hex-underline", {
      preset: "hexagon",
      x: pos2f.x,
      y: pos2f.y,
      cx: pos2f.w,
      cy: pos2f.h,
      fillXml: solidFillXml("70AD47"),
      outlineXml: outlineXml(25400, "333333"),
      textBodyXml: textBodyXmlHelper("Hexagon", {
        fontSize: 14,
        underline: true,
        color: "FFFFFF",
      }),
    }),
  );

  const slide2 = wrapSlideXml(shapes2.join("\n"));

  // Slide 3: Transform + Text combinations
  id = 2;
  const shapes3: string[] = [];

  const transformTextTests = [
    { label: "Rot45 + Text", preset: "rect", rotation: 45, fill: "4472C4" },
    {
      label: "Rot30 + Grad",
      preset: "roundRect",
      rotation: 30,
      fill: undefined as string | undefined,
      useGrad: true,
    },
    { label: "FlipH + Text", preset: "rightArrow", flipH: true, fill: "ED7D31" },
    { label: "FlipV + Text", preset: "triangle", flipV: true, fill: "70AD47" },
    { label: "Rot+FlipH", preset: "parallelogram", rotation: 15, flipH: true, fill: "FFC000" },
    { label: "Rot+FlipV", preset: "pentagon", rotation: 60, flipV: true, fill: "9966FF" },
  ];
  transformTextTests.forEach((t, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const pos = gridPosition(col, row, 3, 2);
    const fill = t.useGrad
      ? gradientFillXml(
          [
            { pos: 0, color: "667EEA" },
            { pos: 100000, color: "764BA2" },
          ],
          5400000,
        )
      : solidFillXml(t.fill!);
    shapes3.push(
      shapeXml(id++, t.label, {
        preset: t.preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: fill,
        outlineXml: outlineXml(12700, "333333"),
        rotation: t.rotation,
        flipH: t.flipH,
        flipV: t.flipV,
        textBodyXml: textBodyXmlHelper(t.label, {
          fontSize: 12,
          bold: true,
          color: "FFFFFF",
        }),
      }),
    );
  });
  const slide3 = wrapSlideXml(shapes3.join("\n"));

  // Slide 4: Overlapping shapes (z-order and semi-transparency)
  id = 2;
  const shapes4: string[] = [];

  // Layer 1: large background rectangle
  shapes4.push(
    shapeXml(id++, "bg-rect", {
      preset: "rect",
      x: 500000,
      y: 300000,
      cx: 4000000,
      cy: 3000000,
      fillXml: solidFillXml("4472C4"),
      outlineXml: outlineXml(12700, "333333"),
      textBodyXml: textBodyXmlHelper("Back (Z=1)", {
        fontSize: 14,
        color: "FFFFFF",
        anchor: "t",
        align: "l",
      }),
    }),
  );

  // Layer 2: overlapping ellipse
  shapes4.push(
    shapeXml(id++, "mid-ellipse", {
      preset: "ellipse",
      x: 1500000,
      y: 1000000,
      cx: 3500000,
      cy: 2500000,
      fillXml: solidFillXml("ED7D31"),
      outlineXml: outlineXml(12700, "333333"),
      textBodyXml: textBodyXmlHelper("Middle (Z=2)", {
        fontSize: 14,
        color: "FFFFFF",
      }),
    }),
  );

  // Layer 3: small foreground rounded rect
  shapes4.push(
    shapeXml(id++, "front-roundRect", {
      preset: "roundRect",
      x: 2500000,
      y: 1500000,
      cx: 2500000,
      cy: 1800000,
      fillXml: solidFillXml("70AD47"),
      outlineXml: outlineXml(19050, "333333"),
      textBodyXml: textBodyXmlHelper("Front (Z=3)", {
        fontSize: 16,
        bold: true,
        color: "FFFFFF",
      }),
    }),
  );

  // Semi-transparent overlapping group (right side)
  shapes4.push(
    shapeXml(id++, "alpha-rect1", {
      preset: "rect",
      x: 5500000,
      y: 500000,
      cx: 2500000,
      cy: 2500000,
      fillXml: `<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(12700, "990000"),
      textBodyXml: textBodyXmlHelper("Red 50%", {
        fontSize: 12,
        color: "000000",
        anchor: "t",
      }),
    }),
  );

  shapes4.push(
    shapeXml(id++, "alpha-rect2", {
      preset: "rect",
      x: 6200000,
      y: 1200000,
      cx: 2500000,
      cy: 2500000,
      fillXml: `<a:solidFill><a:srgbClr val="0000FF"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(12700, "000099"),
      textBodyXml: textBodyXmlHelper("Blue 50%", {
        fontSize: 12,
        color: "000000",
        anchor: "b",
      }),
    }),
  );

  shapes4.push(
    shapeXml(id++, "alpha-ellipse", {
      preset: "ellipse",
      x: 5800000,
      y: 2000000,
      cx: 2200000,
      cy: 2200000,
      fillXml: `<a:solidFill><a:srgbClr val="00FF00"><a:alpha val="40000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(12700, "006600"),
      textBodyXml: textBodyXmlHelper("Green 40%", {
        fontSize: 12,
        color: "000000",
      }),
    }),
  );

  const slide4 = wrapSlideXml(shapes4.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
      { xml: slide3, rels },
      { xml: slide4, rels },
    ],
  });
  savePptx(buffer, "composite.pptx");
}

// --- 20. Text Decoration (superscript / subscript) ---

async function createHyperlinksFixture(): Promise<void> {
  const margin = 457200; // 0.5 inch
  const shapeW = 8229600; // 8.5 inch
  const shapeH = 914400; // 1 inch

  // Shape 1: Normal text (no hyperlink)
  const shape1 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Shape 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${margin}" y="${margin}"/><a:ext cx="${shapeW}" cy="${shapeH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" sz="1400">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
        </a:rPr>
        <a:t>Normal text without hyperlink</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`;

  // Shape 2: Text with hyperlink
  const shape2 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="Shape 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${margin}" y="${margin + shapeH + margin}"/><a:ext cx="${shapeW}" cy="${shapeH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" sz="1400">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
        </a:rPr>
        <a:t>Click here: </a:t>
      </a:r>
      <a:r>
        <a:rPr lang="en-US" sz="1400" u="sng">
          <a:solidFill><a:srgbClr val="0563C1"/></a:solidFill>
          <a:hlinkClick r:id="rId2"/>
        </a:rPr>
        <a:t>https://example.com</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`;

  // Shape 3: Text with hyperlink and tooltip
  const shape3 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="4" name="Shape 3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${margin}" y="${margin + (shapeH + margin) * 2}"/><a:ext cx="${shapeW}" cy="${shapeH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" sz="1400" u="sng">
          <a:solidFill><a:srgbClr val="0563C1"/></a:solidFill>
          <a:hlinkClick r:id="rId3" tooltip="Visit Example"/>
        </a:rPr>
        <a:t>Link with tooltip</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`;

  // Shape 4: Multiple hyperlinks in one paragraph
  const shape4 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="5" name="Shape 4"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${margin}" y="${margin + (shapeH + margin) * 3}"/><a:ext cx="${shapeW}" cy="${shapeH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" sz="1400" u="sng">
          <a:solidFill><a:srgbClr val="0563C1"/></a:solidFill>
          <a:hlinkClick r:id="rId2"/>
        </a:rPr>
        <a:t>Link 1</a:t>
      </a:r>
      <a:r>
        <a:rPr lang="en-US" sz="1400">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
        </a:rPr>
        <a:t> and </a:t>
      </a:r>
      <a:r>
        <a:rPr lang="en-US" sz="1400" u="sng">
          <a:solidFill><a:srgbClr val="0563C1"/></a:solidFill>
          <a:hlinkClick r:id="rId3"/>
        </a:rPr>
        <a:t>Link 2</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`;

  const slide = wrapSlideXml([shape1, shape2, shape3, shape4].join("\n"));
  const rels = slideRelsXml([
    {
      id: "rId2",
      type: REL_TYPES.hyperlink,
      target: "https://example.com",
      targetMode: "External",
    },
    {
      id: "rId3",
      type: REL_TYPES.hyperlink,
      target: "https://example.org",
      targetMode: "External",
    },
  ]);

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "hyperlinks.pptx");
}

// --- Pattern / Image Fill / Radial Gradient ---

async function createSmartArtFixture(): Promise<void> {
  const margin = 300000;
  const diagramW = SLIDE_W - margin * 2;
  const diagramH = SLIDE_H - margin * 2;

  // SmartArt drawing XML: process type with 3 rounded rectangles and 2 arrows placed horizontally
  const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
             xmlns:a="${NS.a}">
  <dsp:spTree>
    <dsp:nvGrpSpPr>
      <dsp:cNvPr id="0" name=""/>
      <dsp:cNvGrpSpPr/>
    </dsp:nvGrpSpPr>
    <dsp:grpSpPr>
      <a:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="${diagramW}" cy="${diagramH}"/>
        <a:chOff x="0" y="0"/>
        <a:chExt cx="${diagramW}" cy="${diagramH}"/>
      </a:xfrm>
    </dsp:grpSpPr>
    <dsp:sp>
      <dsp:nvSpPr><dsp:cNvPr id="10" name="Box1"/><dsp:cNvSpPr/></dsp:nvSpPr>
      <dsp:spPr>
        <a:xfrm><a:off x="0" y="500000"/><a:ext cx="2500000" cy="3500000"/></a:xfrm>
        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
      </dsp:spPr>
      <dsp:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Step 1</a:t></a:r></a:p>
      </dsp:txBody>
    </dsp:sp>
    <dsp:sp>
      <dsp:nvSpPr><dsp:cNvPr id="11" name="Arrow1"/><dsp:cNvSpPr/></dsp:nvSpPr>
      <dsp:spPr>
        <a:xfrm><a:off x="2700000" y="1500000"/><a:ext cx="600000" cy="600000"/></a:xfrm>
        <a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="A5A5A5"/></a:solidFill>
      </dsp:spPr>
    </dsp:sp>
    <dsp:sp>
      <dsp:nvSpPr><dsp:cNvPr id="12" name="Box2"/><dsp:cNvSpPr/></dsp:nvSpPr>
      <dsp:spPr>
        <a:xfrm><a:off x="3500000" y="500000"/><a:ext cx="2500000" cy="3500000"/></a:xfrm>
        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>
      </dsp:spPr>
      <dsp:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Step 2</a:t></a:r></a:p>
      </dsp:txBody>
    </dsp:sp>
    <dsp:sp>
      <dsp:nvSpPr><dsp:cNvPr id="13" name="Arrow2"/><dsp:cNvSpPr/></dsp:nvSpPr>
      <dsp:spPr>
        <a:xfrm><a:off x="6200000" y="1500000"/><a:ext cx="600000" cy="600000"/></a:xfrm>
        <a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="A5A5A5"/></a:solidFill>
      </dsp:spPr>
    </dsp:sp>
    <dsp:sp>
      <dsp:nvSpPr><dsp:cNvPr id="14" name="Box3"/><dsp:cNvSpPr/></dsp:nvSpPr>
      <dsp:spPr>
        <a:xfrm><a:off x="7000000" y="500000"/><a:ext cx="2500000" cy="3500000"/></a:xfrm>
        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>
      </dsp:spPr>
      <dsp:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Step 3</a:t></a:r></a:p>
      </dsp:txBody>
    </dsp:sp>
  </dsp:spTree>
</dsp:drawing>`;

  const dataXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
  <dgm:ptLst><dgm:pt modelId="0" type="doc"/></dgm:ptLst>
</dgm:dataModel>`;

  const minimalLayout = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>`;
  const minimalStyles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="${NS.a}"/>`;
  const minimalColors = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="${NS.a}"/>`;

  const dataRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="drawing1.xml"/>
</Relationships>`;

  // SmartArt graphicFrame wrapped in mc:AlternateContent
  const smartArtXml = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                                            xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
    <mc:Choice Requires="dgm">
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="2" name="SmartArt1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="${margin}" y="${margin}"/><a:ext cx="${diagramW}" cy="${diagramH}"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
            <dgm:relIds xmlns:r="${NS.r}" r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
    </mc:Choice>
    <mc:Fallback/>
  </mc:AlternateContent>`;

  const slideXml = wrapSlideXml(smartArtXml);
  const slideRels = slideRelsXml([
    {
      id: "rId2",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
      target: "../diagrams/data1.xml",
    },
    {
      id: "rId3",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
      target: "../diagrams/layout1.xml",
    },
    {
      id: "rId4",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramStyle",
      target: "../diagrams/quickStyles1.xml",
    },
    {
      id: "rId5",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
      target: "../diagrams/colors1.xml",
    },
  ]);

  const extraFiles = new Map<string, string>();
  extraFiles.set("ppt/diagrams/data1.xml", dataXml);
  extraFiles.set("ppt/diagrams/drawing1.xml", drawingXml);
  extraFiles.set("ppt/diagrams/layout1.xml", minimalLayout);
  extraFiles.set("ppt/diagrams/quickStyles1.xml", minimalStyles);
  extraFiles.set("ppt/diagrams/colors1.xml", minimalColors);
  extraFiles.set("ppt/diagrams/_rels/data1.xml.rels", dataRels);

  const buffer = await buildPptx({
    slides: [{ xml: slideXml, rels: slideRels }],
    charts: extraFiles,
    contentTypesExtra: [
      `<Override PartName="/ppt/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>`,
      `<Override PartName="/ppt/diagrams/drawing1.xml" ContentType="application/vnd.ms-office.drawingml.diagramDrawing+xml"/>`,
    ],
  });

  savePptx(buffer, "smartart.pptx");
}

// --- Image Crop (srcRect) ---

export const miscFixtureCreators: FixtureCreatorMap = {
  "composite.pptx": createCompositeFixture,
  "hyperlinks.pptx": createHyperlinksFixture,
  "smartart.pptx": createSmartArtFixture,
};
