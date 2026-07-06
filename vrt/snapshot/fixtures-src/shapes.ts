import type { FixtureCreatorMap, GridPos } from "../fixture-builder.js";
import {
  buildPptx,
  COLORS,
  gradientFillXml,
  gridPosition,
  outlineXml,
  savePptx,
  shapeXml,
  SLIDE_H_4_3,
  SLIDE_W,
  SLIDE_W_4_3,
  slideRelsXml,
  solidFillXml,
  textBodyXmlHelper,
  wrapSlideXml,
} from "../fixture-builder.js";

async function createShapesFixture(): Promise<void> {
  const presets1 = [
    "rect",
    "ellipse",
    "roundRect",
    "triangle",
    "rtTriangle",
    "diamond",
    "parallelogram",
    "trapezoid",
    "pentagon",
    "hexagon",
  ];
  const presets2 = [
    "star4",
    "star5",
    "rightArrow",
    "leftArrow",
    "upArrow",
    "downArrow",
    "line",
    "cloud",
    "heart",
  ];

  const makeSlide = (presets: string[], cols: number, rows: number): string => {
    let id = 2;
    const shapes = presets.map((preset, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const pos = gridPosition(col, row, cols, rows);
      const lineXml = preset === "line" ? outlineXml(25400, "333333") : outlineXml(12700, "333333");
      return shapeXml(id++, preset, {
        preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(COLORS[i % COLORS.length]),
        outlineXml: lineXml,
      });
    });
    return wrapSlideXml(shapes.join("\n"));
  };

  const slide1 = makeSlide(presets1, 5, 2);
  const slide2 = makeSlide(presets2, 5, 2);
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "shapes.pptx");
}

async function createFillAndLinesFixture(): Promise<void> {
  // Slide 1: Fill types
  let id = 2;
  const fillShapes: string[] = [];

  // Solid fills
  const solidColors = ["FF6384", "36A2EB", "FFCE56"];
  solidColors.forEach((c, i) => {
    const pos = gridPosition(i, 0, 3, 2);
    fillShapes.push(
      shapeXml(id++, `solid-${c}`, {
        preset: "roundRect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(c),
      }),
    );
  });

  // Gradient fills
  const gradients = [
    { angle: 0, label: "horizontal" },
    { angle: 5400000, label: "vertical" },
    { angle: 2700000, label: "diagonal" },
  ];
  gradients.forEach((g, i) => {
    const pos = gridPosition(i, 1, 3, 2);
    fillShapes.push(
      shapeXml(id++, `grad-${g.label}`, {
        preset: "roundRect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: gradientFillXml(
          [
            { pos: 0, color: "4472C4" },
            { pos: 100000, color: "ED7D31" },
          ],
          g.angle,
        ),
      }),
    );
  });

  const slide1 = wrapSlideXml(fillShapes.join("\n"));

  // Slide 2: Line dash styles
  id = 2;
  const dashStyles = [
    "solid",
    "dash",
    "dot",
    "dashDot",
    "lgDash",
    "lgDashDot",
    "sysDash",
    "sysDot",
  ];
  const lineShapes = dashStyles.map((ds, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const pos = gridPosition(col, row, 4, 2);
    return shapeXml(id++, `line-${ds}`, {
      preset: "rect",
      x: pos.x,
      y: pos.y,
      cx: pos.w,
      cy: pos.h,
      fillXml: `<a:noFill/>`,
      outlineXml: outlineXml(25400, "333333", ds === "solid" ? undefined : ds),
      textBodyXml: textBodyXmlHelper(ds, { fontSize: 12, color: "333333" }),
    });
  });

  const slide2 = wrapSlideXml(lineShapes.join("\n"));

  // Slide 3: Line cap and join styles
  id = 2;
  const capJoinShapes: string[] = [];

  // Row 0: Line cap styles (flat, sq, rnd)
  const capStyles: { cap: string; label: string }[] = [
    { cap: "flat", label: "cap: flat (butt)" },
    { cap: "sq", label: "cap: sq (square)" },
    { cap: "rnd", label: "cap: rnd (round)" },
  ];
  capStyles.forEach((cs, i) => {
    const pos = gridPosition(i, 0, 3, 2);
    capJoinShapes.push(
      shapeXml(id++, `cap-${cs.cap}`, {
        preset: "rect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: `<a:noFill/>`,
        outlineXml: outlineXml(50800, "4472C4", undefined, { cap: cs.cap }),
        textBodyXml: textBodyXmlHelper(cs.label, { fontSize: 10, color: "333333" }),
      }),
    );
  });

  // Row 1: Line join styles (miter, bevel, round)
  const joinStyles: { join: string; label: string }[] = [
    { join: "miter", label: "join: miter" },
    { join: "bevel", label: "join: bevel" },
    { join: "round", label: "join: round" },
  ];
  joinStyles.forEach((js, i) => {
    const pos = gridPosition(i, 1, 3, 2);
    capJoinShapes.push(
      shapeXml(id++, `join-${js.join}`, {
        preset: "rect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: `<a:noFill/>`,
        outlineXml: outlineXml(50800, "ED7D31", undefined, { join: js.join }),
        textBodyXml: textBodyXmlHelper(js.label, { fontSize: 10, color: "333333" }),
      }),
    );
  });

  const slide3 = wrapSlideXml(capJoinShapes.join("\n"));

  // Slide 4: Gradient line fill + custom dash patterns
  id = 2;
  const gradDashShapes: string[] = [];

  // Row 0: Gradient line fills
  const gradientLines = [
    {
      label: "Linear grad stroke",
      outlineXml: `<a:ln w="38100"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill></a:ln>`,
    },
    {
      label: "Vertical grad stroke",
      outlineXml: `<a:ln w="38100"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="FF00FF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill></a:ln>`,
    },
    {
      label: "3-stop grad stroke",
      outlineXml: `<a:ln w="38100"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill></a:ln>`,
    },
  ];
  gradientLines.forEach((gl, i) => {
    const pos = gridPosition(i, 0, 3, 2);
    gradDashShapes.push(
      shapeXml(id++, `grad-line-${i}`, {
        preset: "rect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: `<a:noFill/>`,
        outlineXml: gl.outlineXml,
        textBodyXml: textBodyXmlHelper(gl.label, { fontSize: 9, color: "333333" }),
      }),
    );
  });

  // Row 1: Custom dash patterns
  const customDashLines = [
    {
      label: "custDash 3:1",
      outlineXml: `<a:ln w="25400"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:custDash><a:ds d="300000" sp="100000"/></a:custDash></a:ln>`,
    },
    {
      label: "custDash 3:1:1:1",
      outlineXml: `<a:ln w="25400"><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill><a:custDash><a:ds d="300000" sp="100000"/><a:ds d="100000" sp="100000"/></a:custDash></a:ln>`,
    },
    {
      label: "custDash 5:2:2:2",
      outlineXml: `<a:ln w="25400"><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill><a:custDash><a:ds d="500000" sp="200000"/><a:ds d="200000" sp="200000"/></a:custDash></a:ln>`,
    },
  ];
  customDashLines.forEach((cdl, i) => {
    const pos = gridPosition(i, 1, 3, 2);
    gradDashShapes.push(
      shapeXml(id++, `cust-dash-${i}`, {
        preset: "rect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: `<a:noFill/>`,
        outlineXml: cdl.outlineXml,
        textBodyXml: textBodyXmlHelper(cdl.label, { fontSize: 9, color: "333333" }),
      }),
    );
  });

  const slide4 = wrapSlideXml(gradDashShapes.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
      { xml: slide3, rels },
      { xml: slide4, rels },
    ],
  });
  savePptx(buffer, "fill-and-lines.pptx");
}

async function createTransformFixture(): Promise<void> {
  let id = 2;
  const shapes: string[] = [];
  const tests = [
    { label: "Rot 45", rotation: 45 },
    { label: "Rot 90", rotation: 90 },
    { label: "FlipH", flipH: true },
    { label: "FlipV", flipV: true },
    { label: "FlipH+V", flipH: true, flipV: true },
    { label: "Rot+FlipH", rotation: 30, flipH: true },
  ];
  tests.forEach((t, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const pos = gridPosition(col, row, 3, 2);
    shapes.push(
      shapeXml(id++, t.label, {
        preset: "rightArrow",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(COLORS[i]),
        outlineXml: outlineXml(12700, "333333"),
        rotation: t.rotation,
        flipH: t.flipH,
        flipV: t.flipV,
        textBodyXml: textBodyXmlHelper(t.label, {
          fontSize: 12,
          color: "FFFFFF",
        }),
      }),
    );
  });
  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "transform.pptx");
}

async function createGroupsFixture(): Promise<void> {
  // Slide 1: Group containing a rect, ellipse, and text shape
  const grpX = 500000;
  const grpY = 500000;
  const grpW = 8000000;
  const grpH = 4000000;

  const groupContent = `<p:grpSp>
  <p:nvGrpSpPr><p:cNvPr id="2" name="Group 1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm>
      <a:off x="${grpX}" y="${grpY}"/>
      <a:ext cx="${grpW}" cy="${grpH}"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="${grpW}" cy="${grpH}"/>
    </a:xfrm>
  </p:grpSpPr>
  ${shapeXml(3, "GroupRect", {
    preset: "rect",
    x: 0,
    y: 0,
    cx: 3500000,
    cy: 3500000,
    fillXml: solidFillXml("4472C4"),
    textBodyXml: textBodyXmlHelper("In Group", { fontSize: 18, color: "FFFFFF" }),
  })}
  ${shapeXml(4, "GroupEllipse", {
    preset: "ellipse",
    x: 4000000,
    y: 0,
    cx: 3500000,
    cy: 3500000,
    fillXml: solidFillXml("ED7D31"),
    textBodyXml: textBodyXmlHelper("Ellipse", { fontSize: 18, color: "FFFFFF" }),
  })}
</p:grpSp>`;

  const slide1 = wrapSlideXml(groupContent);
  const rels1 = slideRelsXml();

  // Slide 2: Group rotation and flip
  const grp2W = 3500000;
  const grp2H = 2000000;
  const childShapes = (id: number) => `
  ${shapeXml(id, "Rect", {
    preset: "rect",
    x: 0,
    y: 0,
    cx: 1500000,
    cy: 1500000,
    fillXml: solidFillXml("4472C4"),
  })}
  ${shapeXml(id + 1, "Ellipse", {
    preset: "ellipse",
    x: 1800000,
    y: 200000,
    cx: 1500000,
    cy: 1500000,
    fillXml: solidFillXml("ED7D31"),
  })}`;

  // Group with 90 degree rotation
  const rotatedGroup = `<p:grpSp>
  <p:nvGrpSpPr><p:cNvPr id="10" name="Rotated Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm rot="${90 * 60000}">
      <a:off x="200000" y="200000"/>
      <a:ext cx="${grp2W}" cy="${grp2H}"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="${grp2W}" cy="${grp2H}"/>
    </a:xfrm>
  </p:grpSpPr>
  ${childShapes(11)}
</p:grpSp>`;

  // Group with flipH
  const flipHGroup = `<p:grpSp>
  <p:nvGrpSpPr><p:cNvPr id="20" name="FlipH Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm flipH="1">
      <a:off x="4800000" y="200000"/>
      <a:ext cx="${grp2W}" cy="${grp2H}"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="${grp2W}" cy="${grp2H}"/>
    </a:xfrm>
  </p:grpSpPr>
  ${childShapes(21)}
</p:grpSp>`;

  // Group with flipV
  const flipVGroup = `<p:grpSp>
  <p:nvGrpSpPr><p:cNvPr id="30" name="FlipV Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm flipV="1">
      <a:off x="200000" y="2800000"/>
      <a:ext cx="${grp2W}" cy="${grp2H}"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="${grp2W}" cy="${grp2H}"/>
    </a:xfrm>
  </p:grpSpPr>
  ${childShapes(31)}
</p:grpSp>`;

  const slide2 = wrapSlideXml(`${rotatedGroup}${flipHGroup}${flipVGroup}`);
  const rels2 = slideRelsXml();

  // Slide 3: Group transform combinations (rot+flipH, rot+flipV, flipH+flipV)
  const grp3W = 3500000;
  const grp3H = 2000000;
  const comboChildShapes = (id: number, label: string) => `
  ${shapeXml(id, "Rect", {
    preset: "rect",
    x: 0,
    y: 0,
    cx: 1500000,
    cy: 1500000,
    fillXml: solidFillXml("70AD47"),
    textBodyXml: textBodyXmlHelper(label, { fontSize: 10, color: "FFFFFF" }),
  })}
  ${shapeXml(id + 1, "Ellipse", {
    preset: "ellipse",
    x: 1800000,
    y: 200000,
    cx: 1500000,
    cy: 1500000,
    fillXml: solidFillXml("FFC000"),
  })}`;

  // Group with rot45 + flipH
  const rotFlipHGroup = `<p:grpSp>
  <p:nvGrpSpPr><p:cNvPr id="40" name="Rot45+FlipH"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm rot="${45 * 60000}" flipH="1">
      <a:off x="200000" y="200000"/>
      <a:ext cx="${grp3W}" cy="${grp3H}"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="${grp3W}" cy="${grp3H}"/>
    </a:xfrm>
  </p:grpSpPr>
  ${comboChildShapes(41, "R45+FH")}
</p:grpSp>`;

  // Group with rot30 + flipV
  const rotFlipVGroup = `<p:grpSp>
  <p:nvGrpSpPr><p:cNvPr id="50" name="Rot30+FlipV"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm rot="${30 * 60000}" flipV="1">
      <a:off x="4800000" y="200000"/>
      <a:ext cx="${grp3W}" cy="${grp3H}"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="${grp3W}" cy="${grp3H}"/>
    </a:xfrm>
  </p:grpSpPr>
  ${comboChildShapes(51, "R30+FV")}
</p:grpSp>`;

  // Group with flipH + flipV (equivalent to 180 degree rotation)
  const flipHVGroup = `<p:grpSp>
  <p:nvGrpSpPr><p:cNvPr id="60" name="FlipH+FlipV"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm flipH="1" flipV="1">
      <a:off x="2500000" y="2800000"/>
      <a:ext cx="${grp3W}" cy="${grp3H}"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="${grp3W}" cy="${grp3H}"/>
    </a:xfrm>
  </p:grpSpPr>
  ${comboChildShapes(61, "FH+FV")}
</p:grpSp>`;

  const slide3 = wrapSlideXml(`${rotFlipHGroup}${rotFlipVGroup}${flipHVGroup}`);
  const rels3 = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels: rels1 },
      { xml: slide2, rels: rels2 },
      { xml: slide3, rels: rels3 },
    ],
  });
  savePptx(buffer, "groups.pptx");
}

async function createConnectorsFixture(): Promise<void> {
  // Row 1: Basic connectors (straight, dash, dot)
  const basicConnectors = [
    `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="2" name="Straight"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="300000"/><a:ext cx="2500000" cy="0"/></a:xfrm>
    <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln>
  </p:spPr>
</p:cxnSp>`,
    `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="3" name="StraightDiag"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="600000"/><a:ext cx="2500000" cy="1200000"/></a:xfrm>
    <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill><a:prstDash val="dash"/></a:ln>
  </p:spPr>
</p:cxnSp>`,
  ];

  // Row 2: Arrow endpoints
  const arrowConnectors = [
    `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="4" name="TriangleTail"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="2200000"/><a:ext cx="2500000" cy="0"/></a:xfrm>
    <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:tailEnd type="triangle" w="med" len="med"/></a:ln>
  </p:spPr>
</p:cxnSp>`,
    `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="5" name="BothArrows"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="2700000"/><a:ext cx="2500000" cy="0"/></a:xfrm>
    <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill><a:headEnd type="triangle" w="med" len="med"/><a:tailEnd type="stealth" w="lg" len="lg"/></a:ln>
  </p:spPr>
</p:cxnSp>`,
    `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="6" name="DiamondOval"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="3200000"/><a:ext cx="2500000" cy="0"/></a:xfrm>
    <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill><a:headEnd type="diamond" w="med" len="med"/><a:tailEnd type="oval" w="med" len="med"/></a:ln>
  </p:spPr>
</p:cxnSp>`,
  ];

  // Row 3: Bent and curved connectors
  const geometryConnectors = [
    `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="7" name="Bent3"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="4000000" y="300000"/><a:ext cx="2500000" cy="1500000"/></a:xfrm>
    <a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst></a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="5B9BD5"/></a:solidFill></a:ln>
  </p:spPr>
</p:cxnSp>`,
    `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="8" name="Curved3"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="4000000" y="2200000"/><a:ext cx="2500000" cy="1500000"/></a:xfrm>
    <a:prstGeom prst="curvedConnector3"><a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst></a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="FFC000"/></a:solidFill></a:ln>
  </p:spPr>
</p:cxnSp>`,
    `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="9" name="Bent3Arrow"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="7000000" y="300000"/><a:ext cx="1500000" cy="1500000"/></a:xfrm>
    <a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst></a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="FF6384"/></a:solidFill><a:tailEnd type="triangle" w="med" len="med"/></a:ln>
  </p:spPr>
</p:cxnSp>`,
    `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="10" name="Curved3Arrow"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="7000000" y="2200000"/><a:ext cx="1500000" cy="1500000"/></a:xfrm>
    <a:prstGeom prst="curvedConnector3"><a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst></a:prstGeom>
    <a:ln w="25400"><a:solidFill><a:srgbClr val="9966FF"/></a:solidFill><a:headEnd type="oval" w="sm" len="sm"/><a:tailEnd type="triangle" w="lg" len="lg"/></a:ln>
  </p:spPr>
</p:cxnSp>`,
  ];

  const allXml = [...basicConnectors, ...arrowConnectors, ...geometryConnectors].join("\n");
  const slide = wrapSlideXml(allXml);
  const rels = slideRelsXml();

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "connectors.pptx");
}

async function createCustomGeometryFixture(): Promise<void> {
  // A custom star-like path and a custom curve
  const customShape1 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="CustomStar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3500000" cy="3500000"/></a:xfrm>
    <a:custGeom>
      <a:avLst/>
      <a:gdLst/>
      <a:ahLst/>
      <a:cxnLst/>
      <a:rect l="0" t="0" r="0" b="0"/>
      <a:pathLst>
        <a:path w="1000" h="1000">
          <a:moveTo><a:pt x="500" y="0"/></a:moveTo>
          <a:lnTo><a:pt x="650" y="350"/></a:lnTo>
          <a:lnTo><a:pt x="1000" y="400"/></a:lnTo>
          <a:lnTo><a:pt x="750" y="650"/></a:lnTo>
          <a:lnTo><a:pt x="800" y="1000"/></a:lnTo>
          <a:lnTo><a:pt x="500" y="850"/></a:lnTo>
          <a:lnTo><a:pt x="200" y="1000"/></a:lnTo>
          <a:lnTo><a:pt x="250" y="650"/></a:lnTo>
          <a:lnTo><a:pt x="0" y="400"/></a:lnTo>
          <a:lnTo><a:pt x="350" y="350"/></a:lnTo>
          <a:close/>
        </a:path>
      </a:pathLst>
    </a:custGeom>
    <a:solidFill><a:srgbClr val="FFC000"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
  </p:spPr>
</p:sp>`;

  const customShape2 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="CustomCurve"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="5000000" y="500000"/><a:ext cx="3500000" cy="3500000"/></a:xfrm>
    <a:custGeom>
      <a:avLst/>
      <a:gdLst/>
      <a:ahLst/>
      <a:cxnLst/>
      <a:rect l="0" t="0" r="0" b="0"/>
      <a:pathLst>
        <a:path w="1000" h="1000">
          <a:moveTo><a:pt x="0" y="500"/></a:moveTo>
          <a:cubicBezTo>
            <a:pt x="250" y="0"/>
            <a:pt x="750" y="1000"/>
            <a:pt x="1000" y="500"/>
          </a:cubicBezTo>
          <a:close/>
        </a:path>
      </a:pathLst>
    </a:custGeom>
    <a:solidFill><a:srgbClr val="5B9BD5"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
  </p:spPr>
</p:sp>`;

  // quadBezTo: quadratic Bezier curve
  const customShape3 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="4" name="QuadBez"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="3200000"/><a:ext cx="2500000" cy="1500000"/></a:xfrm>
    <a:custGeom>
      <a:avLst/>
      <a:gdLst/>
      <a:pathLst>
        <a:path w="1000" h="1000">
          <a:moveTo><a:pt x="0" y="1000"/></a:moveTo>
          <a:quadBezTo>
            <a:pt x="500" y="0"/>
            <a:pt x="1000" y="1000"/>
          </a:quadBezTo>
          <a:close/>
        </a:path>
      </a:pathLst>
    </a:custGeom>
    <a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
  </p:spPr>
</p:sp>`;

  // arcTo: Elliptical arc (semicircle)
  const customShape4 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="5" name="ArcShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="3500000" y="3200000"/><a:ext cx="2500000" cy="1500000"/></a:xfrm>
    <a:custGeom>
      <a:avLst/>
      <a:gdLst/>
      <a:pathLst>
        <a:path w="1000" h="500">
          <a:moveTo><a:pt x="0" y="500"/></a:moveTo>
          <a:arcTo wR="500" hR="500" stAng="10800000" swAng="10800000"/>
          <a:close/>
        </a:path>
      </a:pathLst>
    </a:custGeom>
    <a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
  </p:spPr>
</p:sp>`;

  // adjustValues: dynamic coordinates with guide values
  const customShape5 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="6" name="AdjustShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="6500000" y="3200000"/><a:ext cx="2500000" cy="1500000"/></a:xfrm>
    <a:custGeom>
      <a:avLst>
        <a:gd name="adj" fmla="val 50000"/>
      </a:avLst>
      <a:gdLst>
        <a:gd name="midX" fmla="*/ w adj 100000"/>
        <a:gd name="midY" fmla="*/ h adj 100000"/>
      </a:gdLst>
      <a:pathLst>
        <a:path w="1000" h="1000">
          <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
          <a:lnTo><a:pt x="w" y="0"/></a:lnTo>
          <a:lnTo><a:pt x="w" y="midY"/></a:lnTo>
          <a:lnTo><a:pt x="midX" y="midY"/></a:lnTo>
          <a:lnTo><a:pt x="midX" y="h"/></a:lnTo>
          <a:lnTo><a:pt x="0" y="h"/></a:lnTo>
          <a:close/>
        </a:path>
      </a:pathLst>
    </a:custGeom>
    <a:solidFill><a:srgbClr val="9B59B6"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
  </p:spPr>
</p:sp>`;

  const slide = wrapSlideXml(
    [customShape1, customShape2, customShape3, customShape4, customShape5].join("\n"),
  );
  const rels = slideRelsXml();

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "custom-geometry.pptx");
}

async function createFlowchartFixture(): Promise<void> {
  const flowchartPresets1 = [
    "flowChartProcess",
    "flowChartAlternateProcess",
    "flowChartDecision",
    "flowChartInputOutput",
    "flowChartPredefinedProcess",
    "flowChartInternalStorage",
    "flowChartDocument",
    "flowChartMultidocument",
    "flowChartTerminator",
    "flowChartPreparation",
    "flowChartManualInput",
    "flowChartManualOperation",
    "flowChartConnector",
    "flowChartOffpageConnector",
  ];
  const flowchartPresets2 = [
    "flowChartPunchedCard",
    "flowChartPunchedTape",
    "flowChartCollate",
    "flowChartSort",
    "flowChartExtract",
    "flowChartMerge",
    "flowChartOnlineStorage",
    "flowChartDelay",
    "flowChartDisplay",
    "flowChartMagneticTape",
    "flowChartMagneticDisk",
    "flowChartMagneticDrum",
    "flowChartSummingJunction",
    "flowChartOr",
  ];

  const makeSlide = (presets: string[], cols: number, rows: number): string => {
    let id = 2;
    const shapes = presets.map((preset, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const pos = gridPosition(col, row, cols, rows);
      return shapeXml(id++, preset, {
        preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(COLORS[i % COLORS.length]),
        outlineXml: outlineXml(12700, "333333"),
      });
    });
    return wrapSlideXml(shapes.join("\n"));
  };

  const slide1 = makeSlide(flowchartPresets1, 7, 2);
  const slide2 = makeSlide(flowchartPresets2, 7, 2);
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "flowchart.pptx");
}

async function createCalloutsArcsFixture(): Promise<void> {
  const presets = [
    "wedgeRectCallout",
    "wedgeRoundRectCallout",
    "wedgeEllipseCallout",
    "cloudCallout",
    "borderCallout1",
    "borderCallout2",
    "borderCallout3",
    "arc",
    "chord",
    "pie",
    "blockArc",
  ];

  let id = 2;
  const shapes = presets.map((preset, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const pos = gridPosition(col, row, 4, 3);
    return shapeXml(id++, preset, {
      preset,
      x: pos.x,
      y: pos.y,
      cx: pos.w,
      cy: pos.h,
      fillXml: solidFillXml(COLORS[i % COLORS.length]),
      outlineXml: outlineXml(12700, "333333"),
    });
  });

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "callouts-arcs.pptx");
}

async function createArrowsStarsFixture(): Promise<void> {
  const arrowPresets = [
    "leftRightArrow",
    "upDownArrow",
    "notchedRightArrow",
    "stripedRightArrow",
    "chevron",
    "homePlate",
    "bentArrow",
    "bendUpArrow",
    "quadArrow",
    "leftUpArrow",
  ];
  const starPresets = [
    "star6",
    "star8",
    "star10",
    "star12",
    "star16",
    "star24",
    "star32",
    "irregularSeal1",
    "irregularSeal2",
    "sun",
  ];

  const makeSlide = (presets: string[], cols: number, rows: number): string => {
    let id = 2;
    const shapes = presets.map((preset, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const pos = gridPosition(col, row, cols, rows);
      return shapeXml(id++, preset, {
        preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(COLORS[i % COLORS.length]),
        outlineXml: outlineXml(12700, "333333"),
      });
    });
    return wrapSlideXml(shapes.join("\n"));
  };

  const slide1 = makeSlide(arrowPresets, 5, 2);
  const slide2 = makeSlide(starPresets, 5, 2);

  // Slide 3: Wide (banner-shaped) arrows to test adj calculation with w >> h
  const wideArrowPresets = ["chevron", "homePlate", "notchedRightArrow", "stripedRightArrow"];
  const slide3 = (() => {
    let id = 2;
    const margin = 200000;
    const shapeW = SLIDE_W - margin * 2;
    const shapeH = 650000; // ~68px - much shorter than wide
    const shapes = wideArrowPresets.map((preset, i) => {
      const y = margin + i * (shapeH + margin);
      return shapeXml(id++, preset, {
        preset,
        x: margin,
        y,
        cx: shapeW,
        cy: shapeH,
        fillXml: solidFillXml(COLORS[i % COLORS.length]),
        outlineXml: outlineXml(12700, "333333"),
      });
    });
    return wrapSlideXml(shapes.join("\n"));
  })();

  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
      { xml: slide3, rels },
    ],
  });
  savePptx(buffer, "arrows-stars.pptx");
}

async function createMathOtherFixture(): Promise<void> {
  const mathPresets = [
    "mathPlus",
    "mathMinus",
    "mathMultiply",
    "mathDivide",
    "mathEqual",
    "mathNotEqual",
  ];
  const otherPresets1 = [
    "plus",
    "heptagon",
    "octagon",
    "decagon",
    "dodecagon",
    "plaque",
    "can",
    "cube",
    "donut",
    "noSmoking",
    "smileyFace",
    "foldedCorner",
    "frame",
    "bevel",
  ];
  const otherPresets2 = [
    "halfFrame",
    "corner",
    "diagStripe",
    "snip1Rect",
    "snip2SameRect",
    "snip2DiagRect",
    "snipRoundRect",
    "round1Rect",
    "round2SameRect",
    "round2DiagRect",
    "lightningBolt",
    "moon",
    "teardrop",
    "wave",
    "doubleWave",
    "ribbon",
    "ribbon2",
    "bracketPair",
    "bracePair",
    "leftBracket",
  ];

  const makeSlide = (presets: string[], cols: number, rows: number): string => {
    let id = 2;
    const shapes = presets.map((preset, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const pos = gridPosition(col, row, cols, rows);
      return shapeXml(id++, preset, {
        preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml(COLORS[i % COLORS.length]),
        outlineXml: outlineXml(12700, "333333"),
      });
    });
    return wrapSlideXml(shapes.join("\n"));
  };

  const slide1 = makeSlide(mathPresets, 3, 2);
  const slide2 = makeSlide(otherPresets1, 7, 2);
  const slide3 = makeSlide(otherPresets2, 7, 3);
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
      { xml: slide3, rels },
    ],
  });
  savePptx(buffer, "math-other.pptx");
}

function gridPosition43(
  col: number,
  row: number,
  cols: number,
  rows: number,
  margin = 200000,
): GridPos {
  const cellW = (SLIDE_W_4_3 - margin * (cols + 1)) / cols;
  const cellH = (SLIDE_H_4_3 - margin * (rows + 1)) / rows;
  return {
    x: margin + col * (cellW + margin),
    y: margin + row * (cellH + margin),
    w: cellW,
    h: cellH,
  };
}

async function createSlideSize43Fixture(): Promise<void> {
  const slideSize43 = { cx: SLIDE_W_4_3, cy: SLIDE_H_4_3, type: "screen4x3" };

  // Slide 1: Basic shapes on 4:3
  const shapeDefs = [
    { preset: "rect", label: "Rectangle" },
    { preset: "ellipse", label: "Ellipse" },
    { preset: "roundRect", label: "RoundRect" },
    { preset: "diamond", label: "Diamond" },
    { preset: "triangle", label: "Triangle" },
    { preset: "hexagon", label: "Hexagon" },
  ];

  let id = 2;
  const shapes1: string[] = [];
  shapeDefs.forEach((def, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const pos = gridPosition43(col, row, 3, 2);
    shapes1.push(
      shapeXml(id++, def.label, {
        preset: def.preset,
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        textBodyXml: textBodyXmlHelper(def.label, { fontSize: 14, color: "FFFFFF" }),
      }),
    );
  });

  const slide1 = wrapSlideXml(shapes1.join("\n"));

  // Slide 2: Text layout on 4:3
  id = 2;
  const shapes2: string[] = [];

  // Large title text at top
  const titlePos = gridPosition43(0, 0, 1, 3);
  shapes2.push(
    shapeXml(id++, "Title", {
      preset: "rect",
      x: titlePos.x,
      y: titlePos.y,
      cx: titlePos.w,
      cy: titlePos.h,
      fillXml: solidFillXml("1F4E79"),
      textBodyXml: textBodyXmlHelper("4:3 Slide Title", {
        fontSize: 32,
        bold: true,
        color: "FFFFFF",
        align: "ctr",
      }),
    }),
  );

  // Text boxes in 2x1 grid for body area
  const leftPos = gridPosition43(0, 1, 2, 3);
  shapes2.push(
    shapeXml(id++, "LeftText", {
      preset: "roundRect",
      x: leftPos.x,
      y: leftPos.y,
      cx: leftPos.w,
      cy: leftPos.h,
      fillXml: solidFillXml("E8F0FE"),
      outlineXml: outlineXml(12700, "4472C4"),
      textBodyXml: textBodyXmlHelper("Left content area with text that spans multiple lines", {
        fontSize: 14,
        color: "333333",
        anchor: "t",
      }),
    }),
  );

  const rightPos = gridPosition43(1, 1, 2, 3);
  shapes2.push(
    shapeXml(id++, "RightText", {
      preset: "roundRect",
      x: rightPos.x,
      y: rightPos.y,
      cx: rightPos.w,
      cy: rightPos.h,
      fillXml: solidFillXml("FFF2CC"),
      outlineXml: outlineXml(12700, "FFC000"),
      textBodyXml: textBodyXmlHelper("Right content area with different styling", {
        fontSize: 14,
        italic: true,
        color: "333333",
        anchor: "t",
      }),
    }),
  );

  // Bottom bar
  const bottomPos = gridPosition43(0, 2, 1, 3);
  shapes2.push(
    shapeXml(id++, "Footer", {
      preset: "rect",
      x: bottomPos.x,
      y: bottomPos.y,
      cx: bottomPos.w,
      cy: bottomPos.h,
      fillXml: solidFillXml("44546A"),
      textBodyXml: textBodyXmlHelper("Footer on 4:3 slide", {
        fontSize: 10,
        color: "FFFFFF",
        align: "ctr",
      }),
    }),
  );

  const slide2 = wrapSlideXml(shapes2.join("\n"));

  // Slide 3: Background on 4:3
  const bgXml = `<p:bg>
    <p:bgPr>
      <a:gradFill>
        <a:gsLst>
          <a:gs pos="0"><a:srgbClr val="1A237E"/></a:gs>
          <a:gs pos="50000"><a:srgbClr val="4A148C"/></a:gs>
          <a:gs pos="100000"><a:srgbClr val="880E4F"/></a:gs>
        </a:gsLst>
        <a:lin ang="2700000" scaled="1"/>
      </a:gradFill>
      <a:effectLst/>
    </p:bgPr>
  </p:bg>`;

  id = 2;
  const shapes3: string[] = [];
  const centerPos = gridPosition43(0, 0, 1, 1);
  shapes3.push(
    shapeXml(id++, "CenterShape", {
      preset: "ellipse",
      x: centerPos.x + centerPos.w / 4,
      y: centerPos.y + centerPos.h / 4,
      cx: centerPos.w / 2,
      cy: centerPos.h / 2,
      fillXml: `<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="70000"/></a:srgbClr></a:solidFill>`,
      textBodyXml: textBodyXmlHelper("4:3 BG", {
        fontSize: 24,
        bold: true,
        color: "1A237E",
        align: "ctr",
      }),
    }),
  );

  const slide3 = wrapSlideXml(shapes3.join("\n"), bgXml);

  const rels = slideRelsXml();
  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
      { xml: slide3, rels },
    ],
    slideSize: slideSize43,
  });
  savePptx(buffer, "slide-size-4-3.pptx");
}

// --- Main ---
async function createEffectsFixture(): Promise<void> {
  let id = 2;
  const shapes: string[] = [];

  // Outer shadow (downward)
  const pos0 = gridPosition(0, 0, 3, 2);
  shapes.push(
    shapeXml(id++, "OuterShadow", {
      preset: "roundRect",
      x: pos0.x,
      y: pos0.y,
      cx: pos0.w,
      cy: pos0.h,
      fillXml: solidFillXml("4472C4"),
      effectsXml: `<a:outerShdw blurRad="40000" dist="20000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw>`,
    }),
  );

  // Outer shadow (bottom right, larger)
  const pos1 = gridPosition(1, 0, 3, 2);
  shapes.push(
    shapeXml(id++, "OuterShadow-Large", {
      preset: "rect",
      x: pos1.x,
      y: pos1.y,
      cx: pos1.w,
      cy: pos1.h,
      fillXml: solidFillXml("ED7D31"),
      effectsXml: `<a:outerShdw blurRad="76200" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw>`,
    }),
  );

  // Glow
  const pos2 = gridPosition(2, 0, 3, 2);
  shapes.push(
    shapeXml(id++, "Glow", {
      preset: "ellipse",
      x: pos2.x,
      y: pos2.y,
      cx: pos2.w,
      cy: pos2.h,
      fillXml: solidFillXml("5B9BD5"),
      effectsXml: `<a:glow rad="63500"><a:srgbClr val="4472C4"><a:alpha val="40000"/></a:srgbClr></a:glow>`,
    }),
  );

  // Inner shadow
  const pos3 = gridPosition(0, 1, 3, 2);
  shapes.push(
    shapeXml(id++, "InnerShadow", {
      preset: "roundRect",
      x: pos3.x,
      y: pos3.y,
      cx: pos3.w,
      cy: pos3.h,
      fillXml: solidFillXml("A5A5A5"),
      effectsXml: `<a:innerShdw blurRad="63500" dist="50800" dir="2700000"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:innerShdw>`,
    }),
  );

  // Soft edge
  const pos4 = gridPosition(1, 1, 3, 2);
  shapes.push(
    shapeXml(id++, "SoftEdge", {
      preset: "rect",
      x: pos4.x,
      y: pos4.y,
      cx: pos4.w,
      cy: pos4.h,
      fillXml: solidFillXml("FFC000"),
      effectsXml: `<a:softEdge rad="31750"/>`,
    }),
  );

  // Combined (shadow + glow)
  const pos5 = gridPosition(2, 1, 3, 2);
  shapes.push(
    shapeXml(id++, "Combined", {
      preset: "diamond",
      x: pos5.x,
      y: pos5.y,
      cx: pos5.w,
      cy: pos5.h,
      fillXml: solidFillXml("70AD47"),
      effectsXml: `<a:outerShdw blurRad="40000" dist="20000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw><a:glow rad="63500"><a:srgbClr val="70AD47"><a:alpha val="30000"/></a:srgbClr></a:glow>`,
    }),
  );

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();
  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "effects.pptx");
}

async function createColorTransformsFixture(): Promise<void> {
  const baseColor = "4472C4";
  const shapes: string[] = [];

  const cases: { label: string; fillXml: string }[] = [
    // Row 0
    {
      label: "Base",
      fillXml: `<a:solidFill><a:srgbClr val="${baseColor}"/></a:solidFill>`,
    },
    {
      label: "tint 50%",
      fillXml: `<a:solidFill><a:srgbClr val="${baseColor}"><a:tint val="50000"/></a:srgbClr></a:solidFill>`,
    },
    {
      label: "tint 80%",
      fillXml: `<a:solidFill><a:srgbClr val="${baseColor}"><a:tint val="80000"/></a:srgbClr></a:solidFill>`,
    },
    // Row 1
    {
      label: "shade 50%",
      fillXml: `<a:solidFill><a:srgbClr val="${baseColor}"><a:shade val="50000"/></a:srgbClr></a:solidFill>`,
    },
    {
      label: "shade 80%",
      fillXml: `<a:solidFill><a:srgbClr val="${baseColor}"><a:shade val="80000"/></a:srgbClr></a:solidFill>`,
    },
    {
      label: "lumMod 75%\n(schemeClr)",
      fillXml: `<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>`,
    },
    // Row 2
    {
      label: "lumMod 50%\nlumOff 50%",
      fillXml: `<a:solidFill><a:srgbClr val="${baseColor}"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:srgbClr></a:solidFill>`,
    },
    {
      label: "alpha 50%",
      fillXml: `<a:solidFill><a:srgbClr val="${baseColor}"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`,
    },
    {
      label: "tint 40%\nshade 80%",
      fillXml: `<a:solidFill><a:srgbClr val="${baseColor}"><a:tint val="40000"/><a:shade val="80000"/></a:srgbClr></a:solidFill>`,
    },
  ];

  cases.forEach((c, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const pos = gridPosition(col, row, 3, 3);
    shapes.push(
      shapeXml(i + 2, c.label, {
        preset: "roundRect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: c.fillXml,
        textBodyXml: textBodyXmlHelper(c.label, { fontSize: 10, color: "FFFFFF" }),
      }),
    );
  });

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();
  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "color-transforms.pptx");
}

export const shapeFixtureCreators: FixtureCreatorMap = {
  "shapes.pptx": createShapesFixture,
  "fill-and-lines.pptx": createFillAndLinesFixture,
  "transform.pptx": createTransformFixture,
  "groups.pptx": createGroupsFixture,
  "connectors.pptx": createConnectorsFixture,
  "custom-geometry.pptx": createCustomGeometryFixture,
  "flowchart.pptx": createFlowchartFixture,
  "callouts-arcs.pptx": createCalloutsArcsFixture,
  "arrows-stars.pptx": createArrowsStarsFixture,
  "math-other.pptx": createMathOtherFixture,
  "slide-size-4-3.pptx": createSlideSize43Fixture,
  "effects.pptx": createEffectsFixture,
  "color-transforms.pptx": createColorTransformsFixture,
};
