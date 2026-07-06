import sharp from "sharp";

import type { FixtureCreatorMap } from "../fixture-builder.js";
import {
  buildPptx,
  outlineXml,
  REL_TYPES,
  savePptx,
  shapeXml,
  slideRelsXml,
  solidFillXml,
  textBodyXmlHelper,
  wrapSlideXml,
} from "../fixture-builder.js";

function tableGraphicFrameXml(
  id: number,
  name: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  tblXml: string,
): string {
  return `<p:graphicFrame>
  <p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>
  <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      ${tblXml}
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`;
}

function tableCellXml(
  text: string,
  opts?: {
    fillColor?: string;
    bold?: boolean;
    fontSize?: number;
    fontColor?: string;
    borderColor?: string;
    borderWidth?: number;
    gridSpan?: number;
    rowSpan?: number;
    hMerge?: boolean;
    vMerge?: boolean;
  },
): string {
  const spanAttrs = [
    opts?.gridSpan && opts.gridSpan > 1 ? ` gridSpan="${opts.gridSpan}"` : "",
    opts?.rowSpan && opts.rowSpan > 1 ? ` rowSpan="${opts.rowSpan}"` : "",
  ].join("");

  const sz = opts?.fontSize ? ` sz="${opts.fontSize * 100}"` : ` sz="1200"`;
  const b = opts?.bold ? ` b="1"` : "";
  const fontColor = opts?.fontColor ?? "000000";

  const fillXml = opts?.fillColor
    ? `<a:solidFill><a:srgbClr val="${opts.fillColor}"/></a:solidFill>`
    : "";

  const bw = opts?.borderWidth ?? 12700;
  const bc = opts?.borderColor ?? "000000";
  const borderXml = `<a:lnL w="${bw}"><a:solidFill><a:srgbClr val="${bc}"/></a:solidFill></a:lnL>
      <a:lnR w="${bw}"><a:solidFill><a:srgbClr val="${bc}"/></a:solidFill></a:lnR>
      <a:lnT w="${bw}"><a:solidFill><a:srgbClr val="${bc}"/></a:solidFill></a:lnT>
      <a:lnB w="${bw}"><a:solidFill><a:srgbClr val="${bc}"/></a:solidFill></a:lnB>`;

  const hMergeAttr = opts?.hMerge ? ` hMerge="1"` : "";
  const vMergeAttr = opts?.vMerge ? ` vMerge="1"` : "";

  return `<a:tc${spanAttrs}>
    <a:txBody>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p>
        <a:r>
          <a:rPr lang="en-US"${sz}${b}>
            <a:solidFill><a:srgbClr val="${fontColor}"/></a:solidFill>
          </a:rPr>
          <a:t>${text}</a:t>
        </a:r>
      </a:p>
    </a:txBody>
    <a:tcPr${hMergeAttr}${vMergeAttr}>
      ${borderXml}
      ${fillXml}
    </a:tcPr>
  </a:tc>`;
}

async function createTablesFixture(): Promise<void> {
  const margin = 300000;

  // Slide 1: Basic table with header row
  const colW = 2286000; // 3 columns
  const rowH = 457200;
  const tblW = colW * 3;
  const tblH = rowH * 4;

  const headerRow = `<a:tr h="${rowH}">
    ${tableCellXml("Name", { fillColor: "4472C4", fontColor: "FFFFFF", bold: true })}
    ${tableCellXml("Value", { fillColor: "4472C4", fontColor: "FFFFFF", bold: true })}
    ${tableCellXml("Status", { fillColor: "4472C4", fontColor: "FFFFFF", bold: true })}
  </a:tr>`;

  const dataRows = [
    ["Alpha", "100", "Active"],
    ["Beta", "250", "Pending"],
    ["Gamma", "75", "Done"],
  ]
    .map(
      (row, i) =>
        `<a:tr h="${rowH}">
    ${row.map((cell) => tableCellXml(cell, { fillColor: i % 2 === 0 ? "D6E4F0" : "FFFFFF" })).join("\n    ")}
  </a:tr>`,
    )
    .join("\n  ");

  const tbl1 = `<a:tbl>
    <a:tblPr firstRow="1"/>
    <a:tblGrid>
      <a:gridCol w="${colW}"/>
      <a:gridCol w="${colW}"/>
      <a:gridCol w="${colW}"/>
    </a:tblGrid>
    ${headerRow}
    ${dataRows}
  </a:tbl>`;

  const gf1 = tableGraphicFrameXml(2, "Basic Table", margin, margin, tblW, tblH, tbl1);
  const slide1 = wrapSlideXml(gf1);

  // Slide 2: Table with cell merging
  const col2W = 2286000;
  const row2H = 600000;
  const tbl2W = col2W * 3;
  const tbl2H = row2H * 3;

  const tbl2 = `<a:tbl>
    <a:tblPr/>
    <a:tblGrid>
      <a:gridCol w="${col2W}"/>
      <a:gridCol w="${col2W}"/>
      <a:gridCol w="${col2W}"/>
    </a:tblGrid>
    <a:tr h="${row2H}">
      ${tableCellXml("Merged Header (3 cols)", { fillColor: "ED7D31", fontColor: "FFFFFF", bold: true, gridSpan: 3 })}
      ${tableCellXml("", { hMerge: true })}
      ${tableCellXml("", { hMerge: true })}
    </a:tr>
    <a:tr h="${row2H}">
      ${tableCellXml("Merged\nRows", { fillColor: "A5A5A5", fontColor: "FFFFFF", rowSpan: 2 })}
      ${tableCellXml("Cell B2", { fillColor: "FFFFFF" })}
      ${tableCellXml("Cell C2", { fillColor: "FFFFFF" })}
    </a:tr>
    <a:tr h="${row2H}">
      ${tableCellXml("", { vMerge: true })}
      ${tableCellXml("Cell B3", { fillColor: "D6E4F0" })}
      ${tableCellXml("Cell C3", { fillColor: "D6E4F0" })}
    </a:tr>
  </a:tbl>`;

  const gf2 = tableGraphicFrameXml(2, "Merge Table", margin, margin, tbl2W, tbl2H, tbl2);
  const slide2 = wrapSlideXml(gf2);

  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "tables.pptx");
}

async function createTableComplexMergeFixture(): Promise<void> {
  const margin = 300000;

  // Slide 1: Complex merge (4x4 table)
  // Layout:
  // +----------+----+----+
  // | A (2x2)  | C  | D  |
  // |          |    |    |
  // +----+-----+----+----+
  // | E  | F   | G (2x1) |
  // +----+-----+---------+
  // | H (1x2)  | I  | J  |
  // |          |    |    |
  // +----+-----+----+----+
  const colW = 1714500; // ~4 columns
  const rowH = 600000;

  const tbl1 = `<a:tbl>
    <a:tblPr/>
    <a:tblGrid>
      <a:gridCol w="${colW}"/><a:gridCol w="${colW}"/><a:gridCol w="${colW}"/><a:gridCol w="${colW}"/>
    </a:tblGrid>
    <a:tr h="${rowH}">
      ${tableCellXml("A (2x2)", { fillColor: "4472C4", fontColor: "FFFFFF", bold: true, gridSpan: 2, rowSpan: 2 })}
      ${tableCellXml("", { hMerge: true })}
      ${tableCellXml("C", { fillColor: "D6E4F0" })}
      ${tableCellXml("D", { fillColor: "FFFFFF" })}
    </a:tr>
    <a:tr h="${rowH}">
      ${tableCellXml("", { vMerge: true })}
      ${tableCellXml("", { hMerge: true, vMerge: true })}
      ${tableCellXml("C2", { fillColor: "FFFFFF" })}
      ${tableCellXml("D2", { fillColor: "D6E4F0" })}
    </a:tr>
    <a:tr h="${rowH}">
      ${tableCellXml("E", { fillColor: "D6E4F0" })}
      ${tableCellXml("F", { fillColor: "FFFFFF" })}
      ${tableCellXml("G (2x1)", { fillColor: "ED7D31", fontColor: "FFFFFF", bold: true, gridSpan: 2 })}
      ${tableCellXml("", { hMerge: true })}
    </a:tr>
    <a:tr h="${rowH}">
      ${tableCellXml("H (1x2)", { fillColor: "A5A5A5", fontColor: "FFFFFF", bold: true, rowSpan: 2 })}
      ${tableCellXml("I1", { fillColor: "FFFFFF" })}
      ${tableCellXml("J1", { fillColor: "D6E4F0" })}
      ${tableCellXml("K1", { fillColor: "FFFFFF" })}
    </a:tr>
    <a:tr h="${rowH}">
      ${tableCellXml("", { vMerge: true })}
      ${tableCellXml("I2", { fillColor: "D6E4F0" })}
      ${tableCellXml("J2", { fillColor: "FFFFFF" })}
      ${tableCellXml("K2", { fillColor: "D6E4F0" })}
    </a:tr>
  </a:tbl>`;

  const gf1 = tableGraphicFrameXml(2, "Complex Merge", margin, margin, colW * 4, rowH * 5, tbl1);
  const slide1 = wrapSlideXml(gf1);

  // Slide 2: Uneven column widths and row heights
  const colWidths = [1000000, 3000000, 2858000];
  const rowHeights = [300000, 800000, 500000];
  const cellColors = ["D6E4F0", "FFFFFF"];

  const rows2 = rowHeights
    .map(
      (rh, ri) =>
        `<a:tr h="${rh}">
      ${colWidths.map((_, ci) => tableCellXml(`R${ri + 1}C${ci + 1}`, { fillColor: cellColors[(ri + ci) % 2], fontSize: 10 })).join("\n      ")}
    </a:tr>`,
    )
    .join("\n    ");

  const tbl2 = `<a:tbl>
    <a:tblPr/>
    <a:tblGrid>
      ${colWidths.map((w) => `<a:gridCol w="${w}"/>`).join("")}
    </a:tblGrid>
    ${rows2}
  </a:tbl>`;

  const tbl2W = colWidths.reduce((a, b) => a + b, 0);
  const tbl2H = rowHeights.reduce((a, b) => a + b, 0);
  const gf2 = tableGraphicFrameXml(2, "Uneven Table", margin, margin, tbl2W, tbl2H, tbl2);
  const slide2 = wrapSlideXml(gf2);

  const rels = slideRelsXml();
  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "table-complex-merge.pptx");
}

function tableCellXmlNoBorder(
  text: string,
  opts?: {
    fillColor?: string;
    bold?: boolean;
    fontSize?: number;
    fontColor?: string;
  },
): string {
  const sz = opts?.fontSize ? ` sz="${opts.fontSize * 100}"` : ` sz="1200"`;
  const b = opts?.bold ? ` b="1"` : "";
  const fontColor = opts?.fontColor ?? "000000";

  const fillXml = opts?.fillColor
    ? `<a:solidFill><a:srgbClr val="${opts.fillColor}"/></a:solidFill>`
    : "";

  return `<a:tc>
    <a:txBody>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p>
        <a:r>
          <a:rPr lang="en-US"${sz}${b}>
            <a:solidFill><a:srgbClr val="${fontColor}"/></a:solidFill>
          </a:rPr>
          <a:t>${text}</a:t>
        </a:r>
      </a:p>
    </a:txBody>
    <a:tcPr>
      ${fillXml}
    </a:tcPr>
  </a:tc>`;
}

async function createTableStyleBorderFixture(): Promise<void> {
  const margin = 300000;
  const colW = 2286000;
  const rowH = 457200;
  const tblW = colW * 3;
  const tblH = rowH * 3;

  const tbl1 = `<a:tbl>
    <a:tblPr firstRow="1">
      <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
    </a:tblPr>
    <a:tblGrid>
      <a:gridCol w="${colW}"/>
      <a:gridCol w="${colW}"/>
      <a:gridCol w="${colW}"/>
    </a:tblGrid>
    <a:tr h="${rowH}">
      ${tableCellXmlNoBorder("Header A", { fillColor: "4472C4", fontColor: "FFFFFF", bold: true })}
      ${tableCellXmlNoBorder("Header B", { fillColor: "4472C4", fontColor: "FFFFFF", bold: true })}
      ${tableCellXmlNoBorder("Header C", { fillColor: "4472C4", fontColor: "FFFFFF", bold: true })}
    </a:tr>
    <a:tr h="${rowH}">
      ${tableCellXmlNoBorder("Cell 1", { fillColor: "D6E4F0" })}
      ${tableCellXmlNoBorder("Cell 2", { fillColor: "D6E4F0" })}
      ${tableCellXmlNoBorder("Cell 3", { fillColor: "D6E4F0" })}
    </a:tr>
    <a:tr h="${rowH}">
      ${tableCellXmlNoBorder("Cell 4", { fillColor: "FFFFFF" })}
      ${tableCellXmlNoBorder("Cell 5", { fillColor: "FFFFFF" })}
      ${tableCellXmlNoBorder("Cell 6", { fillColor: "FFFFFF" })}
    </a:tr>
  </a:tbl>`;

  const gf1 = tableGraphicFrameXml(2, "Style Border Table", margin, margin, tblW, tblH, tbl1);
  const slide1 = wrapSlideXml(gf1);
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [{ xml: slide1, rels }],
  });
  savePptx(buffer, "table-style-border.pptx");
}

async function createZOrderMixedFixture(): Promise<void> {
  // Generate image (blue gradient)
  const imgSize = 100;
  const pixels = Buffer.alloc(imgSize * imgSize * 4);
  for (let y = 0; y < imgSize; y++) {
    for (let x = 0; x < imgSize; x++) {
      const idx = (y * imgSize + x) * 4;
      pixels[idx] = 30; // R
      pixels[idx + 1] = 100; // G
      pixels[idx + 2] = Math.floor(150 + (x / imgSize) * 105); // B gradient
      pixels[idx + 3] = 255; // A
    }
  }
  const testImage = await sharp(pixels, {
    raw: { width: imgSize, height: imgSize, channels: 4 },
  })
    .png()
    .toBuffer();

  // Slide 1: sp -> pic -> sp (image is sandwiched between two shapes)
  // Correct Z-order: Red rectangle (backmost) -> Image (middle) -> Green rectangle (frontmost)
  const spTreeContent1 = [
    // 1. Red rectangle (backmost, Z=1)
    shapeXml(2, "back-rect", {
      preset: "rect",
      x: 500000,
      y: 500000,
      cx: 5000000,
      cy: 3500000,
      fillXml: solidFillXml("CC3333"),
      outlineXml: outlineXml(25400, "990000"),
      textBodyXml: textBodyXmlHelper("Back (Z=1)", {
        fontSize: 18,
        bold: true,
        color: "FFFFFF",
        anchor: "t",
        align: "l",
      }),
    }),
    // 2. Image (middle, Z=2)
    `<p:pic>
  <p:nvPicPr><p:cNvPr id="3" name="Image 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId2"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="1500000" y="1000000"/><a:ext cx="4000000" cy="2500000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`,
    // 3. Green rectangle (frontmost, Z=3)
    shapeXml(4, "front-rect", {
      preset: "roundRect",
      x: 3000000,
      y: 1500000,
      cx: 4000000,
      cy: 2500000,
      fillXml: `<a:solidFill><a:srgbClr val="33AA33"><a:alpha val="80000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(25400, "006600"),
      textBodyXml: textBodyXmlHelper("Front (Z=3)", {
        fontSize: 18,
        bold: true,
        color: "FFFFFF",
      }),
    }),
  ].join("\n");

  const slide1 = wrapSlideXml(spTreeContent1);
  const rels1 = slideRelsXml([
    { id: "rId2", type: REL_TYPES.image, target: "../media/image1.png" },
  ]);

  // Slide 2: cxnSp -> sp -> pic -> sp (also includes connectors)
  const spTreeContent2 = [
    // 1. Connector (backmost)
    `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="2" name="Connector 1"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="500000" y="2500000"/><a:ext cx="8000000" cy="0"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    <a:ln w="50800"><a:solidFill><a:srgbClr val="FF6600"/></a:solidFill></a:ln>
  </p:spPr>
</p:cxnSp>`,
    // 2. Yellow rectangle
    shapeXml(3, "yellow-rect", {
      preset: "rect",
      x: 1000000,
      y: 800000,
      cx: 3500000,
      cy: 3000000,
      fillXml: solidFillXml("FFCC00"),
      outlineXml: outlineXml(19050, "CC9900"),
      textBodyXml: textBodyXmlHelper("Shape (Z=2)", {
        fontSize: 14,
        bold: true,
        color: "333333",
        anchor: "t",
      }),
    }),
    // 3. Image
    `<p:pic>
  <p:nvPicPr><p:cNvPr id="4" name="Image 2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId2"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="3000000" y="1500000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`,
    // 4. Purple rectangle (frontmost)
    shapeXml(5, "purple-rect", {
      preset: "ellipse",
      x: 5000000,
      y: 500000,
      cx: 3500000,
      cy: 3500000,
      fillXml: `<a:solidFill><a:srgbClr val="6633CC"><a:alpha val="70000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(19050, "330066"),
      textBodyXml: textBodyXmlHelper("Shape (Z=4)", {
        fontSize: 14,
        bold: true,
        color: "FFFFFF",
      }),
    }),
  ].join("\n");

  const slide2 = wrapSlideXml(spTreeContent2);
  const rels2 = slideRelsXml([
    { id: "rId2", type: REL_TYPES.image, target: "../media/image1.png" },
  ]);

  // Slide 3: All 5 element types mixed (cxnSp -> grpSp -> pic -> graphicFrame(table) -> sp)
  const spTreeContent3 = [
    // 1. cxnSp (backmost, Z=1)
    `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="10" name="Connector BG"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="300000" y="1000000"/><a:ext cx="8500000" cy="3000000"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    <a:ln w="76200"><a:solidFill><a:srgbClr val="FF6600"/></a:solidFill></a:ln>
  </p:spPr>
</p:cxnSp>`,
    // 2. grpSp (Z=2)
    `<p:grpSp>
  <p:nvGrpSpPr><p:cNvPr id="11" name="Group Z2"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm>
      <a:off x="500000" y="500000"/>
      <a:ext cx="3500000" cy="2500000"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="3500000" cy="2500000"/>
    </a:xfrm>
  </p:grpSpPr>
  ${shapeXml(12, "GrpRect", {
    preset: "rect",
    x: 0,
    y: 0,
    cx: 1600000,
    cy: 2500000,
    fillXml: solidFillXml("4472C4"),
    textBodyXml: textBodyXmlHelper("Grp (Z=2)", { fontSize: 12, color: "FFFFFF" }),
  })}
  ${shapeXml(13, "GrpEllipse", {
    preset: "ellipse",
    x: 1800000,
    y: 0,
    cx: 1600000,
    cy: 2500000,
    fillXml: solidFillXml("5B9BD5"),
  })}
</p:grpSp>`,
    // 3. pic (Z=3)
    `<p:pic>
  <p:nvPicPr><p:cNvPr id="14" name="Image Z3"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId2"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="2000000" y="1000000"/><a:ext cx="3000000" cy="2500000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`,
    // 4. graphicFrame - table (Z=4)
    tableGraphicFrameXml(
      15,
      "Table Z4",
      3500000,
      1500000,
      3500000,
      1500000,
      `<a:tbl>
    <a:tblPr firstRow="1"/>
    <a:tblGrid>
      <a:gridCol w="1750000"/>
      <a:gridCol w="1750000"/>
    </a:tblGrid>
    <a:tr h="750000">
      ${tableCellXml("Tbl", { fillColor: "ED7D31", fontColor: "FFFFFF", bold: true })}
      ${tableCellXml("Z=4", { fillColor: "ED7D31", fontColor: "FFFFFF", bold: true })}
    </a:tr>
    <a:tr h="750000">
      ${tableCellXml("A", { fillColor: "FFF2CC" })}
      ${tableCellXml("B", { fillColor: "FFF2CC" })}
    </a:tr>
  </a:tbl>`,
    ),
    // 5. sp (frontmost, Z=5)
    shapeXml(16, "front-shape", {
      preset: "ellipse",
      x: 5000000,
      y: 500000,
      cx: 3500000,
      cy: 3500000,
      fillXml: `<a:solidFill><a:srgbClr val="70AD47"><a:alpha val="75000"/></a:srgbClr></a:solidFill>`,
      outlineXml: outlineXml(25400, "2E7D32"),
      textBodyXml: textBodyXmlHelper("Shape (Z=5)", {
        fontSize: 16,
        bold: true,
        color: "FFFFFF",
      }),
    }),
  ].join("\n");

  const slide3 = wrapSlideXml(spTreeContent3);
  const rels3 = slideRelsXml([
    { id: "rId2", type: REL_TYPES.image, target: "../media/image1.png" },
  ]);

  const media = new Map<string, Buffer>();
  media.set("ppt/media/image1.png", testImage);

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels: rels1 },
      { xml: slide2, rels: rels2 },
      { xml: slide3, rels: rels3 },
    ],
    media,
  });
  savePptx(buffer, "z-order-mixed.pptx");
}

export const tableFixtureCreators: FixtureCreatorMap = {
  "tables.pptx": createTablesFixture,
  "table-complex-merge.pptx": createTableComplexMergeFixture,
  "table-style-border.pptx": createTableStyleBorderFixture,
  "z-order-mixed.pptx": createZOrderMixedFixture,
};
