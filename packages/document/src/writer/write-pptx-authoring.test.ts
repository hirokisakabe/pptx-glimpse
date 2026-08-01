import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

// Import via the actual public surface (`@pptx-glimpse/document`).
import {
  addChart,
  addConnector,
  addEmptySlideFromLayout,
  addPicture,
  addShape,
  addSlideNumber,
  addTable,
  addTextBox,
  asEmu,
  asHundredthPt,
  asOoxmlAngle,
  asOoxmlPercent,
  asPartPath,
  asPt,
  clearBackground,
  createComputedView,
  createPptx,
  deleteShape,
  readPptx,
  setBackground,
  setSlideBackground,
  writePptx,
} from "../index.js";
import {
  BLUE_PNG,
  buildShapeDeleteFixture,
  decoder,
  encoder,
  findImageByName,
  findShapeByName,
  findTextRun,
  getEntry,
  RED_PNG,
  requireHandle,
  requireShape,
} from "./write-pptx.test-helpers.js";

describe("writePptx - from-scratch builder", () => {
  it("writes, reads, and computes a customized theme color and script font scheme", () => {
    let source = createPptx({
      theme: {
        name: "Product & Theme",
        colorScheme: {
          name: "Product Colors",
          dk1: "102030",
          lt1: "F0F1F2",
          dk2: "203040",
          lt2: "E0E1E2",
          accent1: "12ab34",
          accent2: "223344",
          accent3: "334455",
          accent4: "445566",
          accent5: "556677",
          accent6: "667788",
          hlink: "0055cc",
          folHlink: "775599",
        },
        fontScheme: {
          name: "Product Fonts",
          major: {
            latin: "Brand Display",
            eastAsian: "Noto Sans CJK JP",
            complexScript: "Noto Sans Arabic",
          },
          minor: {
            latin: "Brand Text",
            eastAsian: "Yu Gothic",
            complexScript: "Arial",
          },
        },
      },
    });
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(100000),
      offsetY: asEmu(100000),
      width: asEmu(1000000),
      height: asEmu(500000),
      fill: { kind: "solid", color: { kind: "srgb", hex: "FFFFFF" } },
      paragraphs: [{ runs: [{ text: "Script fonts", properties: { fontFace: "+mj-lt" } }] }],
    });

    const output = writePptx(source);
    const themeXml = decoder.decode(getEntry(output, "ppt/theme/theme1.xml"));
    expect(themeXml).toContain(`name="Product &amp; Theme"`);
    expect(themeXml).toContain(`<a:clrScheme name="Product Colors">`);
    expect(themeXml).toContain(`<a:dk1><a:srgbClr val="102030"/></a:dk1>`);
    expect(themeXml).toContain(`<a:accent1><a:srgbClr val="12AB34"/></a:accent1>`);
    const expectedColors = {
      dk1: "102030",
      lt1: "F0F1F2",
      dk2: "203040",
      lt2: "E0E1E2",
      accent1: "12AB34",
      accent2: "223344",
      accent3: "334455",
      accent4: "445566",
      accent5: "556677",
      accent6: "667788",
      hlink: "0055CC",
      folHlink: "775599",
    };
    for (const [slot, hex] of Object.entries(expectedColors)) {
      expect(themeXml).toContain(`<a:${slot}><a:srgbClr val="${hex}"/></a:${slot}>`);
    }
    expect(themeXml).toContain(`<a:fontScheme name="Product Fonts">`);
    expect(themeXml).toContain(
      `<a:majorFont><a:latin typeface="Brand Display"/><a:ea typeface="Noto Sans CJK JP"/>` +
        `<a:cs typeface="Noto Sans Arabic"/></a:majorFont>`,
    );

    const archive = unzipSync(output);
    const slideXml = decoder
      .decode(getEntry(output, "ppt/slides/slide1.xml"))
      .replace(`<a:srgbClr val="FFFFFF"/>`, `<a:schemeClr val="accent1"/>`)
      .replace(
        `<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-lt"/><a:cs typeface="+mj-lt"/>`,
        `<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/>`,
      );
    const reread = readPptx(
      zipSync({ ...archive, "ppt/slides/slide1.xml": encoder.encode(slideXml) }),
    );
    expect(reread.diagnostics).toEqual([]);
    expect(reread.themes[0]).toMatchObject({
      name: "Product & Theme",
      colorScheme: {
        colors: {
          dk1: { kind: "srgb", hex: "102030" },
          accent1: { kind: "srgb", hex: "12AB34" },
          accent2: { kind: "srgb", hex: "223344" },
          hlink: { kind: "srgb", hex: "0055CC" },
        },
      },
      fontScheme: {
        majorLatin: "Brand Display",
        majorEastAsian: "Noto Sans CJK JP",
        majorComplexScript: "Noto Sans Arabic",
        minorLatin: "Brand Text",
        minorEastAsian: "Yu Gothic",
        minorComplexScript: "Arial",
      },
    });
    const computedShape = createComputedView(reread).slides[0]?.elements[0];
    expect(computedShape).toMatchObject({
      kind: "shape",
      fill: { kind: "solid", color: { hex: "#12ab34" } },
      textBody: {
        paragraphs: [
          {
            runs: [
              {
                properties: {
                  typeface: "Brand Display",
                  typefaceEa: "Noto Sans CJK JP",
                  typefaceCs: "Noto Sans Arabic",
                },
              },
            ],
          },
        ],
      },
    });
  });

  it("applies theme defaults per field and rejects invalid theme inputs", () => {
    const source = createPptx({
      theme: {
        colorScheme: { accent1: "abcdef" },
        fontScheme: { major: { eastAsian: "", complexScript: "" } },
      },
    });
    expect(source.themes[0]).toMatchObject({
      name: "Office Theme",
      colorScheme: {
        colors: {
          dk1: { kind: "system", value: "windowText", lastColor: "000000" },
          accent1: { kind: "srgb", hex: "ABCDEF" },
          accent2: { kind: "srgb", hex: "ED7D31" },
        },
      },
      fontScheme: {
        majorLatin: "Aptos Display",
        minorLatin: "Aptos",
        majorEastAsian: "",
        majorComplexScript: "",
      },
    });
    const themeXml = decoder.decode(getEntry(writePptx(source), "ppt/theme/theme1.xml"));
    expect(themeXml).toContain(`<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>`);
    expect(readPptx(writePptx(source)).themes[0]?.fontScheme).toMatchObject({
      majorEastAsian: "",
      minorEastAsian: "",
      majorComplexScript: "",
      minorComplexScript: "",
    });

    expect(() => createPptx({ theme: { colorScheme: { accent1: "12345" } } })).toThrow(
      /theme\.colorScheme\.accent1 must be a 6-digit RGB color/,
    );
    expect(() => createPptx({ theme: { fontScheme: { major: { latin: "" } } } })).toThrow(
      /theme\.fontScheme\.major\.latin must be a non-empty string/,
    );
    expect(() =>
      createPptx({ theme: { fontScheme: { minor: { eastAsian: "bad\nfont" } } } }),
    ).toThrow(/forbidden in an XML attribute/);
  });

  it("keeps empty script typefaces distinct from Jpan fallback during computation", () => {
    let source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(100000),
      offsetY: asEmu(100000),
      width: asEmu(1000000),
      height: asEmu(500000),
      paragraphs: [
        { runs: [{ text: "East Asian", properties: { fontFace: "+mj-ea" } }] },
        { runs: [{ text: "Complex", properties: { fontFace: "+mj-cs" } }] },
      ],
    });
    const directRuns = createComputedView(source).slides[0]?.elements[0];
    expect(directRuns).toMatchObject({
      kind: "shape",
      textBody: {
        paragraphs: [
          { runs: [{ properties: { typeface: "+mj-ea" } }] },
          { runs: [{ properties: { typeface: "+mj-cs" } }] },
        ],
      },
    });

    const archive = unzipSync(writePptx(source));
    const themePath = "ppt/theme/theme1.xml";
    const themeXml = decoder
      .decode(archive[themePath])
      .replace(
        `<a:ea typeface=""/>`,
        `<a:ea typeface=""/><a:font script="Jpan" typeface="Japanese Fallback"/>`,
      );
    const reread = readPptx(zipSync({ ...archive, [themePath]: encoder.encode(themeXml) }));
    expect(reread.themes[0]?.fontScheme).toMatchObject({
      majorEastAsian: "",
      majorJapanese: "Japanese Fallback",
    });
    expect(createComputedView(reread).slides[0]?.elements[0]).toMatchObject({
      kind: "shape",
      textBody: {
        paragraphs: [
          { runs: [{ properties: { typeface: "Japanese Fallback" } }] },
          { runs: [{ properties: { typeface: "+mj-cs" } }] },
        ],
      },
    });
  });

  it("preserves alternating shape and picture sibling order after write and reread", () => {
    let source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
    const shape = (name: string) => ({
      geometry: { kind: "preset" as const, preset: "rect" },
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      name,
    });

    source = addShape(source, slideHandle, shape("First shape"));
    source = addPicture(source, slideHandle, {
      bytes: RED_PNG,
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      name: "Middle picture",
    });
    source = addShape(source, slideHandle, shape("Last shape"));

    const output = writePptx(source);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const spTreeXml = slideXml.slice(
      slideXml.indexOf("<p:spTree"),
      slideXml.indexOf("</p:spTree>"),
    );

    expect(spTreeXml.match(/<p:(?:sp|pic|graphicFrame)(?=[ >])/g)).toEqual([
      "<p:sp",
      "<p:pic",
      "<p:sp",
    ]);
    expect(readPptx(output).slides[0]?.shapes.map((item) => item.kind)).toEqual([
      "shape",
      "image",
      "shape",
    ]);

    const persisted = readPptx(output);
    const firstShapeHandle = persisted.slides[0]?.shapes[0]?.handle;
    if (firstShapeHandle === undefined) throw new Error("first shape handle should exist");
    const deletedOutput = writePptx(deleteShape(persisted, firstShapeHandle));
    const deletedSlideXml = decoder.decode(getEntry(deletedOutput, "ppt/slides/slide1.xml"));
    const deletedSpTreeXml = deletedSlideXml.slice(
      deletedSlideXml.indexOf("<p:spTree"),
      deletedSlideXml.indexOf("</p:spTree>"),
    );
    expect(deletedSpTreeXml.match(/<p:(?:sp|pic|graphicFrame)(?=[ >])/g)).toEqual([
      "<p:pic",
      "<p:sp",
    ]);
  });

  it("preserves shape, picture, chart, and table sibling order after write and reread", () => {
    let source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      name: "Ordered shape",
    });
    source = addPicture(source, slideHandle, {
      bytes: RED_PNG,
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      name: "Ordered picture",
    });
    source = addChart(source, slideHandle, {
      chartType: "bar",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      series: [{ categories: ["A"], values: [1] }],
      name: "Ordered chart",
    });
    source = addTable(source, slideHandle, {
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      columnWidths: [asEmu(1000)],
      rows: [{ height: asEmu(1000), cells: [{ text: "Cell" }] }],
      name: "Ordered table",
    });

    const output = writePptx(source);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const spTreeXml = slideXml.slice(
      slideXml.indexOf("<p:spTree"),
      slideXml.indexOf("</p:spTree>"),
    );

    expect(spTreeXml.match(/<p:(?:sp|pic|graphicFrame)(?=[ >])/g)).toEqual([
      "<p:sp",
      "<p:pic",
      "<p:graphicFrame",
      "<p:graphicFrame",
    ]);
    expect(readPptx(output).slides[0]?.shapes.map((item) => item.kind)).toEqual([
      "shape",
      "image",
      "chart",
      "table",
    ]);
  });

  it("writes all native chart types with editable workbooks and consistent package metadata", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
    const types = ["bar", "line", "pie", "area", "doughnut", "radar"] as const;
    const edited = types.reduce(
      (current, chartType, index) =>
        addChart(current, slideHandle, {
          chartType,
          offsetX: asEmu(index * 1200000),
          offsetY: asEmu(200000),
          width: asEmu(1100000),
          height: asEmu(1800000),
          title: `${chartType} title`,
          showLegend: true,
          legendPosition: "b",
          series: [
            { name: "Revenue", categories: ["Q1", "Q2"], values: [10, 20], color: "#4472C4" },
            { name: "Cost", categories: ["Q1", "Q2"], values: [7, 12], color: "ED7D31" },
          ],
          ...(chartType === "radar" ? { radarStyle: "filled" as const } : {}),
          ...(chartType === "doughnut" ? { holeSize: 60 } : {}),
          ...(chartType === "bar"
            ? {
                categoryAxis: { hidden: true, lineVisible: false, gridLinesVisible: false },
                valueAxis: { hidden: true, lineVisible: false, gridLinesVisible: false },
                plotLayout: { x: 0, y: 0, width: 1, height: 1 },
              }
            : {}),
        }),
      source,
    );

    const output = writePptx(edited);
    const archive = unzipSync(output);
    const contentTypes = decoder.decode(archive["[Content_Types].xml"]);
    const slideXml = decoder.decode(archive["ppt/slides/slide1.xml"]);
    const slideRels = decoder.decode(archive["ppt/slides/_rels/slide1.xml.rels"]);
    expect(contentTypes.match(/drawingml\.chart\+xml/g)).toHaveLength(6);
    expect(contentTypes.match(/spreadsheetml\.sheet/g)).toHaveLength(6);
    expect(slideXml.match(/<c:chart /g)).toHaveLength(6);
    expect(slideRels.match(/relationships\/chart/g)).toHaveLength(6);

    for (let index = 1; index <= 6; index += 1) {
      const chartXml = decoder.decode(archive[`ppt/charts/chart${index}.xml`]);
      const chartRels = decoder.decode(archive[`ppt/charts/_rels/chart${index}.xml.rels`]);
      const workbook = archive[`ppt/embeddings/Microsoft_Excel_Worksheet${index}.xlsx`];
      expect(chartXml).toContain(`<c:${types[index - 1]}Chart>`);
      expect(chartXml).toContain("Sheet1!$B$2:$B$3");
      expect(chartXml).toContain(`<c:v>Revenue</c:v>`);
      expect(chartXml).toContain(`<c:v>Q2</c:v>`);
      expect(chartXml).toContain(`<c:v>20</c:v>`);
      expect(chartRels).toContain(`Target="../embeddings/Microsoft_Excel_Worksheet${index}.xlsx"`);
      const worksheet = decoder.decode(unzipSync(workbook)["xl/worksheets/sheet1.xml"]);
      expect(worksheet).toContain(`<t>Revenue</t>`);
      expect(worksheet).toContain(`<t>Q2</t>`);
      expect(worksheet).toContain(`<c r="B3"><v>20</v></c>`);
    }
    expect(decoder.decode(archive["ppt/charts/chart1.xml"])).toContain(`<c:manualLayout>`);
    expect(decoder.decode(archive["ppt/charts/chart1.xml"])).toContain(`<c:delete val="1"/>`);
    expect(decoder.decode(archive["ppt/charts/chart5.xml"])).toContain(`<c:holeSize val="60"/>`);
    expect(decoder.decode(archive["ppt/charts/chart6.xml"])).toContain(
      `<c:radarStyle val="filled"/>`,
    );

    const reread = readPptx(output);
    expect(reread.diagnostics).toEqual([]);
    expect(createComputedView(reread).slides[0]?.elements.map((element) => element.kind)).toEqual(
      types.map(() => "chart"),
    );
    expect(
      createComputedView(reread).slides[0]?.elements.map((element) =>
        element.kind === "chart" ? element.chartData?.chartType : undefined,
      ),
    ).toEqual(types);
  });

  it("writes typed chart, axis, series, marker, and data point formatting", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
    const color = (hex: string, alpha?: number) => ({
      kind: "srgb" as const,
      hex,
      ...(alpha === undefined
        ? {}
        : { transforms: [{ kind: "alpha" as const, value: asOoxmlPercent(alpha) }] }),
    });
    const solid = (hex: string, alpha?: number) => ({
      kind: "solid" as const,
      color: color(hex, alpha),
    });

    const formatted = addChart(source, slideHandle, {
      chartType: "bar",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(4000000),
      height: asEmu(2500000),
      title: "Formatted chart",
      titleStyle: {
        fontFace: "Aptos Display",
        fontSize: asPt(18),
        color: color("112233"),
        bold: true,
        italic: true,
      },
      displayBlanksAs: "span",
      roundedCorners: true,
      chartArea: {
        fill: solid("F0F0F0", 80000),
        outline: { width: asEmu(12700), fill: solid("223344"), dash: "dash" },
      },
      plotArea: {
        fill: solid("FFFFFF"),
        outline: { fill: { kind: "none" } },
      },
      categoryAxis: {
        hidden: true,
        majorTickMark: "inside",
        labelPosition: "low",
        numberFormat: { formatCode: "0.0%", sourceLinked: false },
        line: { width: asEmu(19050), fill: solid("445566") },
        majorGridline: { fill: solid("D9D9D9"), dash: "dot" },
        gridLinesVisible: true,
        textStyle: {
          fontFace: "Aptos",
          fontSize: asPt(9),
          color: color("667788"),
          bold: true,
          italic: false,
        },
        showMultiLevelLabels: false,
      },
      valueAxis: {
        hidden: true,
        lineVisible: false,
        majorTickMark: "outside",
        labelPosition: "none",
        numberFormat: { formatCode: "#,##0" },
        line: { fill: solid("FF0000") },
        gridLinesVisible: false,
      },
      plotLayout: { coordinateMode: "edge", x: 0, y: 0, width: 1, height: 1 },
      series: [
        {
          name: "Revenue",
          categories: ["Q1", "Q2"],
          values: [10, 20],
          fill: solid("4472C4"),
          outline: { width: asEmu(12700), fill: solid("203864"), dash: "solid" },
          dataPoints: [
            { index: 0, fill: solid("ED7D31"), outline: { fill: solid("843C0C") } },
            { index: 1, fill: solid("70AD47") },
          ],
        },
      ],
    });
    const withLine = addChart(formatted, slideHandle, {
      chartType: "line",
      offsetX: asEmu(4000000),
      offsetY: asEmu(0),
      width: asEmu(4000000),
      height: asEmu(2500000),
      series: [
        {
          categories: ["Q1", "Q2"],
          values: [4, 8],
          fill: { kind: "none" },
          outline: { width: asEmu(25400), fill: solid("5B9BD5"), dash: "dashDot" },
          marker: {
            symbol: "diamond",
            size: 9,
            fill: solid("FFC000"),
            outline: { fill: solid("7F6000") },
          },
        },
      ],
    });
    const withArea = addChart(withLine, slideHandle, {
      chartType: "area",
      offsetX: asEmu(0),
      offsetY: asEmu(2500000),
      width: asEmu(2500000),
      height: asEmu(2000000),
      series: [
        {
          categories: ["Q1", "Q2"],
          values: [2, 3],
          fill: solid("A5A5A5"),
          outline: { fill: solid("404040") },
        },
      ],
    });
    const withRadar = addChart(withArea, slideHandle, {
      chartType: "radar",
      offsetX: asEmu(2500000),
      offsetY: asEmu(2500000),
      width: asEmu(2500000),
      height: asEmu(2000000),
      series: [
        {
          categories: ["Q1", "Q2"],
          values: [5, 6],
          fill: solid("8064A2"),
          outline: { fill: solid("4F3B66") },
          marker: { symbol: "triangle", size: 7, fill: solid("8064A2") },
        },
      ],
    });
    const withPie = addChart(withRadar, slideHandle, {
      chartType: "pie",
      offsetX: asEmu(5000000),
      offsetY: asEmu(2500000),
      width: asEmu(1800000),
      height: asEmu(2000000),
      displayBlanksAs: "zero",
      series: [
        {
          categories: ["A", "B"],
          values: [1, 2],
          dataPoints: [
            { index: 0, fill: solid("C00000") },
            { index: 1, fill: solid("00B050"), outline: { fill: solid("006100") } },
          ],
        },
      ],
    });
    const edited = addChart(withPie, slideHandle, {
      chartType: "doughnut",
      offsetX: asEmu(6800000),
      offsetY: asEmu(2500000),
      width: asEmu(1800000),
      height: asEmu(2000000),
      displayBlanksAs: "gap",
      series: [
        {
          categories: ["A", "B"],
          values: [1, 2],
          dataPoints: [{ index: 0, fill: solid("00B0F0") }],
        },
      ],
    });

    const output = writePptx(edited);
    const archive = unzipSync(output);
    const barXml = decoder.decode(archive["ppt/charts/chart1.xml"]);
    const titleXml = /<c:title>.*?<\/c:title>/.exec(barXml)?.[0];
    const plotAreaXml = /<c:plotArea>.*?<\/c:plotArea>/.exec(barXml)?.[0];
    const barSeriesXml = /<c:ser>.*?<\/c:ser>/.exec(barXml)?.[0];
    expect(barXml).toContain(`<c:roundedCorners val="1"/>`);
    expect(barXml).toContain(`<c:dispBlanksAs val="span"/>`);
    expect(titleXml).toContain(`<a:rPr lang="en-US" sz="1800" b="1" i="1">`);
    expect(titleXml).toContain(`<a:srgbClr val="112233">`);
    expect(titleXml).toContain(`<a:latin typeface="Aptos Display"/>`);
    expect(plotAreaXml).toContain(
      `<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>`,
    );
    expect(barSeriesXml).toContain(
      `<c:spPr><a:solidFill><a:srgbClr val="4472C4"></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="203864"></a:srgbClr></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr>`,
    );
    expect(barXml).toContain(
      `<c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="edge"/><c:hMode val="edge"/><c:x val="0"/><c:y val="0"/><c:w val="1"/><c:h val="1"/></c:manualLayout>`,
    );
    const categoryAxisXml = /<c:catAx>.*?<\/c:catAx>/.exec(barXml)?.[0];
    const valueAxisXml = /<c:valAx>.*?<\/c:valAx>/.exec(barXml)?.[0];
    expect(categoryAxisXml).toContain(`<c:delete val="1"/>`);
    expect(categoryAxisXml).toContain(`<c:majorTickMark val="in"/>`);
    expect(categoryAxisXml).toContain(`<c:tickLblPos val="low"/>`);
    expect(categoryAxisXml).toContain(`<c:numFmt formatCode="0.0%" sourceLinked="0"/>`);
    expect(categoryAxisXml).toContain(`<c:majorGridlines><c:spPr>`);
    expect(categoryAxisXml).toContain(`<a:ln w="19050">`);
    expect(categoryAxisXml).toContain(`<a:defRPr sz="900" b="1" i="0">`);
    expect(categoryAxisXml).toContain(`<c:noMultiLvlLbl val="1"/>`);
    expect(valueAxisXml).toContain(`<c:delete val="1"/>`);
    expect(valueAxisXml).toContain(`<c:majorTickMark val="out"/>`);
    expect(valueAxisXml).toContain(`<c:tickLblPos val="none"/>`);
    expect(valueAxisXml).toContain(`<c:numFmt formatCode="#,##0" sourceLinked="1"/>`);
    expect(valueAxisXml).not.toContain(`<c:majorGridlines>`);
    expect(valueAxisXml).toContain(`<a:ln><a:noFill/></a:ln>`);
    expect(barXml).toContain(
      `<c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="ED7D31"></a:srgbClr></a:solidFill><a:ln><a:solidFill><a:srgbClr val="843C0C"></a:srgbClr></a:solidFill></a:ln></c:spPr></c:dPt>`,
    );
    expect(barXml).toContain(
      `<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="70AD47"></a:srgbClr></a:solidFill></c:spPr></c:dPt>`,
    );
    expect(barXml).toContain(
      `</c:chart><c:spPr><a:solidFill><a:srgbClr val="F0F0F0"><a:alpha val="80000"/></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="223344"></a:srgbClr></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr><c:externalData`,
    );
    const lineXml = decoder.decode(archive["ppt/charts/chart2.xml"]);
    expect(lineXml).toContain(`<c:symbol val="diamond"/><c:size val="9"/>`);
    expect(lineXml).toContain(
      `<c:spPr><a:solidFill><a:srgbClr val="FFC000"></a:srgbClr></a:solidFill><a:ln><a:solidFill><a:srgbClr val="7F6000"></a:srgbClr></a:solidFill></a:ln></c:spPr>`,
    );
    expect(lineXml).toContain(`<a:prstDash val="dashDot"/>`);
    const areaXml = decoder.decode(archive["ppt/charts/chart3.xml"]);
    expect(areaXml).toContain(
      `<a:solidFill><a:srgbClr val="A5A5A5"></a:srgbClr></a:solidFill><a:ln><a:solidFill><a:srgbClr val="404040"></a:srgbClr></a:solidFill></a:ln>`,
    );
    const radarXml = decoder.decode(archive["ppt/charts/chart4.xml"]);
    expect(radarXml).toContain(`<c:symbol val="triangle"/><c:size val="7"/>`);
    expect(radarXml).toContain(
      `<a:solidFill><a:srgbClr val="8064A2"></a:srgbClr></a:solidFill><a:ln><a:solidFill><a:srgbClr val="4F3B66"></a:srgbClr></a:solidFill></a:ln>`,
    );
    const pieXml = decoder.decode(archive["ppt/charts/chart5.xml"]);
    expect(pieXml).toContain(`<c:dispBlanksAs val="zero"/>`);
    expect(pieXml).toContain(
      `<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="00B050"></a:srgbClr></a:solidFill><a:ln><a:solidFill><a:srgbClr val="006100"></a:srgbClr></a:solidFill></a:ln></c:spPr></c:dPt>`,
    );
    const doughnutXml = decoder.decode(archive["ppt/charts/chart6.xml"]);
    expect(doughnutXml).toContain(`<c:dispBlanksAs val="gap"/>`);
    expect(doughnutXml).toContain(
      `<c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="00B0F0"></a:srgbClr></a:solidFill></c:spPr></c:dPt>`,
    );

    const reread = readPptx(output);
    expect(reread.diagnostics).toEqual([]);
    expect(
      reread.packageGraph.parts.filter((part) =>
        /^ppt\/charts\/chart\d+\.xml$/.test(part.partPath),
      ),
    ).toHaveLength(6);
    expect(
      reread.packageGraph.parts.filter((part) => part.partPath.includes("/embeddings/")),
    ).toHaveLength(6);
    expect(
      createComputedView(reread).slides[0]?.elements.map((element) =>
        element.kind === "chart" ? element.chartData?.chartType : undefined,
      ),
    ).toEqual(["bar", "line", "area", "radar", "pie", "doughnut"]);
  });

  it("avoids the shape-tree root ID and rejects package-breaking chart inputs", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const edited = addChart(source, source.slides[0].handle!, {
      chartType: "bar",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      series: [{ categories: [" A "], values: [1], name: " Series " }],
    });
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const chartXml = decoder.decode(getEntry(output, "ppt/charts/chart1.xml"));
    const worksheet = decoder.decode(
      unzipSync(getEntry(output, "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"))[
        "xl/worksheets/sheet1.xml"
      ],
    );
    expect(slideXml).toContain(`<p:cNvPr id="31" name="Chart 31"/>`);
    expect(chartXml).toContain(`<c:v xml:space="preserve"> Series </c:v>`);
    expect(worksheet).toContain(`<t xml:space="preserve"> A </t>`);

    const valid = {
      chartType: "bar",
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      series: [{ categories: ["A"], values: [1] }],
    } as const;
    expect(() =>
      addChart(source, source.slides[0].handle!, { ...valid, title: "bad\u0000" }),
    ).toThrow(/forbidden by XML 1.0/);
    const invalidFillKind = { ...valid, chartArea: { fill: { kind: "none" as const } } };
    Reflect.set(invalidFillKind.chartArea.fill, "kind", "gradient");
    expect(() => addChart(source, source.slides[0].handle!, invalidFillKind)).toThrow(
      /unsupported chartArea\.fill\.kind/,
    );
    const invalidColorKind = {
      ...valid,
      chartArea: {
        fill: { kind: "solid" as const, color: { kind: "srgb" as const, hex: "FFFFFF" } },
      },
    };
    Reflect.set(invalidColorKind.chartArea.fill.color, "kind", "scheme");
    expect(() => addChart(source, source.slides[0].handle!, invalidColorKind)).toThrow(
      /unsupported chartArea\.fill\.color\.kind/,
    );
    expect(() =>
      addChart(source, source.slides[0].handle!, {
        ...valid,
        chartArea: { outline: { width: asEmu(20_116_801) } },
      }),
    ).toThrow(/width must be a finite EMU value/);
    expect(() =>
      addChart(source, source.slides[0].handle!, {
        ...valid,
        chartArea: { outline: { width: asEmu(0.5) } },
      }),
    ).toThrow(/width must be a finite EMU value/);
    expect(() =>
      addChart(source, source.slides[0].handle!, {
        ...valid,
        title: "Too small",
        titleStyle: { fontSize: asPt(0.5) },
      }),
    ).toThrow(/fontSize must be from 1 through 4000 points/);
  });

  it("writes a new presentation after adding a text box through public APIs", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    const edited = addTextBox(source, slideHandle, {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(3657600),
      height: asEmu(914400),
      text: "Hello from scratch",
    });

    const reread = readPptx(writePptx(edited));
    expect(reread.diagnostics).toEqual([]);
    expect(reread.presentation.slidePartPaths).toEqual([asPartPath("ppt/slides/slide1.xml")]);
    expect(reread.slides).toHaveLength(1);
    expect(reread.slideLayouts).toHaveLength(1);
    expect(reread.slideMasters).toHaveLength(1);
    expect(reread.themes).toHaveLength(1);
    expect(findTextRun(reread, "Hello from scratch").text).toBe("Hello from scratch");

    const computed = createComputedView(reread);
    expect(computed.slideSize).toEqual({ width: asEmu(9144000), height: asEmu(5143500) });
    expect(computed.slides[0]?.elements).toHaveLength(1);
  });

  it("authors a named master and layout with backgrounds, objects, slide numbers, and margins", () => {
    let source = createPptx({
      slideMaster: { name: "Product Master", background: { kind: "image", bytes: BLUE_PNG } },
      slideLayout: {
        name: "Product Blank",
        margin: {
          left: asEmu(120000),
          right: asEmu(130000),
          top: asEmu(140000),
          bottom: asEmu(150000),
        },
      },
    });
    const masterHandle = source.slideMasters[0]?.handle;
    const layout = source.slideLayouts[0];
    if (masterHandle === undefined || layout?.handle === undefined) {
      throw new Error("createPptx should create master and layout handles");
    }
    source = addTextBox(source, masterHandle, {
      offsetX: asEmu(200000),
      offsetY: asEmu(100000),
      width: asEmu(1800000),
      height: asEmu(400000),
      text: "Master title",
    });
    source = addShape(source, masterHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(0),
      offsetY: asEmu(5000000),
      width: asEmu(9144000),
      height: asEmu(143500),
      fill: { kind: "solid", color: { kind: "srgb", hex: "112233" } },
    });
    source = addConnector(source, masterHandle, {
      preset: "straightConnector1",
      offsetX: asEmu(200000),
      offsetY: asEmu(700000),
      width: asEmu(1800000),
      height: asEmu(1),
    });
    source = addPicture(source, masterHandle, {
      bytes: RED_PNG,
      offsetX: asEmu(8200000),
      offsetY: asEmu(100000),
      width: asEmu(600000),
      height: asEmu(300000),
    });
    source = addSlideNumber(source, masterHandle, {
      offsetX: asEmu(8200000),
      offsetY: asEmu(4700000),
      width: asEmu(600000),
      height: asEmu(300000),
      align: "right",
      properties: { fontFace: "Aptos", fontSize: asPt(10), color: { kind: "srgb", hex: "334455" } },
    });
    source = addEmptySlideFromLayout(source, { layoutPartPath: layout.partPath });
    source = addEmptySlideFromLayout(source, { layoutPartPath: layout.partPath });
    source = addTextBox(source, source.slides[1].handle!, {
      offsetX: asEmu(500000),
      offsetY: asEmu(500000),
      width: asEmu(2000000),
      height: asEmu(500000),
      text: "Uses layout margins",
      body: { marginRight: asEmu(999000) },
    });

    const output = writePptx(source);
    const masterXml = decoder.decode(getEntry(output, "ppt/slideMasters/slideMaster1.xml"));
    const masterRels = decoder.decode(
      getEntry(output, "ppt/slideMasters/_rels/slideMaster1.xml.rels"),
    );
    const slide2Xml = decoder.decode(getEntry(output, "ppt/slides/slide2.xml"));
    const reread = readPptx(output);

    expect(source.presentation.slidePartPaths).toHaveLength(3);
    expect(source.slideMasters[0]).toMatchObject({ name: "Product Master" });
    expect(source.slideLayouts[0]).toMatchObject({
      name: "Product Blank",
      defaultTextBodyProperties: {
        marginLeft: 120000,
        marginRight: 130000,
        marginTop: 140000,
        marginBottom: 150000,
      },
    });
    expect(masterXml).toContain(`<p:cSld name="Product Master"><p:bg>`);
    expect(masterXml).toContain(`<a:blip r:embed="rId3"/>`);
    expect(masterXml).toContain(`type="slidenum"`);
    expect(masterXml.match(/<p:cNvPr id="[1-5]"/g)).toHaveLength(5);
    expect(masterRels).toContain(`Id="rId3"`);
    expect(masterRels).toContain(`Target="../media/image1.png"`);
    expect(masterRels).toContain(`Id="rId4"`);
    expect(masterRels).toContain(`Target="../media/image2.png"`);
    expect(slide2Xml).toContain(`lIns="120000"`);
    expect(slide2Xml).toContain(`rIns="999000"`);
    expect(slide2Xml).toContain(`tIns="140000"`);
    expect(slide2Xml).toContain(`bIns="150000"`);
    expect(reread.diagnostics).toEqual([]);
    expect(reread.slideMasters[0]).toMatchObject({ name: "Product Master" });
    expect(reread.slideLayouts[0]).toMatchObject({ name: "Product Blank" });
    expect(reread.slideMasters[0]?.shapes).toHaveLength(5);
    expect(createComputedView(reread).slides.map((slide) => slide.elements.length)).toEqual([
      5, 6, 5,
    ]);
    expect(reread.packageGraph.contentTypes.defaults).toContainEqual({
      extension: "png",
      contentType: "image/png",
    });
  });

  it("authors a solid master background", () => {
    const output = writePptx(
      createPptx({
        slideMaster: {
          name: "Solid Master",
          background: { kind: "solid", color: { kind: "srgb", hex: "F8FAFC" } },
        },
      }),
    );
    const reread = readPptx(output);
    expect(reread.slideMasters[0]).toMatchObject({
      name: "Solid Master",
      background: {
        kind: "fill",
        fill: { kind: "solid", color: { kind: "srgb", hex: "F8FAFC" } },
      },
    });
  });

  it("authors solid, linear, radial, PNG, and JPEG backgrounds on individual slides", () => {
    let source = createPptx();
    const masterHandle = source.slideMasters[0]?.handle;
    const layout = source.slideLayouts[0];
    if (masterHandle === undefined || layout === undefined) {
      throw new Error("createPptx should create a master and layout");
    }
    source = addShape(source, masterHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(100000),
      offsetY: asEmu(100000),
      width: asEmu(500000),
      height: asEmu(500000),
      fill: { kind: "solid", color: { kind: "srgb", hex: "FFFFFF" } },
    });
    for (let index = 0; index < 4; index += 1) {
      source = addEmptySlideFromLayout(source, { layoutPartPath: layout.partPath });
    }
    const handles = source.slides.map((slide) => slide.handle);
    if (handles.some((handle) => handle === undefined)) {
      throw new Error("authored slides should have handles");
    }

    source = setSlideBackground(source, handles[0]!, {
      kind: "solid",
      color: { kind: "srgb", hex: "112233" },
    });
    source = setSlideBackground(source, handles[1]!, {
      kind: "gradient",
      gradientType: "linear",
      angle: asOoxmlAngle(2700000),
      stops: [
        { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FF0000" } },
        { position: asOoxmlPercent(100000), color: { kind: "srgb", hex: "0000FF" } },
      ],
    });
    source = setSlideBackground(source, handles[2]!, {
      kind: "gradient",
      gradientType: "radial",
      centerX: asOoxmlPercent(25000),
      centerY: asOoxmlPercent(75000),
      stops: [
        { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FFFFFF" } },
        { position: asOoxmlPercent(100000), color: { kind: "srgb", hex: "00AA44" } },
      ],
    });
    source = setSlideBackground(source, handles[3]!, { kind: "image", bytes: BLUE_PNG });
    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 1, 2, 3]);
    source = setSlideBackground(source, handles[4]!, { kind: "image", bytes: jpeg });

    const output = writePptx(source);
    const archive = unzipSync(output);
    const slideXml = Array.from({ length: 5 }, (_, index) =>
      decoder.decode(archive[`ppt/slides/slide${index + 1}.xml`]),
    );
    const pngRels = decoder.decode(archive["ppt/slides/_rels/slide4.xml.rels"]);
    const jpegRels = decoder.decode(archive["ppt/slides/_rels/slide5.xml.rels"]);
    const contentTypes = decoder.decode(archive["[Content_Types].xml"]);
    const reread = readPptx(output);

    expect(source.slides.map((slide) => slide.background?.kind)).toEqual([
      "fill",
      "fill",
      "fill",
      "fill",
      "fill",
    ]);
    expect(slideXml[0]).toContain(
      `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill>`,
    );
    expect(slideXml[1]).toContain(`<a:lin ang="2700000" scaled="1"/>`);
    expect(slideXml[2]).toContain(
      `<a:path path="circle"><a:fillToRect l="25000" t="75000" r="75000" b="25000"/></a:path>`,
    );
    for (const xml of slideXml) {
      expect(xml.indexOf("<p:bg>")).toBeLessThan(xml.indexOf("<p:spTree>"));
    }
    expect(pngRels).toContain(`Target="../media/image1.png"`);
    expect(jpegRels).toContain(`Target="../media/image1.jpeg"`);
    expect(archive["ppt/media/image1.png"]).toEqual(BLUE_PNG);
    expect(archive["ppt/media/image1.jpeg"]).toEqual(jpeg);
    expect(contentTypes).toContain(`<Default Extension="png" ContentType="image/png"/>`);
    expect(contentTypes).toContain(`<Default Extension="jpeg" ContentType="image/jpeg"/>`);
    expect(reread.diagnostics).toEqual([]);
    expect(reread.slides[1]?.background).toMatchObject({
      kind: "fill",
      fill: { kind: "gradient", gradientType: "linear", angle: 2700000 },
    });
    expect(reread.slides[2]?.background).toMatchObject({
      kind: "fill",
      fill: { kind: "gradient", gradientType: "radial", centerX: 0.25, centerY: 0.75 },
    });
    expect(reread.slides[3]?.background).toMatchObject({
      kind: "fill",
      fill: { kind: "image", blipRelationshipId: "rId2" },
    });
    expect(createComputedView(reread).slides.every((slide) => slide.elements.length === 1)).toBe(
      true,
    );
    expect(
      createComputedView(reread).slides.every(
        (slide) => slide.elements[0]?.sourceLayer === "master",
      ),
    ).toBe(true);
  });

  it("rejects invalid slide background inputs and non-slide handles", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    const masterHandle = source.slideMasters[0]?.handle;
    if (slideHandle === undefined || masterHandle === undefined) {
      throw new Error("createPptx should create slide and master handles");
    }
    expect(() =>
      setSlideBackground(source, slideHandle, {
        kind: "gradient",
        gradientType: "linear",
        angle: asOoxmlAngle(0),
        stops: [],
      }),
    ).toThrow(/at least two/);
    expect(() =>
      setSlideBackground(source, slideHandle, {
        kind: "image",
        bytes: new Uint8Array([1, 2, 3]),
      }),
    ).toThrow(/unsupported or unknown image format/);
    expect(() =>
      setSlideBackground(source, masterHandle, {
        kind: "solid",
        color: { kind: "srgb", hex: "FFFFFF" },
      }),
    ).toThrow(/slide handle was not found/);

    expect(() => {
      Reflect.apply(writePptx, undefined, [
        {
          ...source,
          edits: [
            {
              kind: "setBackground",
              targetPartPath: source.slides[0].partPath,
              relationshipId: "rId2",
              xml: "<p:bg/>",
            },
          ],
        },
      ]);
    }).toThrow(/relationship, media part, and content type must be provided together/);
  });

  it("sets and clears existing master and layout backgrounds through the common API", () => {
    const base = readPptx(writePptx(createPptx()));
    const masterHandle = base.slideMasters[0]?.handle;
    const layoutHandle = base.slideLayouts[0]?.handle;
    if (masterHandle === undefined || layoutHandle === undefined) {
      throw new Error("createPptx should create master and layout handles");
    }

    const masterOutput = writePptx(
      setBackground(base, masterHandle, {
        kind: "gradient",
        gradientType: "linear",
        angle: asOoxmlAngle(5400000),
        stops: [
          { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "112233" } },
          { position: asOoxmlPercent(100000), color: { kind: "srgb", hex: "445566" } },
        ],
      }),
    );
    const masterReread = readPptx(masterOutput);
    expect(masterReread.slideMasters[0]?.background).toMatchObject({
      kind: "fill",
      fill: { kind: "gradient", gradientType: "linear", angle: 5400000 },
    });
    expect(createComputedView(masterReread).slides[0]?.background?.sourceLayer).toBe("master");

    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 1, 2, 3]);
    const layoutOutput = writePptx(
      setBackground(masterReread, requireHandle(masterReread.slideLayouts[0]?.handle), {
        kind: "image",
        bytes: jpeg,
      }),
    );
    const layoutArchive = unzipSync(layoutOutput);
    const layoutReread = readPptx(layoutOutput);
    expect(layoutReread.slideLayouts[0]?.background).toMatchObject({
      kind: "fill",
      fill: { kind: "image" },
    });
    expect(createComputedView(layoutReread).slides[0]?.background?.sourceLayer).toBe("layout");
    expect(layoutArchive["ppt/media/image1.jpeg"]).toEqual(jpeg);

    const clearedLayout = readPptx(
      writePptx(clearBackground(layoutReread, requireHandle(layoutReread.slideLayouts[0]?.handle))),
    );
    expect(clearedLayout.slideLayouts[0]?.background).toBeUndefined();
    expect(createComputedView(clearedLayout).slides[0]?.background?.sourceLayer).toBe("master");
    expect(createComputedView(clearedLayout).slides[0]?.background?.source).toMatchObject({
      kind: "fill",
      fill: { kind: "gradient", gradientType: "linear" },
    });

    const clearedOutput = writePptx(
      clearBackground(masterReread, requireHandle(masterReread.slideMasters[0]?.handle)),
    );
    const clearedArchive = unzipSync(clearedOutput);
    const clearedMasterXml = decoder.decode(clearedArchive["ppt/slideMasters/slideMaster1.xml"]);
    const clearedReread = readPptx(clearedOutput);
    expect(clearedMasterXml).not.toContain("<p:bg>");
    expect(clearedReread.slideMasters[0]?.background).toBeUndefined();
  });

  it("adds PNG media to a master while preserving old media relationships and raw siblings", () => {
    const initialArchive = unzipSync(writePptx(createPptx()));
    const masterPartPath = "ppt/slideMasters/slideMaster1.xml";
    const masterXml = decoder
      .decode(initialArchive[masterPartPath])
      .replace("</p:sldMaster>", '<p:extLst><p:ext uri="keep"/></p:extLst></p:sldMaster>');
    const base = readPptx(
      zipSync({ ...initialArchive, [masterPartPath]: encoder.encode(masterXml) }),
    );
    const withImage = writePptx(
      setBackground(base, requireHandle(base.slideMasters[0]?.handle), {
        kind: "image",
        bytes: BLUE_PNG,
      }),
    );
    const reread = readPptx(withImage);
    const replaced = writePptx(
      setBackground(reread, requireHandle(reread.slideMasters[0]?.handle), {
        kind: "solid",
        color: { kind: "srgb", hex: "ABCDEF" },
      }),
    );
    const archive = unzipSync(replaced);
    const writtenMasterXml = decoder.decode(archive[masterPartPath]);
    const masterRels = decoder.decode(archive["ppt/slideMasters/_rels/slideMaster1.xml.rels"]);

    expect(writtenMasterXml).toContain('<p:extLst><p:ext uri="keep"/></p:extLst>');
    expect(writtenMasterXml).toContain('<a:srgbClr val="ABCDEF"/>');
    expect(masterRels).toContain("../media/image1.png");
    expect(archive["ppt/media/image1.png"]).toEqual(BLUE_PNG);
  });

  it("preserves a non-standard PresentationML prefix when authoring a slide background", () => {
    const archive = unzipSync(writePptx(createPptx()));
    const slidePartPath = "ppt/slides/slide1.xml";
    const slideXml = decoder
      .decode(archive[slidePartPath])
      .replaceAll("p:", "x:")
      .replace("xmlns:p=", "xmlns:x=");
    const source = readPptx(zipSync({ ...archive, [slidePartPath]: encoder.encode(slideXml) }));
    const edited = setSlideBackground(source, source.slides[0].handle!, {
      kind: "solid",
      color: { kind: "srgb", hex: "112233" },
    });
    const writtenSlideXml = decoder.decode(unzipSync(writePptx(edited))[slidePartPath]);

    expect(writtenSlideXml).toContain("<x:bg><x:bgPr>");
    expect(writtenSlideXml).not.toContain("<p:bg");
  });

  it("preserves a namespace declared locally on an existing slide background", () => {
    const archive = unzipSync(writePptx(createPptx()));
    const slidePartPath = "ppt/slides/slide1.xml";
    const presentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const existingBackground =
      `<x:bg xmlns:x="${presentationNamespace}"><x:bgPr>` +
      `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/>` +
      `</x:bgPr></x:bg>`;
    const slideXml = decoder
      .decode(archive[slidePartPath])
      .replace("<p:spTree>", `${existingBackground}<p:spTree>`);
    const source = readPptx(zipSync({ ...archive, [slidePartPath]: encoder.encode(slideXml) }));
    const edited = setSlideBackground(source, source.slides[0].handle!, {
      kind: "solid",
      color: { kind: "srgb", hex: "112233" },
    });
    const output = writePptx(edited);
    const writtenSlideXml = decoder.decode(unzipSync(output)[slidePartPath]);
    const reread = readPptx(output);

    expect(writtenSlideXml).toContain(`<x:bg xmlns:x="${presentationNamespace}"><x:bgPr>`);
    expect(reread.diagnostics).toEqual([]);
    expect(reread.slides[0].background).toMatchObject({
      kind: "fill",
      fill: { kind: "solid", color: { kind: "srgb", hex: "112233" } },
    });
  });

  it("rejects authoring a background when the slide has no shape tree", () => {
    const archive = unzipSync(writePptx(createPptx()));
    const slidePartPath = "ppt/slides/slide1.xml";
    const slideXml = decoder
      .decode(archive[slidePartPath])
      .replace(/<p:spTree>[\s\S]*<\/p:spTree>/, "");
    const source = readPptx(zipSync({ ...archive, [slidePartPath]: encoder.encode(slideXml) }));
    const edited = setSlideBackground(source, source.slides[0].handle!, {
      kind: "solid",
      color: { kind: "srgb", hex: "112233" },
    });

    expect(() => writePptx(edited)).toThrow(/has no p:spTree/);
  });

  it("authors every supported object on a layout with part-unique slide-number fields", () => {
    let source = createPptx();
    const masterHandle = source.slideMasters[0]?.handle;
    const layoutHandle = source.slideLayouts[0]?.handle;
    if (masterHandle === undefined || layoutHandle === undefined) {
      throw new Error("createPptx should create master and layout handles");
    }
    source = addTextBox(source, layoutHandle, {
      offsetX: asEmu(10),
      offsetY: asEmu(20),
      width: asEmu(300),
      height: asEmu(100),
      text: "Layout text",
    });
    source = addShape(source, layoutHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(20),
      offsetY: asEmu(30),
      width: asEmu(300),
      height: asEmu(100),
    });
    source = addConnector(source, layoutHandle, {
      preset: "straightConnector1",
      offsetX: asEmu(30),
      offsetY: asEmu(40),
      width: asEmu(300),
      height: asEmu(1),
    });
    source = addPicture(source, layoutHandle, {
      bytes: RED_PNG,
      offsetX: asEmu(40),
      offsetY: asEmu(50),
      width: asEmu(300),
      height: asEmu(100),
    });
    source = addSlideNumber(source, layoutHandle, {
      offsetX: asEmu(50),
      offsetY: asEmu(60),
      width: asEmu(300),
      height: asEmu(100),
    });
    for (let index = 0; index < 4; index += 1) {
      source = addSlideNumber(source, layoutHandle, {
        offsetX: asEmu(60 + index),
        offsetY: asEmu(70),
        width: asEmu(300),
        height: asEmu(100),
      });
    }
    source = addSlideNumber(source, masterHandle, {
      offsetX: asEmu(50),
      offsetY: asEmu(60),
      width: asEmu(300),
      height: asEmu(100),
    });

    const output = writePptx(source);
    const layoutXml = decoder.decode(getEntry(output, "ppt/slideLayouts/slideLayout1.xml"));
    const masterXml = decoder.decode(getEntry(output, "ppt/slideMasters/slideMaster1.xml"));
    const layoutRels = decoder.decode(
      getEntry(output, "ppt/slideLayouts/_rels/slideLayout1.xml.rels"),
    );
    const layoutFieldIds = [...layoutXml.matchAll(/<a:fld id="([^"]+)" type="slidenum"/g)].map(
      (match) => match[1],
    );
    const masterFieldId = masterXml.match(/<a:fld id="([^"]+)" type="slidenum"/)?.[1];
    const reread = readPptx(output);

    expect(layoutXml).toContain("<p:sp>");
    expect(layoutXml).toContain("<p:cxnSp>");
    expect(layoutXml).toContain("<p:pic>");
    expect(layoutRels).toContain('Target="../media/image1.png"');
    expect(layoutFieldIds).toHaveLength(5);
    expect(new Set(layoutFieldIds).size).toBe(5);
    expect(masterFieldId).toBeDefined();
    expect(layoutFieldIds).not.toContain(masterFieldId);
    expect(reread.slideLayouts[0]?.shapes).toHaveLength(9);
    expect(reread.slideMasters[0]?.shapes).toHaveLength(1);
    expect(reread.diagnostics).toEqual([]);
  });

  it("rejects invalid master and layout authoring options", () => {
    expect(() => {
      Reflect.apply(createPptx, undefined, [{ slideMaster: { background: { kind: "pattern" } } }]);
    }).toThrow(/background\.kind must be solid or image/);
    expect(() => createPptx({ slideMaster: { name: "bad\u0000name" } })).toThrow(
      /forbidden in an XML attribute/,
    );
    expect(() => createPptx({ slideLayout: { name: "bad\nname" } })).toThrow(
      /forbidden in an XML attribute/,
    );
  });

  it("writes run hyperlinks for text boxes and shapes with slide-local relationships", () => {
    const source = createPptx();
    const firstSlideHandle = source.slides[0]?.handle;
    const layoutPartPath = source.slideLayouts[0]?.partPath;
    if (firstSlideHandle === undefined || layoutPartPath === undefined) {
      throw new Error("createPptx should create a slide and layout");
    }

    const withTextBox = addTextBox(source, firstSlideHandle, {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(3657600),
      height: asEmu(914400),
      paragraphs: [
        {
          runs: [
            { text: "plain" },
            { text: "first", hyperlink: "https://example.com/first?a=1&b=2" },
            { text: " middle", properties: { bold: true } },
            { text: " repeated", hyperlink: "https://example.com/first?a=1&b=2" },
            { text: " second", hyperlink: "http://example.com/second" },
          ],
        },
      ],
    });
    const withShape = addShape(withTextBox, firstSlideHandle, {
      geometry: { kind: "preset", preset: "roundRect" },
      offsetX: asEmu(914400),
      offsetY: asEmu(2286000),
      width: asEmu(3657600),
      height: asEmu(914400),
      paragraphs: [
        {
          runs: [
            { text: "shape plain" },
            { text: " shape link", hyperlink: "https://example.com/first?a=1&b=2" },
          ],
        },
      ],
    });
    const withSecondSlide = addEmptySlideFromLayout(withShape, { layoutPartPath });
    const secondSlideHandle = withSecondSlide.slides[1]?.handle;
    if (secondSlideHandle === undefined) throw new Error("expected a second slide");
    const edited = addTextBox(withSecondSlide, secondSlideHandle, {
      offsetX: asEmu(914400),
      offsetY: asEmu(914400),
      width: asEmu(3657600),
      height: asEmu(914400),
      paragraphs: [
        {
          runs: [
            { text: "second slide" },
            { text: " linked", hyperlink: "https://example.com/slide-two" },
          ],
        },
      ],
    });

    const firstSlideRelationships = edited.packageGraph.relationships.find(
      (relationships) => relationships.sourcePartPath === "ppt/slides/slide1.xml",
    );
    const secondSlideRelationships = edited.packageGraph.relationships.find(
      (relationships) => relationships.sourcePartPath === "ppt/slides/slide2.xml",
    );
    expect(firstSlideRelationships?.relationships).toMatchObject([
      { id: "rId1", target: "../slideLayouts/slideLayout1.xml" },
      {
        id: "rId2",
        target: "https://example.com/first?a=1&b=2",
        targetMode: "External",
      },
      { id: "rId3", target: "http://example.com/second", targetMode: "External" },
      {
        id: "rId4",
        target: "https://example.com/first?a=1&b=2",
        targetMode: "External",
      },
    ]);
    expect(secondSlideRelationships?.relationships).toMatchObject([
      { id: "rId1", target: "../slideLayouts/slideLayout1.xml" },
      { id: "rId2", target: "https://example.com/slide-two", targetMode: "External" },
    ]);

    const output = writePptx(edited);
    const firstSlideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const secondSlideXml = decoder.decode(getEntry(output, "ppt/slides/slide2.xml"));
    const firstSlideRelsXml = decoder.decode(getEntry(output, "ppt/slides/_rels/slide1.xml.rels"));
    const secondSlideRelsXml = decoder.decode(getEntry(output, "ppt/slides/_rels/slide2.xml.rels"));

    expect([...firstSlideXml.matchAll(/<a:hlinkClick r:id="rId2"\/>/g)]).toHaveLength(2);
    expect([...firstSlideXml.matchAll(/<a:hlinkClick r:id="rId3"\/>/g)]).toHaveLength(1);
    expect([...firstSlideXml.matchAll(/<a:hlinkClick r:id="rId4"\/>/g)]).toHaveLength(1);
    expect([...firstSlideXml.matchAll(/<a:hlinkClick/g)]).toHaveLength(4);
    expect([...secondSlideXml.matchAll(/<a:hlinkClick r:id="rId2"\/>/g)]).toHaveLength(1);
    expect(firstSlideRelsXml).toContain(
      'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/first?a=1&amp;b=2" TargetMode="External"',
    );
    expect(firstSlideRelsXml).toContain(
      'Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com/second" TargetMode="External"',
    );
    expect(firstSlideRelsXml).toContain(
      'Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/first?a=1&amp;b=2" TargetMode="External"',
    );
    expect(secondSlideRelsXml).toContain(
      'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/slide-two" TargetMode="External"',
    );
  });

  it("rejects non-HTTP run hyperlinks for text boxes and shapes", () => {
    const source = createPptx();
    const handle = source.slides[0]?.handle;
    if (handle === undefined) throw new Error("createPptx should create a first slide");
    const base = {
      offsetX: asEmu(0),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      paragraphs: [{ runs: [{ text: "link", hyperlink: "mailto:test@example.com" }] }],
    } as const;

    expect(() => addTextBox(source, handle, base)).toThrow(
      "addTextBox: paragraphs[0].runs[0].hyperlink must be an absolute HTTP(S) URL",
    );
    expect(() =>
      addShape(source, handle, { ...base, geometry: { kind: "preset", preset: "rect" } }),
    ).toThrow("addShape: paragraphs[0].runs[0].hyperlink must be an absolute HTTP(S) URL");
  });

  it("writes added PNG and JPEG pictures with media parts, content types, and slide rels", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    const withPng = addPicture(source, slideHandle, {
      bytes: RED_PNG,
      offsetX: asEmu(914400),
      offsetY: asEmu(457200),
      width: asEmu(1828800),
      height: asEmu(1371600),
      rotation: asOoxmlAngle(600000),
      crop: {
        left: asOoxmlPercent(1000),
        top: asOoxmlPercent(2000),
        right: asOoxmlPercent(3000),
        bottom: asOoxmlPercent(4000),
      },
      name: "Product PNG",
    });
    const edited = addPicture(withPng, slideHandle, {
      bytes: new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 1, 2, 3]),
      offsetX: asEmu(3200400),
      offsetY: asEmu(457200),
      width: asEmu(1828800),
      height: asEmu(1371600),
      name: "Product JPEG",
    });

    const output = writePptx(edited);
    const contentTypesXml = decoder.decode(getEntry(output, "[Content_Types].xml"));
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const slideRelsXml = decoder.decode(getEntry(output, "ppt/slides/_rels/slide1.xml.rels"));
    const reread = readPptx(output);

    expect(getEntry(output, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(getEntry(output, "ppt/media/image1.jpeg")).toEqual(
      new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 1, 2, 3]),
    );
    expect(contentTypesXml).toContain(`<Default Extension="png" ContentType="image/png"/>`);
    expect(contentTypesXml).toContain(`<Default Extension="jpeg" ContentType="image/jpeg"/>`);
    expect(slideRelsXml).toContain(
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`,
    );
    expect(slideRelsXml).toContain(
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/>`,
    );
    expect(slideXml).toContain(`<p:cNvPr id="1" name="Product PNG"/>`);
    expect(slideXml).toContain(`<a:blip r:embed="rId2"/>`);
    expect(slideXml).toContain(`<a:srcRect l="1000" t="2000" r="3000" b="4000"/>`);
    expect(slideXml).toContain(`<a:xfrm rot="600000">`);
    expect(slideXml).toContain(`<p:cNvPr id="2" name="Product JPEG"/>`);
    expect(slideXml).toContain(`<a:blip r:embed="rId3"/>`);
    expect(reread.packageGraph.media).toEqual([
      { partPath: "ppt/media/image1.png", contentType: "image/png", bytes: RED_PNG },
      {
        partPath: "ppt/media/image1.jpeg",
        contentType: "image/jpeg",
        bytes: new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 1, 2, 3]),
      },
    ]);
    expect(reread.slides[0]?.shapes).toMatchObject([
      {
        kind: "image",
        name: "Product PNG",
        blipRelationshipId: "rId2",
        crop: { left: 1000, top: 2000, right: 3000, bottom: 4000 },
      },
      {
        kind: "image",
        name: "Product JPEG",
        blipRelationshipId: "rId3",
      },
    ]);
  });

  it("writes and rereads shape and picture shadow effects", () => {
    let source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(300),
      height: asEmu(400),
      name: "Outer Shadow Shape",
      effects: {
        glow: {
          radius: asEmu(12700),
          color: { kind: "srgb", hex: "FFCC00" },
        },
        outerShadow: {
          blurRadius: asEmu(40000),
          distance: asEmu(20000),
          direction: asOoxmlAngle(5400000),
          color: {
            kind: "srgb",
            hex: "112233",
            transforms: [{ kind: "alpha", value: asOoxmlPercent(40000) }],
          },
          alignment: "ctr",
          rotateWithShape: false,
        },
      },
    });
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "ellipse" },
      offsetX: asEmu(500),
      offsetY: asEmu(600),
      width: asEmu(700),
      height: asEmu(800),
      name: "Inner Shadow Shape",
      effects: {
        innerShadow: {
          blurRadius: asEmu(50000),
          distance: asEmu(30000),
          direction: asOoxmlAngle(10800000),
          color: { kind: "srgb", hex: "445566" },
        },
      },
    });
    source = addPicture(source, slideHandle, {
      bytes: RED_PNG,
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(1100),
      height: asEmu(1200),
      name: "Shadow Picture",
      effects: {
        innerShadow: {
          blurRadius: asEmu(60000),
          distance: asEmu(35000),
          direction: asOoxmlAngle(9000000),
          color: { kind: "srgb", hex: "778899" },
        },
        outerShadow: {
          blurRadius: asEmu(70000),
          distance: asEmu(45000),
          direction: asOoxmlAngle(13500000),
          color: {
            kind: "srgb",
            hex: "000000",
            transforms: [{ kind: "alpha", value: asOoxmlPercent(25000) }],
          },
          alignment: "br",
          rotateWithShape: true,
        },
      },
    });

    expect(findShapeByName(source, "Outer Shadow Shape").effects).toMatchObject({
      glow: { radius: 12700 },
      outerShadow: {
        blurRadius: 40000,
        distance: 20000,
        direction: 5400000,
        alignment: "ctr",
        rotateWithShape: false,
        color: { kind: "srgb", hex: "112233", transforms: [{ kind: "alpha", value: 40000 }] },
      },
    });
    expect(findShapeByName(source, "Inner Shadow Shape").effects?.innerShadow).toMatchObject({
      blurRadius: 50000,
      distance: 30000,
      direction: 10800000,
      color: { kind: "srgb", hex: "445566" },
    });
    expect(findImageByName(source, "Shadow Picture")).toMatchObject({
      effects: {
        innerShadow: { blurRadius: 60000, distance: 35000, direction: 9000000 },
        outerShadow: {
          blurRadius: 70000,
          distance: 45000,
          direction: 13500000,
          alignment: "br",
          rotateWithShape: true,
        },
      },
    });

    const output = writePptx(source);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    expect(reread.diagnostics).toEqual([]);
    expect(findImageByName(reread, "Shadow Picture")).toMatchObject(
      findImageByName(source, "Shadow Picture"),
    );
    expect(slideXml).toContain(
      `<a:effectLst><a:glow rad="12700"><a:srgbClr val="FFCC00"/></a:glow><a:outerShdw blurRad="40000" dist="20000" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="112233"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>`,
    );
    expect(slideXml).toContain(
      `<a:innerShdw blurRad="50000" dist="30000" dir="10800000"><a:srgbClr val="445566"/></a:innerShdw>`,
    );
    expect(slideXml).toContain(
      `<a:effectLst><a:innerShdw blurRad="60000" dist="35000" dir="9000000"><a:srgbClr val="778899"/></a:innerShdw><a:outerShdw blurRad="70000" dist="45000" dir="13500000" algn="br" rotWithShape="1"><a:srgbClr val="000000"><a:alpha val="25000"/></a:srgbClr></a:outerShdw></a:effectLst>`,
    );
  });

  it("keeps added pictures at the serialized shape-tree end and adds missing relationship namespace", () => {
    const source = readPptx(buildShapeDeleteFixture());
    const edited = addPicture(source, source.slides[0].handle!, {
      bytes: RED_PNG,
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(1100),
      height: asEmu(1200),
      name: "Appended Picture",
    });
    const slideXml = decoder.decode(getEntry(writePptx(edited), "ppt/slides/slide1.xml"));

    expect(slideXml).toContain(
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`,
    );
    expect(slideXml.indexOf(`name="Appended Picture"`)).toBeGreaterThan(
      slideXml.indexOf(`name="Keep Shape"`),
    );
  });

  it("writes formatted text box rPr, pPr, bodyPr, and xfrm from public APIs", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    const edited = addTextBox(source, slideHandle, {
      offsetX: asEmu(914400),
      offsetY: asEmu(457200),
      width: asEmu(4572000),
      height: asEmu(1828800),
      rotation: asOoxmlAngle(5400000),
      name: "Formatted TextBox",
      body: {
        anchor: "middle",
        marginLeft: asEmu(91440),
        marginRight: asEmu(182880),
        marginTop: asEmu(45720),
        marginBottom: asEmu(68580),
        autoFit: "shape",
      },
      paragraphs: [
        {
          properties: {
            align: "center",
            marginLeft: asEmu(342900),
            indent: asEmu(-285750),
            lineSpacing: { type: "points", value: asHundredthPt(1800) },
            bullet: {
              type: "character",
              character: "•",
              fontFace: "Aptos",
              size: asOoxmlPercent(125000),
            },
          },
          runs: [
            {
              text: "Solid styled",
              properties: {
                fontFace: "Aptos",
                fontSize: asPt(28),
                color: { kind: "srgb", hex: "112233" },
                bold: true,
                italic: true,
                underline: { style: "dbl", color: { kind: "srgb", hex: "445566" } },
                strike: true,
                highlight: { kind: "srgb", hex: "ffff00" },
                glow: { radius: asEmu(25400), color: { kind: "srgb", hex: "00aaff" } },
                outline: { width: asEmu(12700), color: { kind: "srgb", hex: "aa00aa" } },
                charSpacing: 120,
              },
            },
            {
              text: " gradient",
              properties: {
                gradientFill: {
                  gradientType: "linear",
                  angle: asOoxmlAngle(2700000),
                  stops: [
                    { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "ff0000" } },
                    {
                      position: asOoxmlPercent(100000),
                      color: { kind: "srgb", hex: "0000ff" },
                    },
                  ],
                },
                baseline: "superscript",
              },
            },
          ],
        },
        {
          properties: {
            align: "right",
            lineSpacing: { type: "percent", value: asOoxmlPercent(90000) },
            bullet: {
              type: "auto-number",
              scheme: "alphaLcParenR",
              startAt: 3,
              fontFace: "Aptos",
              size: asOoxmlPercent(100000),
            },
          },
          runs: [
            {
              text: "Subscript line",
              properties: {
                baseline: { type: "percent", value: asOoxmlPercent(-12500) },
                underline: true,
              },
            },
          ],
        },
      ],
    });
    const withShape = addShape(edited, edited.slides[0].handle!, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(914400),
      offsetY: asEmu(2743200),
      width: asEmu(4572000),
      height: asEmu(914400),
      name: "Auto-fit Shape",
      body: { autoFit: "shape" },
      paragraphs: [
        {
          properties: { bullet: { type: "none" } },
          runs: [{ text: "Shape text" }],
        },
      ],
    });
    const output = writePptx(withShape);
    const reread = readPptx(output);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const added = requireShape(findShapeByName(reread, "Formatted TextBox"));
    const addedShape = requireShape(findShapeByName(reread, "Auto-fit Shape"));

    expect(reread.diagnostics).toEqual([]);
    expect(added.transform).toMatchObject({ rotation: 5400000 });
    expect(added.textBody?.properties).toMatchObject({
      anchor: "middle",
      marginLeft: 91440,
      marginRight: 182880,
      marginTop: 45720,
      marginBottom: 68580,
      autoFit: "spAutofit",
    });
    expect(added.textBody?.paragraphs).toHaveLength(2);
    expect(added.textBody?.paragraphs[0]?.properties).toMatchObject({
      align: "center",
      lineSpacing: { type: "pts", value: 1800 },
      marginLeft: 342900,
      indent: -285750,
      bullet: { type: "char", char: "•" },
      bulletFont: "Aptos",
      bulletSizePct: 125000,
    });
    expect(added.textBody?.paragraphs[0]?.runs.map((run) => run.text)).toEqual([
      "Solid styled",
      " gradient",
    ]);
    expect(added.textBody?.paragraphs[1]?.properties).toMatchObject({
      align: "right",
      lineSpacing: { type: "pct", value: 90000 },
      bullet: { type: "autoNum", scheme: "alphaLcParenR", startAt: 3 },
      bulletFont: "Aptos",
      bulletSizePct: 100000,
    });
    expect(added.textBody?.paragraphs[1]?.runs[0]?.properties?.baseline).toBe(-12.5);
    expect(addedShape.textBody?.properties?.autoFit).toBe("spAutofit");
    expect(addedShape.textBody?.paragraphs[0]?.properties?.bullet).toEqual({ type: "none" });
    expect(slideXml).toContain(`<a:xfrm rot="5400000">`);
    expect(slideXml).toContain(
      `<a:bodyPr wrap="square" anchor="ctr" lIns="91440" rIns="182880" tIns="45720" bIns="68580"><a:spAutoFit/></a:bodyPr>`,
    );
    expect(slideXml).toContain(
      `<a:pPr algn="ctr" marL="342900" indent="-285750"><a:lnSpc><a:spcPts val="1800"/></a:lnSpc><a:buSzPct val="125000"/><a:buFont typeface="Aptos"/><a:buChar char="•"/>`,
    );
    expect(slideXml).toContain(
      `<a:pPr algn="r"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buSzPct val="100000"/><a:buFont typeface="Aptos"/><a:buAutoNum type="alphaLcParenR" startAt="3"/>`,
    );
    expect(slideXml).toContain(`<a:pPr><a:buNone/></a:pPr>`);
    expect(slideXml).toContain(`b="1"`);
    expect(slideXml).toContain(`i="1"`);
    expect(slideXml).toContain(`u="dbl"`);
    expect(slideXml).toContain(`strike="sngStrike"`);
    expect(slideXml).toContain(`sz="2800"`);
    expect(slideXml).toContain(`spc="120"`);
    expect(slideXml).toContain(`<a:solidFill><a:srgbClr val="112233"/></a:solidFill>`);
    expect(slideXml).toContain(
      `<a:uFill><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:uFill>`,
    );
    expect(slideXml).toContain(`<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>`);
    expect(slideXml).toContain(
      `<a:ln w="12700"><a:solidFill><a:srgbClr val="AA00AA"/></a:solidFill></a:ln>`,
    );
    expect(slideXml).toContain(
      `<a:effectLst><a:glow rad="25400"><a:srgbClr val="00AAFF"/></a:glow></a:effectLst>`,
    );
    expect(slideXml).toContain(`<a:latin typeface="Aptos"/>`);
    expect(slideXml).toContain(`<a:ea typeface="Aptos"/>`);
    expect(slideXml).toContain(`<a:cs typeface="Aptos"/>`);
    expect(slideXml).toContain(`<a:gradFill><a:gsLst>`);
    expect(slideXml).toContain(`<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>`);
    expect(slideXml).toContain(`<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>`);
    expect(slideXml).toContain(`<a:lin ang="2700000" scaled="1"/>`);
    expect(slideXml).toContain(`baseline="30000"`);
    expect(slideXml).toContain(`baseline="-12500"`);
  });

  it("writes alpha colors and linear or radial gradients consistently with the source model", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    const withText = addTextBox(source, slideHandle, {
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(300),
      height: asEmu(400),
      name: "Alpha Text",
      paragraphs: [
        {
          runs: [
            {
              text: "alpha",
              properties: {
                color: {
                  kind: "srgb",
                  hex: "112233",
                  transforms: [{ kind: "alpha", value: asOoxmlPercent(0) }],
                },
                glow: {
                  radius: asEmu(12700),
                  color: {
                    kind: "srgb",
                    hex: "445566",
                    transforms: [{ kind: "alpha", value: asOoxmlPercent(100000) }],
                  },
                },
              },
            },
            {
              text: " gradient",
              properties: {
                gradientFill: {
                  gradientType: "linear",
                  angle: asOoxmlAngle(1800000),
                  stops: [
                    {
                      position: asOoxmlPercent(0),
                      color: {
                        kind: "srgb",
                        hex: "FF0000",
                        transforms: [{ kind: "alpha", value: asOoxmlPercent(25000) }],
                      },
                    },
                    {
                      position: asOoxmlPercent(100000),
                      color: { kind: "srgb", hex: "0000FF" },
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    });
    const withSolidAndLinearOutline = addShape(withText, withText.slides[0].handle!, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(500),
      offsetY: asEmu(600),
      width: asEmu(700),
      height: asEmu(800),
      name: "Alpha Solid",
      fill: {
        kind: "solid",
        color: {
          kind: "srgb",
          hex: "778899",
          transforms: [{ kind: "alpha", value: asOoxmlPercent(50000) }],
        },
      },
      outline: {
        fill: {
          kind: "gradient",
          gradientType: "linear",
          stops: [
            {
              position: asOoxmlPercent(0),
              color: {
                kind: "srgb",
                hex: "00AA00",
                transforms: [{ kind: "alpha", value: asOoxmlPercent(40000) }],
              },
            },
            {
              position: asOoxmlPercent(100000),
              color: { kind: "srgb", hex: "AA0000" },
            },
          ],
        },
      },
      effects: {
        glow: {
          radius: asEmu(25400),
          color: {
            kind: "srgb",
            hex: "ABCDEF",
            transforms: [{ kind: "alpha", value: asOoxmlPercent(60000) }],
          },
        },
      },
    });
    const edited = addShape(
      withSolidAndLinearOutline,
      withSolidAndLinearOutline.slides[0].handle!,
      {
        geometry: { kind: "preset", preset: "ellipse" },
        offsetX: asEmu(900),
        offsetY: asEmu(1000),
        width: asEmu(1100),
        height: asEmu(1200),
        name: "Radial Shape",
        fill: {
          kind: "gradient",
          gradientType: "radial",
          centerX: asOoxmlPercent(25000),
          centerY: asOoxmlPercent(75000),
          stops: [
            { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FFFFFF" } },
            {
              position: asOoxmlPercent(100000),
              color: {
                kind: "srgb",
                hex: "000000",
                transforms: [{ kind: "alpha", value: asOoxmlPercent(75000) }],
              },
            },
          ],
        },
        outline: {
          fill: {
            kind: "gradient",
            gradientType: "radial",
            centerX: asOoxmlPercent(50000),
            centerY: asOoxmlPercent(50000),
            stops: [
              { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "FFFF00" } },
              {
                position: asOoxmlPercent(100000),
                color: {
                  kind: "srgb",
                  hex: "00FFFF",
                  transforms: [{ kind: "alpha", value: asOoxmlPercent(30000) }],
                },
              },
            ],
          },
        },
      },
    );

    const authoredText = requireShape(findShapeByName(edited, "Alpha Text"));
    const authoredSolid = requireShape(findShapeByName(edited, "Alpha Solid"));
    const authoredRadial = requireShape(findShapeByName(edited, "Radial Shape"));
    const output = writePptx(edited);
    const reread = readPptx(output);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(reread.diagnostics).toEqual([]);
    expect(authoredText.textBody?.paragraphs[0]?.runs[0]?.properties?.color).toEqual({
      kind: "srgb",
      hex: "112233",
      transforms: [{ kind: "alpha", value: 0 }],
    });
    expect(authoredSolid.fill).toEqual({
      kind: "solid",
      color: {
        kind: "srgb",
        hex: "778899",
        transforms: [{ kind: "alpha", value: 50000 }],
      },
    });
    expect(authoredSolid.effects?.glow?.color).toEqual({
      kind: "srgb",
      hex: "ABCDEF",
      transforms: [{ kind: "alpha", value: 60000 }],
    });
    expect(authoredRadial.fill).toMatchObject({
      kind: "gradient",
      gradientType: "radial",
      centerX: 0.25,
      centerY: 0.75,
    });
    expect(
      requireShape(findShapeByName(reread, "Alpha Text")).textBody?.paragraphs[0]?.runs[0]
        ?.properties?.color,
    ).toEqual(authoredText.textBody?.paragraphs[0]?.runs[0]?.properties?.color);
    expect(requireShape(findShapeByName(reread, "Alpha Solid")).fill).toEqual(authoredSolid.fill);
    expect(requireShape(findShapeByName(reread, "Alpha Solid")).outline).toEqual(
      authoredSolid.outline,
    );
    expect(requireShape(findShapeByName(reread, "Radial Shape")).fill).toEqual(authoredRadial.fill);
    expect(requireShape(findShapeByName(reread, "Radial Shape")).outline).toEqual(
      authoredRadial.outline,
    );
    expect(slideXml).toContain(
      `<a:solidFill><a:srgbClr val="112233"><a:alpha val="0"/></a:srgbClr></a:solidFill>`,
    );
    expect(slideXml).toContain(
      `<a:glow rad="12700"><a:srgbClr val="445566"><a:alpha val="100000"/></a:srgbClr></a:glow>`,
    );
    expect(slideXml).toContain(
      `<a:gs pos="0"><a:srgbClr val="FF0000"><a:alpha val="25000"/></a:srgbClr></a:gs>`,
    );
    expect(slideXml).toContain(
      `<a:solidFill><a:srgbClr val="778899"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`,
    );
    expect(slideXml).toContain(
      `<a:glow rad="25400"><a:srgbClr val="ABCDEF"><a:alpha val="60000"/></a:srgbClr></a:glow>`,
    );
    expect(slideXml).toContain(
      `<a:path path="circle"><a:fillToRect l="25000" t="75000" r="75000" b="25000"/></a:path>`,
    );
    expect(slideXml).toContain(
      `<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>`,
    );
    expect(slideXml.match(/<a:alpha val=/g)).toHaveLength(8);
  });

  it("writes preset geometry shapes with fill, line, glow, rotation, and text", () => {
    const source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");

    const withSolid = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "rect" },
      offsetX: asEmu(1000),
      offsetY: asEmu(2000),
      width: asEmu(3000),
      height: asEmu(4000),
      fill: { kind: "solid", color: { kind: "srgb", hex: "112233" } },
      name: "Solid Rect",
    });
    const edited = addShape(withSolid, withSolid.slides[0].handle!, {
      geometry: { kind: "preset", preset: "roundRect" },
      offsetX: asEmu(914400),
      offsetY: asEmu(457200),
      width: asEmu(2743200),
      height: asEmu(914400),
      rotation: asOoxmlAngle(5400000),
      fill: {
        kind: "gradient",
        gradientType: "linear",
        angle: asOoxmlAngle(2700000),
        stops: [
          { position: asOoxmlPercent(0), color: { kind: "srgb", hex: "ff0000" } },
          { position: asOoxmlPercent(100000), color: { kind: "srgb", hex: "0000ff" } },
        ],
      },
      outline: {
        width: asEmu(12700),
        fill: { kind: "solid", color: { kind: "srgb", hex: "00aa44" } },
        dash: "dash",
        headEnd: { type: "oval", width: "sm", length: "sm" },
        tailEnd: { type: "triangle", width: "med", length: "lg" },
      },
      effects: {
        glow: { radius: asEmu(25400), color: { kind: "srgb", hex: "aa00aa" } },
      },
      body: { anchor: "middle" },
      paragraphs: [
        {
          properties: { align: "center" },
          runs: [{ text: "Shape label", properties: { bold: true } }],
        },
      ],
      name: "Styled Shape",
    });
    const lineShape = addShape(edited, edited.slides[0].handle!, {
      geometry: { kind: "preset", preset: "line" },
      offsetX: asEmu(1),
      offsetY: asEmu(2),
      width: asEmu(3),
      height: asEmu(4),
      outline: {
        tailEnd: { type: "triangle", width: "med", length: "med" },
      },
      name: "Line Shape",
    });
    const ellipseShape = addShape(lineShape, lineShape.slides[0].handle!, {
      geometry: { kind: "preset", preset: "ellipse" },
      offsetX: asEmu(5),
      offsetY: asEmu(6),
      width: asEmu(7),
      height: asEmu(8),
      name: "Ellipse Shape",
    });
    const output = writePptx(ellipseShape);
    const reread = readPptx(output);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const solid = requireShape(findShapeByName(reread, "Solid Rect"));
    const styled = requireShape(findShapeByName(reread, "Styled Shape"));

    expect(reread.diagnostics).toEqual([]);
    expect(solid).toMatchObject({
      geometry: { preset: "rect" },
      fill: { kind: "solid", color: { kind: "srgb", hex: "112233" } },
    });
    expect(styled).toMatchObject({
      geometry: { preset: "roundRect" },
      transform: { rotation: 5400000 },
      fill: {
        kind: "gradient",
        gradientType: "linear",
        angle: 2700000,
      },
      outline: {
        width: 12700,
        fill: { kind: "solid", color: { kind: "srgb", hex: "00AA44" } },
        dashStyle: "dash",
        headEnd: { type: "oval", width: "sm", length: "sm" },
        tailEnd: { type: "triangle", width: "med", length: "lg" },
      },
      effects: {
        glow: { radius: 25400, color: { kind: "srgb", hex: "AA00AA" } },
      },
    });
    expect(styled.textBody?.paragraphs[0]?.runs[0]?.text).toBe("Shape label");
    expect(findShapeByName(reread, "Line Shape").geometry).toEqual({ preset: "line" });
    expect(findShapeByName(reread, "Ellipse Shape").geometry).toEqual({ preset: "ellipse" });
    expect(slideXml).toContain(`<a:prstGeom prst="rect"`);
    expect(slideXml).toContain(`<a:prstGeom prst="roundRect"`);
    expect(slideXml).toContain(`<a:prstGeom prst="line"`);
    expect(slideXml).toContain(`<a:prstGeom prst="ellipse"`);
    expect(slideXml).toContain(`<a:solidFill><a:srgbClr val="112233"/></a:solidFill>`);
    expect(slideXml).toContain(`<a:gradFill><a:gsLst>`);
    expect(slideXml).toContain(`<a:lin ang="2700000" scaled="1"/>`);
    expect(slideXml).toContain(
      `<a:ln w="12700"><a:solidFill><a:srgbClr val="00AA44"/></a:solidFill><a:prstDash val="dash"/>`,
    );
    expect(slideXml).toContain(`<a:headEnd type="oval" w="sm" len="sm"`);
    expect(slideXml).toContain(`<a:tailEnd type="triangle" w="med" len="lg"`);
    expect(slideXml).toContain(
      `<a:effectLst><a:glow rad="25400"><a:srgbClr val="AA00AA"/></a:glow></a:effectLst>`,
    );
    expect(slideXml).toContain(`<p:txBody>`);
    expect(slideXml).toContain(`<a:t>Shape label</a:t>`);
    expect(slideXml).toContain(`<a:xfrm rot="5400000">`);
  });

  it("writes and rereads adjusted, custom, flipped, and zero-extent shape geometry", () => {
    let source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "roundRect", adjustValues: { adj: 25000 } },
      offsetX: asEmu(100),
      offsetY: asEmu(200),
      width: asEmu(300),
      height: asEmu(400),
      flipHorizontal: true,
      name: "Adjusted Geometry",
    });
    source = addShape(source, slideHandle, {
      geometry: {
        kind: "custom",
        paths: [
          {
            width: 100,
            height: 100,
            commands: [
              { kind: "moveTo", x: 0, y: 100 },
              { kind: "lineTo", x: 50, y: 0 },
              { kind: "lineTo", x: 100, y: 100 },
              { kind: "close" },
            ],
          },
        ],
      },
      offsetX: asEmu(500),
      offsetY: asEmu(600),
      width: asEmu(700),
      height: asEmu(800),
      flipVertical: true,
      name: "Custom Geometry",
    });
    source = addShape(source, slideHandle, {
      geometry: { kind: "preset", preset: "line" },
      offsetX: asEmu(900),
      offsetY: asEmu(1000),
      width: asEmu(0),
      height: asEmu(1200),
      flipHorizontal: true,
      flipVertical: true,
      name: "Zero Extent Line",
    });

    const output = writePptx(source);
    const reread = readPptx(output);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(requireShape(findShapeByName(reread, "Adjusted Geometry"))).toMatchObject({
      geometry: { preset: "roundRect", adjustValues: { adj: 25000 } },
      transform: { flipHorizontal: true },
    });
    expect(requireShape(findShapeByName(reread, "Custom Geometry"))).toMatchObject({
      geometry: {
        kind: "custom",
        paths: [{ width: 100, height: 100, commands: "M 0 100 L 50 0 L 100 100 Z" }],
      },
      transform: { flipVertical: true },
    });
    expect(requireShape(findShapeByName(reread, "Zero Extent Line"))).toMatchObject({
      geometry: { preset: "line" },
      transform: { width: 0, height: 1200, flipHorizontal: true, flipVertical: true },
    });
    expect(slideXml).toContain(`<a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst>`);
    expect(slideXml).toContain(`<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>`);
    expect(slideXml).toContain(`<a:moveTo><a:pt x="0" y="100"/></a:moveTo>`);
    expect(slideXml).toContain(`<a:lnTo><a:pt x="50" y="0"/></a:lnTo>`);
    expect(slideXml).toContain(`<a:close/>`);
    expect(slideXml).toContain(`<a:xfrm flipH="1" flipV="1">`);
    expect(slideXml).toContain(`<a:ext cx="0" cy="1200"/>`);
  });

  it("writes custom slide size without fixed 16:9 metadata", () => {
    const source = createPptx({
      slideSize: { width: asEmu(7315200), height: asEmu(5486400) },
    });
    const output = writePptx(source);
    const reread = readPptx(output);

    expect(reread.presentation.slideSize).toEqual({
      width: asEmu(7315200),
      height: asEmu(5486400),
    });
    expect(decoder.decode(getEntry(output, "ppt/presentation.xml"))).toContain(
      `<p:sldSz cx="7315200" cy="5486400"/>`,
    );
    expect(decoder.decode(getEntry(output, "docProps/app.xml"))).not.toContain(
      "On-screen Show (16:9)",
    );
  });

  it("rejects invalid custom slide sizes", () => {
    expect(() =>
      createPptx({
        slideSize: { width: asEmu(Number.NaN), height: asEmu(5486400) },
      }),
    ).toThrow(/slideSize\.width/);
    expect(() =>
      createPptx({
        slideSize: { width: asEmu(7315200), height: asEmu(0) },
      }),
    ).toThrow(/slideSize\.height/);
  });
});
