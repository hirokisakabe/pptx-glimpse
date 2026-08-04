import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  addChart,
  asEmu,
  createComputedView,
  createPptx,
  readPptx,
  updateBubbleChartData,
  updateChartData,
  updateScatterChartData,
  writePptx,
} from "../index.js";

const decoder = new TextDecoder();

describe("updateChartData", () => {
  it("updates chart formulas, caches, workbook data, and computed renderer input", () => {
    const input = buildExistingChart();
    const source = readPptx(input);
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("chart fixture should have a handle");

    const edited = updateChartData(source, chart.handle, {
      series: [
        {
          name: "Edited revenue",
          categories: ["Apr", "May", "Jun"],
          values: [40, 55, 70],
        },
        {
          name: "Edited cost",
          categories: ["Apr", "May", "Jun"],
          values: [25, 30, 42],
        },
      ],
    });

    expect(edited).not.toBe(source);
    expect(source.edits).toBeUndefined();
    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "updateChartData",
      handle: chart.handle,
      chartPartPath: "ppt/charts/chart1.xml",
      workbookPartPath: "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
    });
    const computedChart = createComputedView(edited).slides[0]?.elements.find(
      (element) => element.kind === "chart",
    );
    expect(computedChart?.kind).toBe("chart");
    if (computedChart?.kind !== "chart") throw new Error("computed chart should exist");
    expect(computedChart.chartData).toMatchObject({
      chartType: "bar",
      title: "Preserved title",
      categories: ["Apr", "May", "Jun"],
      series: [
        { name: "Edited revenue", values: [40, 55, 70] },
        { name: "Edited cost", values: [25, 30, 42] },
      ],
    });

    const before = unzipSync(input);
    const beforeWorkbook = unzipSync(before["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]);
    const output = unzipSync(writePptx(edited));
    const chartXml = decoder.decode(output["ppt/charts/chart1.xml"]);
    expect(chartXml).toContain("<c:barChart>");
    expect(chartXml).toContain("<c:title>");
    expect(chartXml).toContain("<c:legend>");
    expect(chartXml).toContain('<a:srgbClr val="4472C4"/>');
    expect(chartXml).toContain('<c:ext uri="preserve-me">');
    expect(chartXml).toContain("Sheet1!$A$2:$A$4");
    expect(chartXml).toContain("Sheet1!$B$2:$B$4");
    expect(chartXml).toContain("Sheet1!$C$2:$C$4");
    expect(chartXml.match(/<c:ptCount val="3"\/>/g)).toHaveLength(4);
    expect(output["ppt/slides/slide1.xml"]).toEqual(before["ppt/slides/slide1.xml"]);
    expect(output["ppt/theme/theme1.xml"]).toEqual(before["ppt/theme/theme1.xml"]);

    const workbook = unzipSync(output["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]);
    const worksheet = decoder.decode(workbook["xl/worksheets/sheet1.xml"]);
    expect(worksheet).toContain('dimension ref="A1:C4"');
    expect(worksheet).toContain("<t>Edited revenue</t>");
    expect(worksheet).toContain('<c r="A4" t="inlineStr"><is><t>Jun</t></is></c>');
    expect(worksheet).toContain('<c r="B4"><v>70</v></c>');
    expect(worksheet).toContain('<c r="C4"><v>42</v></c>');
    expect(workbook["xl/styles.xml"]).toEqual(beforeWorkbook["xl/styles.xml"]);

    const rereadChart = createComputedView(readPptx(writePptx(edited))).slides[0]?.elements.find(
      (element) => element.kind === "chart",
    );
    expect(rereadChart?.kind === "chart" ? rereadChart.chartData?.categories : undefined).toEqual([
      "Apr",
      "May",
      "Jun",
    ]);
  });

  it("updates a fixed-topology bar and line category combo by typed source identity", () => {
    const input = buildExistingComboChart();
    const source = readPptx(input);
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("combo fixture should have a handle");
    const before = unzipSync(input);
    const beforeWorkbook = unzipSync(before["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]);

    const edited = updateChartData(source, chart.handle, {
      series: [
        {
          source: { chartType: "line", index: 7 },
          name: "Edited trend",
          categories: ["Apr", "May", "Jun"],
          values: [35, 48, 63],
        },
        {
          source: { chartType: "bar", index: 0 },
          name: "Edited revenue",
          categories: ["Apr", "May", "Jun"],
          values: [40, 55, 70],
        },
      ],
    });

    const computed = createComputedView(edited).slides[0]?.elements.find(
      (element) => element.kind === "chart",
    );
    expect(computed?.kind === "chart" ? computed.chartData : undefined).toMatchObject({
      chartType: "combo",
      categories: ["Apr", "May", "Jun"],
      series: [
        {
          source: { chartType: "bar", index: 0 },
          name: "Edited revenue",
          values: [40, 55, 70],
        },
        {
          source: { chartType: "line", index: 7 },
          name: "Edited trend",
          values: [35, 48, 63],
        },
      ],
      plotGroups: [
        { chartType: "bar", valueAxisId: "100003" },
        { chartType: "line", valueAxisId: "200003" },
      ],
      valueAxes: [
        { id: "100003", position: "l" },
        { id: "200003", position: "r" },
      ],
    });

    const output = unzipSync(writePptx(edited));
    const chartXml = decoder.decode(output["ppt/charts/chart1.xml"]);
    expect(chartXml.indexOf("<c:barChart>")).toBeLessThan(chartXml.indexOf("<c:lineChart>"));
    expect(chartXml.match(/<c:ser>/g)).toHaveLength(2);
    expect(chartXml).toContain('<c:idx val="0"/><c:order val="0"/>');
    expect(chartXml).toContain('<c:idx val="7"/><c:order val="9"/>');
    expect(chartXml).toContain("Sheet1!$B$1");
    expect(chartXml).toContain("Sheet1!$C$1");
    expect(chartXml).toContain("Sheet1!$A$2:$A$4");
    expect(chartXml).toContain("Sheet1!$B$2:$B$4");
    expect(chartXml).toContain("Sheet1!$C$2:$C$4");
    expect(chartXml.match(/<c:ptCount val="3"\/>/g)).toHaveLength(4);
    expect(chartXml).toContain('<c:axId val="100002"/><c:axId val="100003"/>');
    expect(chartXml).toContain('<c:axId val="100002"/><c:axId val="200003"/>');
    expect(chartXml).toContain('<c:crossAx val="100002"/>');
    expect(chartXml).toContain('<c:crossAx val="100003"/>');
    expect(chartXml.match(/<c:catAx>/g)).toHaveLength(1);
    expect(chartXml.match(/<c:valAx>/g)).toHaveLength(2);
    expect(chartXml).toContain('<c:axPos val="r"/>');
    expect(chartXml).toContain('<c:crosses val="max"/>');
    expect(chartXml).toContain("Preserved title");
    expect(chartXml).toContain("<c:legend>");
    expect(chartXml).toContain('uri="combo-group-preserve"');
    expect(chartXml).toContain('uri="preserve-me"');
    expect(chartXml).toContain('val="ED7D31"');
    expect(output["ppt/slides/slide1.xml"]).toEqual(before["ppt/slides/slide1.xml"]);
    expect(output["ppt/theme/theme1.xml"]).toEqual(before["ppt/theme/theme1.xml"]);

    const workbook = unzipSync(output["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]);
    const worksheet = decoder.decode(workbook["xl/worksheets/sheet1.xml"]);
    expect(worksheet).toContain('dimension ref="A1:C4"');
    expect(worksheet).toContain("<t>Edited revenue</t>");
    expect(worksheet).toContain("<t>Edited trend</t>");
    expect(worksheet).toContain('<c r="B4"><v>70</v></c>');
    expect(worksheet).toContain('<c r="C4"><v>63</v></c>');
    expect(workbook["xl/styles.xml"]).toEqual(beforeWorkbook["xl/styles.xml"]);
    expect(workbook["xl/combo-preserve.xml"]).toEqual(beforeWorkbook["xl/combo-preserve.xml"]);

    const reread = createComputedView(readPptx(writePptx(edited))).slides[0]?.elements.find(
      (element) => element.kind === "chart",
    );
    expect(reread?.kind === "chart" ? reread.chartData?.series : undefined).toMatchObject([
      { name: "Edited revenue", source: { chartType: "bar", index: 0 } },
      { name: "Edited trend", source: { chartType: "line", index: 7 } },
    ]);
  });

  it("rejects combo topology and identity changes atomically", () => {
    const source = readPptx(buildExistingComboChart());
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("combo fixture should have a handle");
    const validSeries = [
      {
        source: { chartType: "bar" as const, index: 0 },
        name: "Bar",
        categories: ["A"],
        values: [1],
      },
      {
        source: { chartType: "line" as const, index: 7 },
        name: "Line",
        categories: ["A"],
        values: [2],
      },
    ];
    const failures = [
      { series: validSeries.slice(0, 1), message: "series count must remain unchanged" },
      {
        series: [validSeries[0], { ...validSeries[1], source: validSeries[0].source }],
        message: "identity is duplicated",
      },
      {
        series: [
          validSeries[0],
          { ...validSeries[1], source: { chartType: "line" as const, index: 8 } },
        ],
        message: "identity does not match",
      },
      {
        series: [validSeries[0], { ...validSeries[1], source: undefined }],
        message: "identity is missing or invalid",
      },
    ];
    for (const failure of failures) {
      expect(() => updateChartData(source, chart.handle, { series: failure.series })).toThrow(
        failure.message,
      );
      expect(source.edits).toBeUndefined();
    }

    for (const replacement of ["</c:lineChart><c:areaChart/>", "</c:lineChart><c:scatterChart/>"]) {
      const files = unzipSync(buildExistingComboChart());
      files["ppt/charts/chart1.xml"] = replaceText(
        files["ppt/charts/chart1.xml"],
        "</c:lineChart>",
        replacement,
      );
      const unsupported = readPptx(zipFixture(files));
      const unsupportedChart = unsupported.slides[0]?.shapes.find(
        (shape) => shape.kind === "chart",
      );
      if (unsupportedChart?.handle === undefined) throw new Error("chart handle is missing");
      expect(() =>
        updateChartData(unsupported, unsupportedChart.handle, { series: validSeries }),
      ).toThrow("chart type or combination is not supported");
      expect(unsupported.edits).toBeUndefined();
    }
  });

  it("rejects a horizontal bar and line combo atomically", () => {
    const files = unzipSync(buildExistingComboChart());
    files["ppt/charts/chart1.xml"] = replaceText(
      files["ppt/charts/chart1.xml"],
      '<c:barDir val="col"/>',
      '<c:barDir val="bar"/>',
    );
    const source = readPptx(zipFixture(files));
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("combo fixture should have a handle");

    expect(() =>
      updateChartData(source, chart.handle, {
        series: [
          {
            source: { chartType: "bar", index: 0 },
            name: "Bar",
            categories: ["A"],
            values: [1],
          },
          {
            source: { chartType: "line", index: 7 },
            name: "Line",
            categories: ["A"],
            values: [2],
          },
        ],
      }),
    ).toThrow("combo charts require barDir=col");
    expect(source.edits).toBeUndefined();
  });

  it("rejects stacked and percent-stacked combo grouping atomically", () => {
    for (const [search, replacement] of [
      ['<c:grouping val="clustered"/>', '<c:grouping val="stacked"/>'],
      ['<c:grouping val="standard"/>', '<c:grouping val="percentStacked"/>'],
    ]) {
      const files = unzipSync(buildExistingComboChart());
      files["ppt/charts/chart1.xml"] = replaceText(
        files["ppt/charts/chart1.xml"],
        search,
        replacement,
      );
      expectEditFailure(readPptx(zipFixture(files)), "combo chart grouping is not supported");
    }
  });

  it("rejects incomplete or ambiguous combo axis topology atomically", () => {
    const mutations = [
      (xml: string) => xml.replace(/<c:lineChart>(.*?)<c:ser>.*?<\/c:ser>/, "<c:lineChart>$1"),
      (xml: string) => xml.replace('<c:axId val="200003"/>', '<c:axId val="999999"/>'),
      (xml: string) =>
        xml.replace(
          '<c:axId val="100002"/><c:axId val="200003"/>',
          '<c:axId val="100003"/><c:axId val="200003"/>',
        ),
      (xml: string) =>
        xml
          .replace(
            '<c:axId val="100002"/><c:axId val="200003"/>',
            '<c:axId val="300002"/><c:axId val="200003"/>',
          )
          .replace("</c:plotArea>", '<c:catAx><c:axId val="300002"/></c:catAx></c:plotArea>'),
      (xml: string) => xml.replaceAll("c:catAx", "c:dateAx"),
      (xml: string) =>
        xml.replace(
          '<c:axId val="100002"/><c:axId val="200003"/>',
          '<c:axId val="100002"/><c:axId val="200003"/><c:axId val="999999"/>',
        ),
      (xml: string) =>
        xml.replace(
          '<c:axId val="100002"/><c:axId val="200003"/>',
          '<c:axId val="100002"/><c:axId/>',
        ),
    ];
    for (const mutate of mutations) {
      const files = unzipSync(buildExistingComboChart());
      files["ppt/charts/chart1.xml"] = new TextEncoder().encode(
        mutate(decoder.decode(files["ppt/charts/chart1.xml"])),
      );
      expectEditFailure(
        readPptx(zipFixture(files)),
        mutate === mutations[0]
          ? "combo plot groups must each contain series"
          : "combo chart axis topology is not supported",
      );
    }
  });

  it("adds series by cloning the last series formatting and synchronizes the workbook", () => {
    const files = unzipSync(buildExistingChart());
    files["ppt/charts/chart1.xml"] = addSeriesExtensionMarkers(files["ppt/charts/chart1.xml"]);
    files["ppt/charts/chart1.xml"] = replaceText(
      files["ppt/charts/chart1.xml"],
      '<c:idx val="1"/><c:order val="1"/>',
      '<c:idx val="7"/><c:order val="9"/>',
    );
    files["ppt/charts/chart1.xml"] = replaceText(
      files["ppt/charts/chart1.xml"],
      "<c:legend>",
      '<c:legend><c:legendEntry><c:idx val="7"/><c:delete val="1"/></c:legendEntry>',
    );
    files["ppt/charts/chart1.xml"] = replaceText(
      files["ppt/charts/chart1.xml"],
      "</c:chart>",
      '<c:extLst><c:ext uri="chart-guid"><c16:uniqueId xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart" val="{00000000-0000-0000-0000-000000000003}"/></c:ext></c:extLst></c:chart>',
    );
    const source = readPptx(zipFixture(files));
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("chart fixture should have a handle");

    const output = unzipSync(
      writePptx(
        updateChartData(source, chart.handle, {
          series: [
            { name: "Revenue", categories: ["Apr", "May"], values: [40, 55] },
            { name: "Cost", categories: ["Apr", "May"], values: [25, 30] },
            { name: "Profit", categories: ["Apr", "May"], values: [15, 25] },
          ],
        }),
      ),
    );
    const chartXml = decoder.decode(output["ppt/charts/chart1.xml"]);
    expect(chartXml.match(/<c:ser>/g)).toHaveLength(3);
    expect(chartXml).toContain('<c:idx val="0"/><c:order val="0"/>');
    expect(chartXml).toContain('<c:idx val="7"/><c:order val="9"/>');
    expect(chartXml).toContain('<c:idx val="8"/><c:order val="10"/>');
    expect(chartXml).toContain(
      '<c:legendEntry><c:idx val="7"/><c:delete val="1"/></c:legendEntry>',
    );
    expect(chartXml).toContain("Sheet1!$D$1");
    expect(chartXml).toContain("Sheet1!$D$2:$D$3");
    expect(chartXml.match(/uri="series-1"/g)).toHaveLength(1);
    expect(chartXml.match(/uri="series-2"/g)).toHaveLength(2);
    expect(chartXml.match(/val="\{00000000-0000-0000-0000-000000000001\}"/g)).toHaveLength(1);
    expect(chartXml.match(/val="\{00000000-0000-0000-0000-000000000002\}"/g)).toHaveLength(1);
    expect(chartXml.match(/val="\{00000000-0000-0000-0000-000000000003\}"/g)).toHaveLength(1);
    expect(chartXml.match(/val="\{00000000-0000-0000-0000-000000000004\}"/g)).toHaveLength(1);
    expect(chartXml.match(/<c16:uniqueId[^>]*val="2"/g)).toHaveLength(2);
    expect(chartXml.match(/val="ED7D31"/g)).toHaveLength(4);

    const workbook = unzipSync(output["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]);
    const worksheet = decoder.decode(workbook["xl/worksheets/sheet1.xml"]);
    expect(worksheet).toContain('dimension ref="A1:D3"');
    expect(worksheet).toContain('<c r="D1" t="inlineStr"><is><t>Profit</t></is></c>');
    expect(worksheet).toContain('<c r="D3"><v>25</v></c>');

    const rereadChart = createComputedView(readPptx(zipFixture(output))).slides[0]?.elements.find(
      (element) => element.kind === "chart",
    );
    expect(rereadChart?.kind === "chart" ? rereadChart.chartData?.series : undefined).toMatchObject(
      [
        { name: "Revenue", values: [40, 55] },
        { name: "Cost", values: [25, 30] },
        { name: "Profit", values: [15, 25] },
      ],
    );
  });

  it("removes trailing series with their XML while preserving retained series XML", () => {
    const files = unzipSync(buildExistingChart());
    files["ppt/charts/chart1.xml"] = addSeriesExtensionMarkers(files["ppt/charts/chart1.xml"]);
    const source = readPptx(zipFixture(files));
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("chart fixture should have a handle");

    const output = unzipSync(
      writePptx(
        updateChartData(source, chart.handle, {
          series: [{ name: "Only", categories: ["Apr", "May"], values: [40, 55] }],
        }),
      ),
    );
    const chartXml = decoder.decode(output["ppt/charts/chart1.xml"]);
    expect(chartXml.match(/<c:ser>/g)).toHaveLength(1);
    expect(chartXml).toContain('uri="series-1"');
    expect(chartXml).not.toContain('uri="series-2"');
    expect(chartXml).toContain('<c:idx val="0"/><c:order val="0"/>');
    expect(chartXml).not.toContain("Sheet1!$C$1");
    expect(chartXml).toContain('uri="preserve-me"');

    const workbook = unzipSync(output["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]);
    const worksheet = decoder.decode(workbook["xl/worksheets/sheet1.xml"]);
    expect(worksheet).toContain('dimension ref="A1:B3"');
    expect(worksheet).not.toContain('<c r="C1"');
  });

  it("updates every supported existing category chart type", () => {
    let source = createPptx();
    const slideHandle = source.slides[0]?.handle;
    if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
    const types = ["bar", "line", "pie", "area", "doughnut", "radar"] as const;
    for (const [index, chartType] of types.entries()) {
      source = addChart(source, slideHandle, {
        chartType,
        offsetX: asEmu(index * 1000),
        offsetY: asEmu(0),
        width: asEmu(1000),
        height: asEmu(1000),
        series: [{ name: "Original", categories: ["A", "B"], values: [1, 2] }],
      });
    }
    source = readPptx(writePptx(source));
    const charts = source.slides[0]?.shapes.filter((shape) => shape.kind === "chart") ?? [];
    for (const chart of charts) {
      if (chart.handle === undefined) throw new Error("chart fixture should have a handle");
      source = updateChartData(source, chart.handle, {
        series: [
          { name: "Edited", categories: ["X", "Y", "Z"], values: [3, 5, 8] },
          { name: "Added", categories: ["X", "Y", "Z"], values: [2, 4, 6] },
        ],
      });
    }

    const computedCharts = createComputedView(
      readPptx(writePptx(source)),
    ).slides[0]?.elements.filter((element) => element.kind === "chart");
    expect(computedCharts?.map((chart) => chart.chartData?.chartType)).toEqual(types);
    expect(computedCharts?.every((chart) => chart.chartData?.categories.length === 3)).toBe(true);
    expect(computedCharts?.every((chart) => chart.chartData?.series.length === 2)).toBe(true);
  });

  it("preserves and targets a non-default worksheet name", () => {
    const files = unzipSync(buildExistingChart());
    const chartPath = "ppt/charts/chart1.xml";
    files[chartPath] = replaceAllText(files[chartPath], "Sheet1!", "'Sales Data'!");
    const workbookPath = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
    const workbook = unzipSync(files[workbookPath]);
    workbook["xl/workbook.xml"] = replaceText(
      workbook["xl/workbook.xml"],
      'name="Sheet1"',
      'name="Sales Data"',
    );
    const worksheetBytes = workbook["xl/worksheets/sheet1.xml"];
    delete workbook["xl/worksheets/sheet1.xml"];
    workbook["xl/chart-data/data.xml"] = worksheetBytes;
    workbook["xl/_rels/workbook.xml.rels"] = replaceText(
      workbook["xl/_rels/workbook.xml.rels"],
      'Target="worksheets/sheet1.xml"',
      'Target="chart-data/data.xml"',
    );
    const workbookXmlBefore = workbook["xl/workbook.xml"];
    const workbookRelationshipsBefore = workbook["xl/_rels/workbook.xml.rels"];
    files[workbookPath] = zipFixture(workbook);
    const source = readPptx(zipFixture(files));
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("chart fixture should have a handle");

    const output = unzipSync(
      writePptx(
        updateChartData(source, chart.handle, {
          series: [
            { name: "One", categories: ["A"], values: [1] },
            { name: "Two", categories: ["A"], values: [2] },
          ],
        }),
      ),
    );
    expect(decoder.decode(output[chartPath])).toContain("&apos;Sales Data&apos;!$B$2:$B$2");
    const outputWorkbook = unzipSync(output[workbookPath]);
    expect(outputWorkbook["xl/workbook.xml"]).toEqual(workbookXmlBefore);
    expect(outputWorkbook["xl/_rels/workbook.xml.rels"]).toEqual(workbookRelationshipsBefore);
    expect(decoder.decode(outputWorkbook["xl/chart-data/data.xml"])).toContain("<t>One</t>");
  });

  it("preserves worksheet prefixes and row and cell formatting attributes", () => {
    const files = unzipSync(buildExistingChart());
    const workbookPath = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
    const workbook = unzipSync(files[workbookPath]);
    let worksheet = decoder.decode(workbook["xl/worksheets/sheet1.xml"]);
    for (const element of [
      "worksheet",
      "dimension",
      "sheetViews",
      "sheetView",
      "sheetFormatPr",
      "sheetData",
      "row",
      "c",
      "is",
      "t",
      "v",
    ]) {
      worksheet = worksheet
        .replaceAll(`<${element}`, `<x:${element}`)
        .replaceAll(`</${element}>`, `</x:${element}>`);
    }
    worksheet = worksheet.replace(
      'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
      'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
    );
    worksheet = worksheet
      .replace('<x:row r="1">', '<x:row r="1" ht="22" customHeight="1">')
      .replace('<x:c r="B1"', '<x:c r="B1" s="0"');
    workbook["xl/worksheets/sheet1.xml"] = new TextEncoder().encode(worksheet);
    files[workbookPath] = zipFixture(workbook);
    const source = readPptx(zipFixture(files));
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("chart fixture should have a handle");

    const output = unzipSync(
      writePptx(
        updateChartData(source, chart.handle, {
          series: [
            { name: "One", categories: ["A"], values: [1] },
            { name: "Two", categories: ["A"], values: [2] },
          ],
        }),
      ),
    );
    const outputWorksheet = decoder.decode(
      unzipSync(output[workbookPath])["xl/worksheets/sheet1.xml"],
    );
    expect(outputWorksheet).toContain("<x:sheetData><x:row");
    expect(outputWorksheet).toContain('ht="22"');
    expect(outputWorksheet).toContain('customHeight="1"');
    expect(outputWorksheet).toContain('<x:c r="B1" s="0" t="inlineStr">');
  });

  it("rejects missing workbooks, external workbook relationships, and unsupported layouts", () => {
    const input = buildExistingChart();
    const missingWorkbookFiles = unzipSync(input);
    delete missingWorkbookFiles["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"];
    const missingWorkbookSource = readPptx(zipFixture(missingWorkbookFiles));
    expectEditFailure(missingWorkbookSource, "embedded workbook part was not found");

    const externalFiles = unzipSync(input);
    externalFiles["ppt/charts/_rels/chart1.xml.rels"] = replaceText(
      externalFiles["ppt/charts/_rels/chart1.xml.rels"],
      'Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"',
      'Target="https://example.com/data.xlsx" TargetMode="External"',
    );
    const externalSource = readPptx(zipFixture(externalFiles));
    expectEditFailure(externalSource, "external workbook data is not supported");

    const unsupportedFiles = unzipSync(input);
    unsupportedFiles["ppt/charts/chart1.xml"] = replaceText(
      unsupportedFiles["ppt/charts/chart1.xml"],
      "Sheet1!$B$1",
      "Sheet1!$C$1",
    );
    const unsupportedSource = readPptx(zipFixture(unsupportedFiles));
    expectEditFailure(unsupportedSource, "unsupported data layout");

    const formulaFiles = unzipSync(input);
    const workbookFiles = unzipSync(formulaFiles["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]);
    workbookFiles["xl/worksheets/sheet1.xml"] = replaceText(
      workbookFiles["xl/worksheets/sheet1.xml"],
      "<v>10</v>",
      "<f>1+9</f><v>10</v>",
    );
    formulaFiles["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"] = zipFixture(workbookFiles);
    const formulaSource = readPptx(zipFixture(formulaFiles));
    expectEditFailure(formulaSource, "formulas in the chart data range are not supported");

    for (const [search, replacement, context] of [
      ['<row r="1">', '<row r="1"><extLst/>', "row"],
      [
        '<c r="A1" t="inlineStr"><is><t>Category</t></is></c>',
        '<c r="A1" t="inlineStr"><is><t>Category</t></is><extLst/></c>',
        "cell",
      ],
    ] as const) {
      expectEditFailure(
        readPptx(replaceEmbeddedWorksheetText(input, search, replacement)),
        `worksheet ${context} child XML is not supported`,
      );
    }

    const formattedOutsideRangeFiles = unzipSync(input);
    const formattedWorkbook = unzipSync(
      formattedOutsideRangeFiles["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"],
    );
    formattedWorkbook["xl/worksheets/sheet1.xml"] = replaceText(
      formattedWorkbook["xl/worksheets/sheet1.xml"],
      "</sheetData>",
      '<row r="10" hidden="1" customHeight="1"/></sheetData>',
    );
    formattedOutsideRangeFiles["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"] =
      zipFixture(formattedWorkbook);
    const formattedOutsideRangeSource = readPptx(zipFixture(formattedOutsideRangeFiles));
    expectEditFailure(formattedOutsideRangeSource, "unsupported data layout");

    const comboFiles = unzipSync(input);
    comboFiles["ppt/charts/chart1.xml"] = replaceText(
      comboFiles["ppt/charts/chart1.xml"],
      "</c:barChart>",
      "</c:barChart><c:scatterChart/>",
    );
    const comboSource = readPptx(zipFixture(comboFiles));
    expectEditFailure(comboSource, "chart type or combination is not supported");

    const sharedSource = readPptx(buildSharedWorkbookChart());
    expectEditFailure(sharedSource, "embedded workbook is shared by another package part");

    const duplicateIdentityFiles = unzipSync(input);
    duplicateIdentityFiles["ppt/charts/chart1.xml"] = replaceText(
      duplicateIdentityFiles["ppt/charts/chart1.xml"],
      '<c:idx val="1"/>',
      '<c:idx val="0"/>',
    );
    const duplicateIdentitySource = readPptx(zipFixture(duplicateIdentityFiles));
    const duplicateIdentityChart = duplicateIdentitySource.slides[0]?.shapes.find(
      (shape) => shape.kind === "chart",
    );
    if (duplicateIdentityChart?.handle === undefined)
      throw new Error("chart fixture should have a handle");
    expect(() =>
      updateChartData(duplicateIdentitySource, duplicateIdentityChart.handle, {
        series: [
          { name: "One", categories: ["A"], values: [1] },
          { name: "Two", categories: ["A"], values: [2] },
          { name: "Three", categories: ["A"], values: [3] },
        ],
      }),
    ).toThrow("chart series idx/order values must be unique");

    const explicitLegendFiles = unzipSync(input);
    explicitLegendFiles["ppt/charts/chart1.xml"] = replaceText(
      explicitLegendFiles["ppt/charts/chart1.xml"],
      "<c:legend>",
      '<c:legend><c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry>',
    );
    const explicitLegendSource = readPptx(zipFixture(explicitLegendFiles));
    const explicitLegendChart = explicitLegendSource.slides[0]?.shapes.find(
      (shape) => shape.kind === "chart",
    );
    if (explicitLegendChart?.handle === undefined)
      throw new Error("chart fixture should have a handle");
    expect(() =>
      updateChartData(explicitLegendSource, explicitLegendChart.handle, {
        series: [{ name: "Only", categories: ["A"], values: [1] }],
      }),
    ).toThrow("removing series with explicit legend entries is not supported");

    const source = readPptx(input);
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("chart fixture should have a handle");
    expect(() =>
      updateChartData(source, chart.handle, {
        series: Array.from({ length: 16_384 }, (_, index) => ({
          name: `Series ${index}`,
          categories: ["A"],
          values: [index],
        })),
      }),
    ).toThrow("series must not exceed 16383");
    expect(() =>
      updateChartData(source, chart.handle, {
        series: [
          {
            name: "Too many points",
            categories: new Array<string>(1_048_576),
            values: new Array<number>(1_048_576),
          },
        ],
      }),
    ).toThrow("data points must not exceed 1048575");
  });
});

describe("updateBubbleChartData", () => {
  it("updates XYZ formulas, caches, workbook tables, and series topology", () => {
    const input = buildExistingBubbleChart();
    const source = readPptx(input);
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("bubble fixture should have a handle");

    const edited = updateBubbleChartData(source, chart.handle, {
      series: [
        {
          name: "Edited revenue",
          xValues: [1, 2, 3],
          yValues: [40, 55, 70],
          bubbleSizes: [5, 8, 13],
        },
        { name: "Edited cost", xValues: [10], yValues: [25], bubbleSizes: [21] },
        { name: "Edited profit", xValues: [100], yValues: [28], bubbleSizes: [34] },
      ],
    });

    expect(source.edits).toBeUndefined();
    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "updateBubbleChartData",
      handle: chart.handle,
      chartPartPath: "ppt/charts/chart1.xml",
      workbookPartPath: "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
    });
    const computedChart = createComputedView(edited).slides[0]?.elements.find(
      (element) => element.kind === "chart",
    );
    expect(computedChart?.kind === "chart" ? computedChart.chartData : undefined).toMatchObject({
      chartType: "bubble",
      title: "Preserved title",
      series: [
        {
          name: "Edited revenue",
          xValues: [1, 2, 3],
          values: [40, 55, 70],
          bubbleSizes: [5, 8, 13],
        },
        { name: "Edited cost", xValues: [10], values: [25], bubbleSizes: [21] },
        { name: "Edited profit", xValues: [100], values: [28], bubbleSizes: [34] },
      ],
    });

    const before = unzipSync(input);
    const output = unzipSync(writePptx(edited));
    const chartXml = decoder.decode(output["ppt/charts/chart1.xml"]);
    expect(chartXml).toContain("<c:bubbleChart>");
    expect(chartXml).toContain("<c:title>");
    expect(chartXml).toContain("<c:legend>");
    expect(chartXml).toContain('<c:ext uri="preserve-me">');
    expect(chartXml.match(/<c:ser>/g)).toHaveLength(3);
    expect(chartXml).toContain('<c:idx val="2"/><c:order val="2"/>');
    expect(chartXml).toContain("Sheet1!$C$2:$C$4");
    expect(chartXml).toContain("Sheet1!$C$7:$C$7");
    expect(chartXml).toContain("Sheet1!$C$10:$C$10");
    expect(chartXml.match(/<c:ptCount val="3"\/>/g)).toHaveLength(3);
    expect(output["ppt/slides/slide1.xml"]).toEqual(before["ppt/slides/slide1.xml"]);
    expect(output["ppt/theme/theme1.xml"]).toEqual(before["ppt/theme/theme1.xml"]);

    const workbook = unzipSync(output["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]);
    const worksheet = decoder.decode(workbook["xl/worksheets/sheet1.xml"]);
    expect(worksheet).toContain('dimension ref="A1:C10"');
    expect(worksheet).toContain('<c r="C1" t="inlineStr"><is><t>Size</t></is></c>');
    expect(worksheet).toContain('<c r="C4"><v>13</v></c>');
    expect(worksheet).toContain('<c r="C6" t="inlineStr"><is><t>Size</t></is></c>');
    expect(worksheet).not.toContain('<row r="5"');
  });

  it("rejects invalid input, non-standard layout, and other chart types atomically", () => {
    const source = readPptx(buildExistingBubbleChart());
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("bubble fixture should have a handle");
    expect(() =>
      updateBubbleChartData(source, chart.handle, {
        series: [{ name: "Mismatch", xValues: [1, 2], yValues: [3, 4], bubbleSizes: [5] }],
      }),
    ).toThrow("matching non-empty X, Y, and bubble size counts");
    expect(() =>
      updateBubbleChartData(source, chart.handle, {
        series: [{ name: "Invalid", xValues: [1], yValues: [2], bubbleSizes: [Infinity] }],
      }),
    ).toThrow("X, Y, and bubble size values must be finite numbers");

    const malformed = unzipSync(buildExistingBubbleChart());
    malformed["ppt/charts/chart1.xml"] = replaceText(
      malformed["ppt/charts/chart1.xml"],
      "Sheet1!$C$6:$C$7",
      "Sheet1!$D$6:$D$7",
    );
    expectBubbleEditFailure(readPptx(zipFixture(malformed)), "unsupported data layout");

    const wrongHeader = unzipSync(buildExistingBubbleChart());
    const wrongHeaderWorkbook = unzipSync(
      wrongHeader["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"],
    );
    wrongHeaderWorkbook["xl/worksheets/sheet1.xml"] = replaceText(
      wrongHeaderWorkbook["xl/worksheets/sheet1.xml"],
      "<t>Size</t>",
      "<t>Radius</t>",
    );
    wrongHeader["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"] = zipFixture(wrongHeaderWorkbook);
    expectBubbleEditFailure(readPptx(zipFixture(wrongHeader)), "bubble size header must be 'Size'");

    const scatterSource = readPptx(buildExistingScatterChart());
    const scatter = scatterSource.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (scatter?.handle === undefined) throw new Error("scatter fixture should have a handle");
    expect(() =>
      updateBubbleChartData(scatterSource, scatter.handle, {
        series: [{ name: "Wrong", xValues: [1], yValues: [2], bubbleSizes: [3] }],
      }),
    ).toThrow("updateBubbleChartData: chart type or combination is not supported");

    const categorySource = readPptx(buildExistingChart());
    const category = categorySource.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (category?.handle === undefined) throw new Error("category fixture should have a handle");
    expect(() =>
      updateBubbleChartData(categorySource, category.handle, {
        series: [{ name: "Wrong", xValues: [1], yValues: [2], bubbleSizes: [3] }],
      }),
    ).toThrow("updateBubbleChartData: chart type or combination is not supported");

    const combo = unzipSync(buildExistingBubbleChart());
    combo["ppt/charts/chart1.xml"] = replaceText(
      combo["ppt/charts/chart1.xml"],
      "</c:bubbleChart>",
      "</c:bubbleChart><c:scatterChart/>",
    );
    expectBubbleEditFailure(
      readPptx(zipFixture(combo)),
      "chart type or combination is not supported",
    );
    expect(source.edits).toBeUndefined();
  });

  it("rejects formula cells and external, shared, or multi-sheet workbooks atomically", () => {
    const workbookPath = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
    const formulaFiles = unzipSync(buildExistingBubbleChart());
    const formulaWorkbook = unzipSync(formulaFiles[workbookPath]);
    formulaWorkbook["xl/worksheets/sheet1.xml"] = replaceText(
      formulaWorkbook["xl/worksheets/sheet1.xml"],
      "<v>4</v>",
      "<f>2+2</f><v>4</v>",
    );
    formulaFiles[workbookPath] = zipFixture(formulaWorkbook);
    expectBubbleEditFailure(
      readPptx(zipFixture(formulaFiles)),
      "formulas in the chart data range are not supported",
    );

    const external = unzipSync(buildExistingBubbleChart());
    external["ppt/charts/_rels/chart1.xml.rels"] = replaceText(
      external["ppt/charts/_rels/chart1.xml.rels"],
      'Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"',
      'Target="https://example.com/data.xlsx" TargetMode="External"',
    );
    expectBubbleEditFailure(
      readPptx(zipFixture(external)),
      "external workbook data is not supported",
    );
    expectBubbleEditFailure(
      readPptx(buildSharedBubbleWorkbookChart()),
      "embedded workbook is shared by another package part",
    );

    const multiple = unzipSync(buildExistingBubbleChart());
    const multipleWorkbook = unzipSync(multiple[workbookPath]);
    multipleWorkbook["xl/workbook.xml"] = replaceText(
      multipleWorkbook["xl/workbook.xml"],
      "</sheets>",
      '<sheet name="Other" sheetId="2" r:id="rId99"/></sheets>',
    );
    multiple[workbookPath] = zipFixture(multipleWorkbook);
    expectBubbleEditFailure(
      readPptx(zipFixture(multiple)),
      "embedded workbook must contain one matching worksheet",
    );
  });
});

describe("updateScatterChartData", () => {
  it("updates XY formulas, caches, workbook tables, and series topology", () => {
    const input = buildExistingScatterChart();
    const source = readPptx(input);
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("scatter fixture should have a handle");

    const edited = updateScatterChartData(source, chart.handle, {
      series: [
        { name: "Edited revenue", xValues: [1, 2, 3], yValues: [40, 55, 70] },
        { name: "Edited cost", xValues: [10, 20], yValues: [25, 42] },
        { name: "Edited profit", xValues: [100], yValues: [28] },
      ],
    });

    expect(edited).not.toBe(source);
    expect(source.edits).toBeUndefined();
    expect(edited.edits?.at(-1)).toMatchObject({
      kind: "updateScatterChartData",
      handle: chart.handle,
      chartPartPath: "ppt/charts/chart1.xml",
      workbookPartPath: "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
    });
    const computedChart = createComputedView(edited).slides[0]?.elements.find(
      (element) => element.kind === "chart",
    );
    expect(computedChart?.kind === "chart" ? computedChart.chartData : undefined).toMatchObject({
      chartType: "scatter",
      title: "Preserved title",
      series: [
        { name: "Edited revenue", xValues: [1, 2, 3], values: [40, 55, 70] },
        { name: "Edited cost", xValues: [10, 20], values: [25, 42] },
        { name: "Edited profit", xValues: [100], values: [28] },
      ],
    });

    const before = unzipSync(input);
    const beforeWorkbook = unzipSync(before["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]);
    const output = unzipSync(writePptx(edited));
    const chartXml = decoder.decode(output["ppt/charts/chart1.xml"]);
    expect(chartXml).toContain("<c:scatterChart>");
    expect(chartXml).toContain("<c:title>");
    expect(chartXml).toContain("<c:legend>");
    expect(chartXml).toContain('<c:ext uri="preserve-me">');
    expect(chartXml.match(/<c:valAx>/g)).toHaveLength(2);
    expect(chartXml).not.toContain("<c:catAx>");
    expect(chartXml).toContain("Preserved category axis");
    expect(chartXml).toContain("Preserved value axis");
    expect(chartXml.match(/<c:ser>/g)).toHaveLength(3);
    expect(chartXml).toContain('<c:idx val="0"/><c:order val="0"/>');
    expect(chartXml).toContain('<c:idx val="1"/><c:order val="1"/>');
    expect(chartXml).toContain('<c:idx val="2"/><c:order val="2"/>');
    expect(chartXml).toContain("Sheet1!$B$1");
    expect(chartXml).toContain("Sheet1!$A$2:$A$4");
    expect(chartXml).toContain("Sheet1!$B$2:$B$4");
    expect(chartXml).toContain("Sheet1!$B$6");
    expect(chartXml).toContain("Sheet1!$A$7:$A$8");
    expect(chartXml).toContain("Sheet1!$B$10");
    expect(chartXml).toContain("Sheet1!$A$11:$A$11");
    expect(chartXml.match(/<c:ptCount val="3"\/>/g)).toHaveLength(2);
    expect(chartXml.match(/<c:ptCount val="2"\/>/g)).toHaveLength(2);
    expect(chartXml.match(/uri="series-2"/g)).toHaveLength(2);
    expect(output["ppt/slides/slide1.xml"]).toEqual(before["ppt/slides/slide1.xml"]);
    expect(output["ppt/theme/theme1.xml"]).toEqual(before["ppt/theme/theme1.xml"]);

    const workbook = unzipSync(output["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]);
    const worksheet = decoder.decode(workbook["xl/worksheets/sheet1.xml"]);
    expect(worksheet).toContain('dimension ref="A1:B11"');
    expect(worksheet).toContain('<c r="B1" t="inlineStr"><is><t>Edited revenue</t></is></c>');
    expect(worksheet).toContain('<c r="A4"><v>3</v></c><c r="B4"><v>70</v></c>');
    expect(worksheet).toContain('<c r="B6" t="inlineStr"><is><t>Edited cost</t></is></c>');
    expect(worksheet).toContain('<c r="B10" t="inlineStr"><is><t>Edited profit</t></is></c>');
    expect(worksheet).not.toContain('<row r="5"');
    expect(worksheet).not.toContain('<row r="9"');
    expect(workbook["xl/styles.xml"]).toEqual(beforeWorkbook["xl/styles.xml"]);

    const reread = createComputedView(readPptx(writePptx(edited))).slides[0]?.elements.find(
      (element) => element.kind === "chart",
    );
    expect(reread?.kind === "chart" ? reread.chartData?.series[2] : undefined).toMatchObject({
      name: "Edited profit",
      xValues: [100],
      values: [28],
    });
  });

  it("removes trailing series and rejects unsupported or invalid inputs atomically", () => {
    const source = readPptx(buildExistingScatterChart());
    const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (chart?.handle === undefined) throw new Error("scatter fixture should have a handle");
    const removed = updateScatterChartData(source, chart.handle, {
      series: [{ name: "Only", xValues: [1, 2], yValues: [8, 13] }],
    });
    const removedFiles = unzipSync(writePptx(removed));
    expect(decoder.decode(removedFiles["ppt/charts/chart1.xml"]).match(/<c:ser>/g)).toHaveLength(1);
    expect(decoder.decode(removedFiles["ppt/charts/chart1.xml"])).not.toContain('uri="series-2"');
    expect(
      decoder.decode(
        unzipSync(removedFiles["ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"])[
          "xl/worksheets/sheet1.xml"
        ],
      ),
    ).toContain('dimension ref="A1:B3"');

    expect(() =>
      updateScatterChartData(source, chart.handle, {
        series: [{ name: "Mismatch", xValues: [1, 2], yValues: [3] }],
      }),
    ).toThrow("matching non-empty X and Y value counts");
    expect(() =>
      updateScatterChartData(source, chart.handle, {
        series: [{ name: "Non-finite", xValues: [1], yValues: [Number.NaN] }],
      }),
    ).toThrow("X and Y values must be finite numbers");
    expect(source.edits).toBeUndefined();

    const categorySource = readPptx(buildExistingChart());
    const categoryChart = categorySource.slides[0]?.shapes.find((shape) => shape.kind === "chart");
    if (categoryChart?.handle === undefined)
      throw new Error("category fixture should have a handle");
    expect(() =>
      updateScatterChartData(categorySource, categoryChart.handle, {
        series: [{ name: "Wrong operation", xValues: [1], yValues: [2] }],
      }),
    ).toThrow("updateScatterChartData: chart type or combination is not supported");
    expect(() =>
      updateChartData(source, chart.handle, {
        series: [{ name: "Wrong operation", categories: ["A"], values: [2] }],
      }),
    ).toThrow("updateChartData: chart type or combination is not supported");
  });

  it("rejects non-standard worksheet layouts and formula cells atomically", () => {
    const nonStandardFiles = unzipSync(buildExistingScatterChart());
    nonStandardFiles["ppt/charts/chart1.xml"] = replaceText(
      nonStandardFiles["ppt/charts/chart1.xml"],
      "Sheet1!$B$5",
      "Sheet1!$B$4",
    );
    expectScatterEditFailure(readPptx(zipFixture(nonStandardFiles)), "unsupported data layout");

    const formulaFiles = unzipSync(buildExistingScatterChart());
    const workbookPath = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
    const workbook = unzipSync(formulaFiles[workbookPath]);
    workbook["xl/worksheets/sheet1.xml"] = replaceText(
      workbook["xl/worksheets/sheet1.xml"],
      "<v>10</v>",
      "<f>5+5</f><v>10</v>",
    );
    formulaFiles[workbookPath] = zipFixture(workbook);
    expectScatterEditFailure(
      readPptx(zipFixture(formulaFiles)),
      "formulas in the chart data range are not supported",
    );

    for (const [search, replacement, context] of [
      ['<row r="1">', '<row r="1"><extLst/>', "row"],
      [
        '<c r="B1" t="inlineStr"><is><t>Revenue</t></is></c>',
        '<c r="B1" t="inlineStr"><is><t>Revenue</t></is><extLst/></c>',
        "cell",
      ],
    ] as const) {
      expectScatterEditFailure(
        readPptx(replaceEmbeddedWorksheetText(buildExistingScatterChart(), search, replacement)),
        `worksheet ${context} child XML is not supported`,
      );
    }

    const externalFiles = unzipSync(buildExistingScatterChart());
    externalFiles["ppt/charts/_rels/chart1.xml.rels"] = replaceText(
      externalFiles["ppt/charts/_rels/chart1.xml.rels"],
      'Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"',
      'Target="https://example.com/data.xlsx" TargetMode="External"',
    );
    expectScatterEditFailure(
      readPptx(zipFixture(externalFiles)),
      "external workbook data is not supported",
    );
    expectScatterEditFailure(
      readPptx(buildSharedScatterWorkbookChart()),
      "embedded workbook is shared by another package part",
    );

    const multipleWorksheetFiles = unzipSync(buildExistingScatterChart());
    const multipleWorksheetWorkbook = unzipSync(multipleWorksheetFiles[workbookPath]);
    multipleWorksheetWorkbook["xl/workbook.xml"] = replaceText(
      multipleWorksheetWorkbook["xl/workbook.xml"],
      "</sheets>",
      '<sheet name="Other" sheetId="2" r:id="rId99"/></sheets>',
    );
    multipleWorksheetFiles[workbookPath] = zipFixture(multipleWorksheetWorkbook);
    expectScatterEditFailure(
      readPptx(zipFixture(multipleWorksheetFiles)),
      "embedded workbook must contain one matching worksheet",
    );

    for (const unsupportedGroup of ["bubbleChart", "barChart"]) {
      const unsupportedFiles = unzipSync(buildExistingScatterChart());
      unsupportedFiles["ppt/charts/chart1.xml"] = replaceAllText(
        unsupportedFiles["ppt/charts/chart1.xml"],
        "scatterChart",
        unsupportedGroup,
      );
      expectScatterEditFailure(
        readPptx(zipFixture(unsupportedFiles)),
        "chart type or combination is not supported",
      );
    }

    const comboFiles = unzipSync(buildExistingScatterChart());
    comboFiles["ppt/charts/chart1.xml"] = replaceText(
      comboFiles["ppt/charts/chart1.xml"],
      "</c:scatterChart>",
      "</c:scatterChart><c:barChart/>",
    );
    expectScatterEditFailure(
      readPptx(zipFixture(comboFiles)),
      "chart type or combination is not supported",
    );
  });
});

function buildExistingScatterChart(): Uint8Array {
  const files = unzipSync(buildExistingChart());
  const chartPath = "ppt/charts/chart1.xml";
  files[chartPath] = replaceMatchingText(
    files[chartPath],
    /<c:barChart>.*?<\/c:barChart>/,
    `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>
      ${scatterSeriesXml(0, "Revenue", 1, [1, 2], [10, 20], "4472C4")}
      ${scatterSeriesXml(1, "Cost", 5, [3, 4], [7, 12], "ED7D31")}
      <c:axId val="100002"/><c:axId val="100003"/></c:scatterChart>`,
  );
  files[chartPath] = replaceCategoryAxisWithValueAxis(files[chartPath]);
  files[chartPath] = addSeriesExtensionMarkers(files[chartPath]);
  const workbookPath = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
  const workbook = unzipSync(files[workbookPath]);
  workbook["xl/worksheets/sheet1.xml"] = new TextEncoder().encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B7"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>20</v></c></row><row r="5"><c r="B5" t="inlineStr"><is><t>Cost</t></is></c></row><row r="6"><c r="A6"><v>3</v></c><c r="B6"><v>7</v></c></row><row r="7"><c r="A7"><v>4</v></c><c r="B7"><v>12</v></c></row></sheetData></worksheet>`,
  );
  files[workbookPath] = zipFixture(workbook);
  return zipFixture(files);
}

function buildExistingBubbleChart(): Uint8Array {
  const files = unzipSync(buildExistingChart());
  const chartPath = "ppt/charts/chart1.xml";
  files[chartPath] = replaceMatchingText(
    files[chartPath],
    /<c:barChart>.*?<\/c:barChart>/,
    `<c:bubbleChart><c:varyColors val="0"/>
      ${bubbleSeriesXml(0, "Revenue", 1, [1, 2], [10, 20], [4, 8], "4472C4")}
      ${bubbleSeriesXml(1, "Cost", 5, [3, 4], [7, 12], [6, 9], "ED7D31")}
      <c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:axId val="100002"/><c:axId val="100003"/></c:bubbleChart>`,
  );
  files[chartPath] = replaceCategoryAxisWithValueAxis(files[chartPath]);
  files[chartPath] = addSeriesExtensionMarkers(files[chartPath]);
  const workbookPath = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
  const workbook = unzipSync(files[workbookPath]);
  workbook["xl/worksheets/sheet1.xml"] = new TextEncoder().encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C7"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c><c r="C1" t="inlineStr"><is><t>Size</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>10</v></c><c r="C2"><v>4</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>20</v></c><c r="C3"><v>8</v></c></row><row r="5"><c r="B5" t="inlineStr"><is><t>Cost</t></is></c><c r="C5" t="inlineStr"><is><t>Size</t></is></c></row><row r="6"><c r="A6"><v>3</v></c><c r="B6"><v>7</v></c><c r="C6"><v>6</v></c></row><row r="7"><c r="A7"><v>4</v></c><c r="B7"><v>12</v></c><c r="C7"><v>9</v></c></row></sheetData></worksheet>`,
  );
  files[workbookPath] = zipFixture(workbook);
  return zipFixture(files);
}

function bubbleSeriesXml(
  index: number,
  name: string,
  headerRow: number,
  xValues: readonly number[],
  yValues: readonly number[],
  bubbleSizes: readonly number[],
  color: string,
): string {
  return `${scatterSeriesXml(index, name, headerRow, xValues, yValues, color).replace(
    "</c:ser>",
    `<c:bubbleSize><c:numRef><c:f>Sheet1!$C$${headerRow + 1}:$C$${headerRow + xValues.length}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${bubbleSizes.length}"/>${bubbleSizes.map((value, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${value}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:bubbleSize></c:ser>`,
  )}`;
}

function scatterSeriesXml(
  index: number,
  name: string,
  headerRow: number,
  xValues: readonly number[],
  yValues: readonly number[],
  color: string,
): string {
  const lastRow = headerRow + xValues.length;
  const points = (values: readonly number[]) =>
    values
      .map((value, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${value}</c:v></c:pt>`)
      .join("");
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:strRef><c:f>Sheet1!$B$${headerRow}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr><c:xVal><c:numRef><c:f>Sheet1!$A$${headerRow + 1}:$A$${lastRow}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${xValues.length}"/>${points(xValues)}</c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>Sheet1!$B$${headerRow + 1}:$B$${lastRow}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${yValues.length}"/>${points(yValues)}</c:numCache></c:numRef></c:yVal></c:ser>`;
}

function buildExistingChart(): Uint8Array {
  let source = createPptx();
  const slideHandle = source.slides[0]?.handle;
  if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
  source = addChart(source, slideHandle, {
    chartType: "bar",
    offsetX: asEmu(500000),
    offsetY: asEmu(500000),
    width: asEmu(5000000),
    height: asEmu(3000000),
    title: "Preserved title",
    showLegend: true,
    categoryAxis: { title: "Preserved category axis" },
    valueAxis: { title: "Preserved value axis" },
    series: [
      {
        name: "Revenue",
        categories: ["Jan", "Feb"],
        values: [10, 20],
        color: "4472C4",
      },
      {
        name: "Cost",
        categories: ["Jan", "Feb"],
        values: [7, 12],
        color: "ED7D31",
      },
    ],
  });
  const files = unzipSync(writePptx(source));
  files["ppt/charts/chart1.xml"] = replaceText(
    files["ppt/charts/chart1.xml"],
    "</c:chartSpace>",
    '<c:extLst><c:ext uri="preserve-me"><c:unknown val="kept"/></c:ext></c:extLst></c:chartSpace>',
  );
  return zipFixture(files);
}

function buildExistingComboChart(): Uint8Array {
  const files = unzipSync(buildExistingChart());
  const chartPath = "ppt/charts/chart1.xml";
  let chartXml = decoder.decode(files[chartPath]);
  const barChart = /<c:barChart>.*?<\/c:barChart>/.exec(chartXml)?.[0];
  const series = barChart?.match(/<c:ser>.*?<\/c:ser>/g);
  if (barChart === undefined || series?.length !== 2) {
    throw new Error("category chart fixture does not contain two series");
  }
  const lineSeries = series[1]
    .replace('<c:idx val="1"/><c:order val="1"/>', '<c:idx val="7"/><c:order val="9"/>')
    .replace("</c:ser>", '<c:extLst><c:ext uri="line-series-preserve"/></c:extLst></c:ser>');
  const barOnly = barChart.replace(series[1], "");
  const lineChart = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${lineSeries}<c:axId val="100002"/><c:axId val="200003"/><c:extLst><c:ext uri="combo-group-preserve"/></c:extLst></c:lineChart>`;
  chartXml = chartXml.replace(barChart, `${barOnly}${lineChart}`);
  chartXml = chartXml.replace(
    "</c:plotArea>",
    '<c:valAx><c:axId val="200003"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="100002"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx></c:plotArea>',
  );
  files[chartPath] = new TextEncoder().encode(chartXml);
  const workbookPath = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
  const workbook = unzipSync(files[workbookPath]);
  workbook["xl/combo-preserve.xml"] = new TextEncoder().encode("<preserved/>");
  files[workbookPath] = zipFixture(workbook);
  return zipFixture(files);
}

function buildSharedWorkbookChart(): Uint8Array {
  let source = createPptx();
  const slideHandle = source.slides[0]?.handle;
  if (slideHandle === undefined) throw new Error("createPptx should create a first slide");
  for (let index = 0; index < 2; index += 1) {
    source = addChart(source, slideHandle, {
      chartType: "bar",
      offsetX: asEmu(index * 1000),
      offsetY: asEmu(0),
      width: asEmu(1000),
      height: asEmu(1000),
      series: [
        { name: "One", categories: ["A"], values: [1] },
        { name: "Two", categories: ["A"], values: [2] },
      ],
    });
  }
  const files = unzipSync(writePptx(source));
  files["ppt/charts/_rels/chart2.xml.rels"] = replaceText(
    files["ppt/charts/_rels/chart2.xml.rels"],
    "../embeddings/Microsoft_Excel_Worksheet2.xlsx",
    "../embeddings/Microsoft_Excel_Worksheet1.xlsx",
  );
  return zipFixture(files);
}

function buildSharedScatterWorkbookChart(): Uint8Array {
  const files = unzipSync(buildSharedWorkbookChart());
  files["ppt/charts/chart1.xml"] = replaceMatchingText(
    files["ppt/charts/chart1.xml"],
    /<c:barChart>.*?<\/c:barChart>/,
    `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>
      ${scatterSeriesXml(0, "One", 1, [1], [1], "4472C4")}
      ${scatterSeriesXml(1, "Two", 4, [2], [2], "ED7D31")}
      <c:axId val="100002"/><c:axId val="100003"/></c:scatterChart>`,
  );
  files["ppt/charts/chart1.xml"] = replaceCategoryAxisWithValueAxis(files["ppt/charts/chart1.xml"]);
  return zipFixture(files);
}

function buildSharedBubbleWorkbookChart(): Uint8Array {
  const files = unzipSync(buildSharedWorkbookChart());
  const chartPath = "ppt/charts/chart1.xml";
  files[chartPath] = replaceMatchingText(
    files[chartPath],
    /<c:barChart>.*?<\/c:barChart>/,
    `<c:bubbleChart><c:varyColors val="0"/>
      ${bubbleSeriesXml(0, "One", 1, [1], [1], [2], "4472C4")}
      ${bubbleSeriesXml(1, "Two", 4, [2], [2], [3], "ED7D31")}
      <c:axId val="100002"/><c:axId val="100003"/></c:bubbleChart>`,
  );
  files[chartPath] = replaceCategoryAxisWithValueAxis(files[chartPath]);
  return zipFixture(files);
}

function expectEditFailure(source: ReturnType<typeof readPptx>, message: string): void {
  const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
  if (chart?.handle === undefined) throw new Error("chart fixture should have a handle");
  expect(() =>
    updateChartData(source, chart.handle, {
      series: [
        { name: "Edited 1", categories: ["A"], values: [1] },
        { name: "Edited 2", categories: ["A"], values: [2] },
      ],
    }),
  ).toThrow(message);
  expect(source.edits).toBeUndefined();
}

function expectScatterEditFailure(source: ReturnType<typeof readPptx>, message: string): void {
  const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
  if (chart?.handle === undefined) throw new Error("scatter fixture should have a handle");
  expect(() =>
    updateScatterChartData(source, chart.handle, {
      series: [
        { name: "Edited 1", xValues: [1], yValues: [2] },
        { name: "Edited 2", xValues: [3], yValues: [4] },
      ],
    }),
  ).toThrow(message);
  expect(source.edits).toBeUndefined();
}

function expectBubbleEditFailure(source: ReturnType<typeof readPptx>, message: string): void {
  const chart = source.slides[0]?.shapes.find((shape) => shape.kind === "chart");
  if (chart?.handle === undefined) throw new Error("bubble fixture should have a handle");
  expect(() =>
    updateBubbleChartData(source, chart.handle, {
      series: [
        { name: "Edited 1", xValues: [1], yValues: [2], bubbleSizes: [3] },
        { name: "Edited 2", xValues: [4], yValues: [5], bubbleSizes: [6] },
      ],
    }),
  ).toThrow(message);
  expect(source.edits).toBeUndefined();
}

function replaceText(bytes: Uint8Array, search: string, replacement: string): Uint8Array {
  const value = decoder.decode(bytes);
  if (!value.includes(search)) throw new Error(`fixture text not found: ${search}`);
  return new TextEncoder().encode(value.replace(search, replacement));
}

function replaceAllText(bytes: Uint8Array, search: string, replacement: string): Uint8Array {
  const value = decoder.decode(bytes);
  if (!value.includes(search)) throw new Error(`fixture text not found: ${search}`);
  return new TextEncoder().encode(value.replaceAll(search, replacement));
}

function replaceMatchingText(bytes: Uint8Array, search: RegExp, replacement: string): Uint8Array {
  const value = decoder.decode(bytes);
  if (!search.test(value)) throw new Error(`fixture pattern not found: ${String(search)}`);
  return new TextEncoder().encode(value.replace(search, replacement));
}

function replaceCategoryAxisWithValueAxis(bytes: Uint8Array): Uint8Array {
  const value = decoder.decode(bytes);
  const categoryAxis = /<c:catAx>.*?<\/c:catAx>/.exec(value)?.[0];
  if (categoryAxis === undefined) throw new Error("fixture category axis not found");
  const valueAxis = categoryAxis
    .replace("<c:catAx>", "<c:valAx>")
    .replace("</c:catAx>", '<c:crossBetween val="midCat"/></c:valAx>')
    .replace(/<c:(?:auto|lblAlgn|lblOffset|noMultiLvlLbl)\b[^>]*\/>/g, "");
  return new TextEncoder().encode(value.replace(categoryAxis, valueAxis));
}

function replaceEmbeddedWorksheetText(
  pptxBytes: Uint8Array,
  search: string,
  replacement: string,
): Uint8Array {
  const files = unzipSync(pptxBytes);
  const workbookPath = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
  const workbook = unzipSync(files[workbookPath]);
  workbook["xl/worksheets/sheet1.xml"] = replaceText(
    workbook["xl/worksheets/sheet1.xml"],
    search,
    replacement,
  );
  files[workbookPath] = zipFixture(workbook);
  return zipFixture(files);
}

function addSeriesExtensionMarkers(bytes: Uint8Array): Uint8Array {
  let index = 0;
  return new TextEncoder().encode(
    decoder.decode(bytes).replaceAll("</c:ser>", () => {
      index += 1;
      const uniqueId = String(index).padStart(12, "0");
      return `<c:extLst><c:ext uri="series-${index}"><c16:uniqueId xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart" val="${index}"/><c:unknown val="kept"/></c:ext><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"><c16:uniqueId xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart" val="{00000000-0000-0000-0000-${uniqueId}}"/></c:ext></c:extLst></c:ser>`;
    }),
  );
}

function zipFixture(files: Record<string, Uint8Array>): Uint8Array {
  return zipSync(files);
}
