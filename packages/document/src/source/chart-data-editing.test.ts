import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  addChart,
  asEmu,
  createComputedView,
  createPptx,
  readPptx,
  updateChartData,
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
        series: [{ name: "Edited", categories: ["X", "Y", "Z"], values: [3, 5, 8] }],
      });
    }

    const computedCharts = createComputedView(
      readPptx(writePptx(source)),
    ).slides[0]?.elements.filter((element) => element.kind === "chart");
    expect(computedCharts?.map((chart) => chart.chartData?.chartType)).toEqual(types);
    expect(computedCharts?.every((chart) => chart.chartData?.categories.length === 3)).toBe(true);
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

function zipFixture(files: Record<string, Uint8Array>): Uint8Array {
  return zipSync(files);
}
