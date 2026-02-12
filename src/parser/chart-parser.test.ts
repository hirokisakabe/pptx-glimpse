import { describe, it, expect } from "vitest";
import { parseChart } from "./chart-parser.js";
import { ColorResolver } from "../color/color-resolver.js";

function createColorResolver() {
  return new ColorResolver(
    {
      dk1: "#000000",
      lt1: "#FFFFFF",
      dk2: "#44546A",
      lt2: "#E7E6E6",
      accent1: "#4472C4",
      accent2: "#ED7D31",
      accent3: "#A5A5A5",
      accent4: "#FFC000",
      accent5: "#5B9BD5",
      accent6: "#70AD47",
      hlink: "#0563C1",
      folHlink: "#954F72",
    },
    {
      bg1: "lt1",
      tx1: "dk1",
      bg2: "lt2",
      tx2: "dk2",
      accent1: "accent1",
      accent2: "accent2",
      accent3: "accent3",
      accent4: "accent4",
      accent5: "accent5",
      accent6: "accent6",
      hlink: "hlink",
      folHlink: "folHlink",
    },
  );
}

function barChartXml(options?: { title?: string; barDir?: string; legend?: string }) {
  const titleXml = options?.title
    ? `<c:title><c:tx><c:rich><a:p><a:r><a:t>${options.title}</a:t></a:r></a:p></c:rich></c:tx></c:title>`
    : "";
  const legendXml = options?.legend
    ? `<c:legend><c:legendPos val="${options.legend}"/></c:legend>`
    : "";
  const barDir = options?.barDir ?? "col";

  return `<?xml version="1.0" encoding="UTF-8"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <c:chart>
        ${titleXml}
        <c:plotArea>
          <c:barChart>
            <c:barDir val="${barDir}"/>
            <c:ser>
              <c:idx val="0"/>
              <c:order val="0"/>
              <c:tx>
                <c:strRef>
                  <c:strCache>
                    <c:ptCount val="1"/>
                    <c:pt idx="0"><c:v>Series 1</c:v></c:pt>
                  </c:strCache>
                </c:strRef>
              </c:tx>
              <c:cat>
                <c:strRef>
                  <c:strCache>
                    <c:ptCount val="3"/>
                    <c:pt idx="0"><c:v>Cat A</c:v></c:pt>
                    <c:pt idx="1"><c:v>Cat B</c:v></c:pt>
                    <c:pt idx="2"><c:v>Cat C</c:v></c:pt>
                  </c:strCache>
                </c:strRef>
              </c:cat>
              <c:val>
                <c:numRef>
                  <c:numCache>
                    <c:ptCount val="3"/>
                    <c:pt idx="0"><c:v>4.3</c:v></c:pt>
                    <c:pt idx="1"><c:v>2.5</c:v></c:pt>
                    <c:pt idx="2"><c:v>3.5</c:v></c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
            <c:ser>
              <c:idx val="1"/>
              <c:order val="1"/>
              <c:tx>
                <c:strRef>
                  <c:strCache>
                    <c:ptCount val="1"/>
                    <c:pt idx="0"><c:v>Series 2</c:v></c:pt>
                  </c:strCache>
                </c:strRef>
              </c:tx>
              <c:val>
                <c:numRef>
                  <c:numCache>
                    <c:ptCount val="3"/>
                    <c:pt idx="0"><c:v>2.4</c:v></c:pt>
                    <c:pt idx="1"><c:v>4.4</c:v></c:pt>
                    <c:pt idx="2"><c:v>1.8</c:v></c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
          </c:barChart>
        </c:plotArea>
        ${legendXml}
      </c:chart>
    </c:chartSpace>`;
}

describe("parseChart", () => {
  it("parses a bar chart with series and categories", () => {
    const result = parseChart(barChartXml(), createColorResolver());

    expect(result).not.toBeNull();
    expect(result!.chartType).toBe("bar");
    expect(result!.barDirection).toBe("col");
    expect(result!.categories).toEqual(["Cat A", "Cat B", "Cat C"]);
    expect(result!.series).toHaveLength(2);

    expect(result!.series[0].name).toBe("Series 1");
    expect(result!.series[0].values).toEqual([4.3, 2.5, 3.5]);

    expect(result!.series[1].name).toBe("Series 2");
    expect(result!.series[1].values).toEqual([2.4, 4.4, 1.8]);
  });

  it("parses horizontal bar direction", () => {
    const result = parseChart(barChartXml({ barDir: "bar" }), createColorResolver());
    expect(result!.barDirection).toBe("bar");
  });

  it("parses chart title", () => {
    const result = parseChart(barChartXml({ title: "Sales Data" }), createColorResolver());
    expect(result!.title).toBe("Sales Data");
  });

  it("returns null title when not present", () => {
    const result = parseChart(barChartXml(), createColorResolver());
    expect(result!.title).toBeNull();
  });

  it("parses legend position", () => {
    const result = parseChart(barChartXml({ legend: "t" }), createColorResolver());
    expect(result!.legend).toEqual({ position: "t" });
  });

  it("returns null legend when not present", () => {
    const result = parseChart(barChartXml(), createColorResolver());
    expect(result!.legend).toBeNull();
  });

  it("assigns default accent colors to series without explicit color", () => {
    const result = parseChart(barChartXml(), createColorResolver());
    expect(result!.series[0].color.hex).toBe("#4472C4"); // accent1
    expect(result!.series[1].color.hex).toBe("#ED7D31"); // accent2
  });

  it("parses a line chart", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart>
          <c:plotArea>
            <c:lineChart>
              <c:ser>
                <c:idx val="0"/>
                <c:tx><c:strRef><c:strCache>
                  <c:pt idx="0"><c:v>Line 1</c:v></c:pt>
                </c:strCache></c:strRef></c:tx>
                <c:val><c:numRef><c:numCache>
                  <c:pt idx="0"><c:v>10</c:v></c:pt>
                  <c:pt idx="1"><c:v>20</c:v></c:pt>
                </c:numCache></c:numRef></c:val>
              </c:ser>
            </c:lineChart>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>`;

    const result = parseChart(xml, createColorResolver());
    expect(result!.chartType).toBe("line");
    expect(result!.series[0].name).toBe("Line 1");
    expect(result!.series[0].values).toEqual([10, 20]);
  });

  it("parses a pie chart", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart>
          <c:plotArea>
            <c:pieChart>
              <c:ser>
                <c:idx val="0"/>
                <c:cat><c:strRef><c:strCache>
                  <c:pt idx="0"><c:v>A</c:v></c:pt>
                  <c:pt idx="1"><c:v>B</c:v></c:pt>
                </c:strCache></c:strRef></c:cat>
                <c:val><c:numRef><c:numCache>
                  <c:pt idx="0"><c:v>60</c:v></c:pt>
                  <c:pt idx="1"><c:v>40</c:v></c:pt>
                </c:numCache></c:numRef></c:val>
              </c:ser>
            </c:pieChart>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>`;

    const result = parseChart(xml, createColorResolver());
    expect(result!.chartType).toBe("pie");
    expect(result!.categories).toEqual(["A", "B"]);
    expect(result!.series[0].values).toEqual([60, 40]);
  });

  it("parses a doughnut chart with holeSize", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart>
          <c:plotArea>
            <c:doughnutChart>
              <c:ser>
                <c:idx val="0"/>
                <c:cat><c:strRef><c:strCache>
                  <c:pt idx="0"><c:v>A</c:v></c:pt>
                  <c:pt idx="1"><c:v>B</c:v></c:pt>
                </c:strCache></c:strRef></c:cat>
                <c:val><c:numRef><c:numCache>
                  <c:pt idx="0"><c:v>60</c:v></c:pt>
                  <c:pt idx="1"><c:v>40</c:v></c:pt>
                </c:numCache></c:numRef></c:val>
              </c:ser>
              <c:holeSize val="75"/>
            </c:doughnutChart>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>`;

    const result = parseChart(xml, createColorResolver());
    expect(result!.chartType).toBe("doughnut");
    expect(result!.holeSize).toBe(75);
    expect(result!.categories).toEqual(["A", "B"]);
    expect(result!.series[0].values).toEqual([60, 40]);
  });

  it("parses a doughnut chart with default holeSize", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart>
          <c:plotArea>
            <c:doughnutChart>
              <c:ser>
                <c:idx val="0"/>
                <c:val><c:numRef><c:numCache>
                  <c:pt idx="0"><c:v>100</c:v></c:pt>
                </c:numCache></c:numRef></c:val>
              </c:ser>
            </c:doughnutChart>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>`;

    const result = parseChart(xml, createColorResolver());
    expect(result!.chartType).toBe("doughnut");
    expect(result!.holeSize).toBe(50);
  });

  it("parses a scatter chart with xValues", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart>
          <c:plotArea>
            <c:scatterChart>
              <c:ser>
                <c:idx val="0"/>
                <c:xVal><c:numRef><c:numCache>
                  <c:pt idx="0"><c:v>1</c:v></c:pt>
                  <c:pt idx="1"><c:v>2</c:v></c:pt>
                  <c:pt idx="2"><c:v>3</c:v></c:pt>
                </c:numCache></c:numRef></c:xVal>
                <c:yVal><c:numRef><c:numCache>
                  <c:pt idx="0"><c:v>10</c:v></c:pt>
                  <c:pt idx="1"><c:v>20</c:v></c:pt>
                  <c:pt idx="2"><c:v>15</c:v></c:pt>
                </c:numCache></c:numRef></c:yVal>
              </c:ser>
            </c:scatterChart>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>`;

    const result = parseChart(xml, createColorResolver());
    expect(result!.chartType).toBe("scatter");
    expect(result!.series[0].xValues).toEqual([1, 2, 3]);
    expect(result!.series[0].values).toEqual([10, 20, 15]);
  });

  it("detects 3D chart types", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart>
          <c:plotArea>
            <c:bar3DChart>
              <c:barDir val="col"/>
              <c:ser>
                <c:idx val="0"/>
                <c:val><c:numRef><c:numCache>
                  <c:pt idx="0"><c:v>5</c:v></c:pt>
                </c:numCache></c:numRef></c:val>
              </c:ser>
            </c:bar3DChart>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>`;

    const result = parseChart(xml, createColorResolver());
    expect(result!.chartType).toBe("bar");
  });

  it("returns null for empty chartSpace", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
      </c:chartSpace>`;

    const result = parseChart(xml, createColorResolver());
    expect(result).toBeNull();
  });

  it("returns null when no recognized chart type in plotArea", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart>
          <c:plotArea>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>`;

    const result = parseChart(xml, createColorResolver());
    expect(result).toBeNull();
  });

  it("parses series with explicit solidFill color", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:chart>
          <c:plotArea>
            <c:barChart>
              <c:barDir val="col"/>
              <c:ser>
                <c:idx val="0"/>
                <c:spPr>
                  <a:solidFill>
                    <a:srgbClr val="FF0000"/>
                  </a:solidFill>
                </c:spPr>
                <c:val><c:numRef><c:numCache>
                  <c:pt idx="0"><c:v>5</c:v></c:pt>
                </c:numCache></c:numRef></c:val>
              </c:ser>
            </c:barChart>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>`;

    const result = parseChart(xml, createColorResolver());
    expect(result!.series[0].color.hex).toBe("#FF0000");
  });
});
