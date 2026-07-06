import {
  buildPptx,
  NS,
  REL_TYPES,
  savePptx,
  SLIDE_H,
  SLIDE_W,
  slideRelsXml,
  wrapSlideXml,
} from "../fixture-builder.js";

function chartXml(
  chartType: string,
  opts: {
    barDir?: string;
    holeSize?: number;
    radarStyle?: string;
    ofPieType?: string;
    secondPieSize?: number;
    splitPos?: number;
    title?: string;
    legendPos?: string;
    series: {
      name: string;
      categories?: string[];
      values: number[];
      xValues?: number[];
      bubbleSizes?: number[];
    }[];
  },
): string {
  const titleXml = opts.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${opts.title}</a:t></a:r></a:p></c:rich></c:tx></c:title>`
    : "";
  const legendXml = opts.legendPos
    ? `<c:legend><c:legendPos val="${opts.legendPos}"/></c:legend>`
    : "";

  const seriesXml = opts.series
    .map((s, i) => {
      const nameXml = `<c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>${s.name}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
      const catXml = s.categories
        ? `<c:cat><c:strRef><c:strCache>${s.categories.map((c, j) => `<c:pt idx="${j}"><c:v>${c}</c:v></c:pt>`).join("")}</c:strCache></c:strRef></c:cat>`
        : "";
      const usesXY = chartType === "scatterChart" || chartType === "bubbleChart";
      const valTag = usesXY ? "c:yVal" : "c:val";
      const valXml = `<${valTag}><c:numRef><c:numCache>${s.values.map((v, j) => `<c:pt idx="${j}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></${valTag}>`;
      const xValXml = s.xValues
        ? `<c:xVal><c:numRef><c:numCache>${s.xValues.map((v, j) => `<c:pt idx="${j}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:xVal>`
        : "";
      const bubbleSizeXml = s.bubbleSizes
        ? `<c:bubbleSize><c:numRef><c:numCache>${s.bubbleSizes.map((v, j) => `<c:pt idx="${j}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:bubbleSize>`
        : "";
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${nameXml}${catXml}${xValXml}${valXml}${bubbleSizeXml}</c:ser>`;
    })
    .join("");

  const barDirXml = opts.barDir ? `<c:barDir val="${opts.barDir}"/>` : "";
  const holeSizeXml = opts.holeSize !== undefined ? `<c:holeSize val="${opts.holeSize}"/>` : "";
  const radarStyleXml = opts.radarStyle ? `<c:radarStyle val="${opts.radarStyle}"/>` : "";
  const ofPieTypeXml = opts.ofPieType ? `<c:ofPieType val="${opts.ofPieType}"/>` : "";
  const secondPieSizeXml =
    opts.secondPieSize !== undefined ? `<c:secondPieSize val="${opts.secondPieSize}"/>` : "";
  const splitPosXml = opts.splitPos !== undefined ? `<c:splitPos val="${opts.splitPos}"/>` : "";

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}">
  <c:chart>
    ${titleXml}
    <c:plotArea>
      <c:${chartType}>
        ${radarStyleXml}
        ${barDirXml}
        ${ofPieTypeXml}
        ${seriesXml}
        ${holeSizeXml}
        ${splitPosXml}
        ${secondPieSizeXml}
      </c:${chartType}>
    </c:plotArea>
    ${legendXml}
  </c:chart>
</c:chartSpace>`;
}

function graphicFrameXml(
  id: number,
  name: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  chartRId: string,
): string {
  return `<p:graphicFrame>
  <p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
  <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:chart xmlns:c="${NS.c}" r:id="${chartRId}"/>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`;
}

async function createChartsFixture(): Promise<void> {
  const charts = new Map<string, string>();
  const slides: SlideData[] = [];

  const margin = 300000;

  // Slide 1: Bar chart
  const barChart = chartXml("barChart", {
    barDir: "col",
    title: "Sales by Quarter",
    legendPos: "b",
    series: [
      { name: "FY2024", categories: ["Q1", "Q2", "Q3", "Q4"], values: [10, 25, 15, 30] },
      { name: "FY2025", categories: ["Q1", "Q2", "Q3", "Q4"], values: [15, 20, 25, 35] },
    ],
  });
  charts.set("ppt/charts/chart1.xml", barChart);
  const gf1 = graphicFrameXml(
    2,
    "Bar Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf1),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart1.xml" }]),
  });

  // Slide 2: Line chart
  const lineChart = chartXml("lineChart", {
    title: "Monthly Trend",
    legendPos: "b",
    series: [
      {
        name: "Revenue",
        categories: ["Jan", "Feb", "Mar", "Apr", "May"],
        values: [100, 120, 90, 150, 130],
      },
    ],
  });
  charts.set("ppt/charts/chart2.xml", lineChart);
  const gf2 = graphicFrameXml(
    2,
    "Line Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf2),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart2.xml" }]),
  });

  // Slide 3: Pie chart
  const pieChart = chartXml("pieChart", {
    title: "Market Share",
    legendPos: "r",
    series: [{ name: "Share", categories: ["A", "B", "C", "D"], values: [40, 25, 20, 15] }],
  });
  charts.set("ppt/charts/chart3.xml", pieChart);
  const gf3 = graphicFrameXml(
    2,
    "Pie Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf3),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart3.xml" }]),
  });

  // Slide 4: Scatter chart
  const scatterChart = chartXml("scatterChart", {
    title: "Data Points",
    legendPos: "b",
    series: [{ name: "Dataset", xValues: [1, 2, 3, 5, 8], values: [2, 4, 3, 7, 6] }],
  });
  charts.set("ppt/charts/chart4.xml", scatterChart);
  const gf4 = graphicFrameXml(
    2,
    "Scatter Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf4),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart4.xml" }]),
  });

  // Slide 5: Doughnut chart
  const doughnutChart = chartXml("doughnutChart", {
    title: "Budget Allocation",
    legendPos: "r",
    holeSize: 60,
    series: [
      {
        name: "Budget",
        categories: ["Dev", "Marketing", "Sales", "Support"],
        values: [35, 25, 25, 15],
      },
    ],
  });
  charts.set("ppt/charts/chart5.xml", doughnutChart);
  const gf5 = graphicFrameXml(
    2,
    "Doughnut Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf5),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart5.xml" }]),
  });

  // Slide 6: Bubble chart
  const bubbleChart = chartXml("bubbleChart", {
    title: "Bubble Data",
    legendPos: "b",
    series: [
      {
        name: "Dataset A",
        xValues: [1, 3, 5, 7, 9],
        values: [10, 30, 20, 40, 25],
        bubbleSizes: [4, 8, 12, 6, 16],
      },
      {
        name: "Dataset B",
        xValues: [2, 4, 6, 8],
        values: [15, 25, 35, 10],
        bubbleSizes: [10, 5, 14, 8],
      },
    ],
  });
  charts.set("ppt/charts/chart6.xml", bubbleChart);
  const gf6 = graphicFrameXml(
    2,
    "Bubble Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf6),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart6.xml" }]),
  });

  // Slide 7: Area chart
  const areaChart = chartXml("areaChart", {
    title: "Website Traffic",
    legendPos: "b",
    series: [
      {
        name: "Visitors",
        categories: ["Mon", "Tue", "Wed", "Thu", "Fri"],
        values: [200, 350, 280, 420, 380],
      },
      {
        name: "Page Views",
        categories: ["Mon", "Tue", "Wed", "Thu", "Fri"],
        values: [400, 500, 450, 600, 550],
      },
    ],
  });
  charts.set("ppt/charts/chart7.xml", areaChart);
  const gf7 = graphicFrameXml(
    2,
    "Area Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf7),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart7.xml" }]),
  });

  // Slide 8: Radar chart
  const radarChart = chartXml("radarChart", {
    radarStyle: "marker",
    title: "Skill Assessment",
    legendPos: "b",
    series: [
      {
        name: "Team A",
        categories: ["Speed", "Power", "Accuracy", "Endurance", "Agility"],
        values: [8, 7, 9, 6, 8],
      },
      {
        name: "Team B",
        categories: ["Speed", "Power", "Accuracy", "Endurance", "Agility"],
        values: [7, 8, 7, 8, 7],
      },
    ],
  });
  charts.set("ppt/charts/chart8.xml", radarChart);
  const gf8 = graphicFrameXml(
    2,
    "Radar Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf8),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart8.xml" }]),
  });

  // Slide 9: Stock chart (HLC)
  const stockChart = chartXml("stockChart", {
    title: "Stock Price (HLC)",
    legendPos: "b",
    series: [
      {
        name: "High",
        categories: ["Mon", "Tue", "Wed", "Thu", "Fri"],
        values: [150, 160, 155, 170, 165],
      },
      {
        name: "Low",
        categories: ["Mon", "Tue", "Wed", "Thu", "Fri"],
        values: [100, 110, 105, 120, 115],
      },
      {
        name: "Close",
        categories: ["Mon", "Tue", "Wed", "Thu", "Fri"],
        values: [130, 140, 135, 150, 145],
      },
    ],
  });
  charts.set("ppt/charts/chart9.xml", stockChart);
  const gf9 = graphicFrameXml(
    2,
    "Stock Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf9),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart9.xml" }]),
  });

  // Slide 10: Surface chart (contour/heatmap)
  const surfaceChart = chartXml("surfaceChart", {
    title: "Temperature Map",
    series: [
      { name: "North", categories: ["Jan", "Apr", "Jul", "Oct"], values: [5, 15, 30, 12] },
      { name: "Central", categories: ["Jan", "Apr", "Jul", "Oct"], values: [10, 20, 35, 18] },
      { name: "South", categories: ["Jan", "Apr", "Jul", "Oct"], values: [15, 25, 40, 22] },
    ],
  });
  charts.set("ppt/charts/chart10.xml", surfaceChart);
  const gf10 = graphicFrameXml(
    2,
    "Surface Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf10),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart10.xml" }]),
  });

  // Slide 11: OfPie chart (pie-of-pie)
  const ofPieChartData = chartXml("ofPieChart", {
    ofPieType: "pie",
    splitPos: 3,
    secondPieSize: 75,
    title: "Revenue Breakdown",
    legendPos: "b",
    series: [
      {
        name: "Revenue",
        categories: ["Product A", "Product B", "Service X", "Service Y", "Other"],
        values: [45, 25, 15, 10, 5],
      },
    ],
  });
  charts.set("ppt/charts/chart11.xml", ofPieChartData);
  const gf11 = graphicFrameXml(
    2,
    "OfPie Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf11),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart11.xml" }]),
  });

  const buffer = await buildPptx({
    slides,
    charts,
    contentTypesExtra: [
      `<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart4.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart5.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart6.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart7.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart8.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart9.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart10.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart11.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
    ],
  });
  savePptx(buffer, "charts.pptx");
}

// --- 8. Connectors ---

async function createCharts3dFallbackFixture(): Promise<void> {
  const charts = new Map<string, string>();
  const slides: SlideData[] = [];
  const margin = 300000;

  // Slide 1: bar3DChart -> bar fallback
  const bar3D = chartXml("bar3DChart", {
    barDir: "col",
    title: "3D Bar (fallback)",
    legendPos: "b",
    series: [
      { name: "FY2024", categories: ["Q1", "Q2", "Q3", "Q4"], values: [10, 25, 15, 30] },
      { name: "FY2025", categories: ["Q1", "Q2", "Q3", "Q4"], values: [15, 20, 25, 35] },
    ],
  });
  charts.set("ppt/charts/chart1.xml", bar3D);
  const gf1 = graphicFrameXml(
    2,
    "3D Bar Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf1),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart1.xml" }]),
  });

  // Slide 2: pie3DChart -> pie fallback
  const pie3D = chartXml("pie3DChart", {
    title: "3D Pie (fallback)",
    legendPos: "r",
    series: [{ name: "Share", categories: ["A", "B", "C", "D"], values: [40, 25, 20, 15] }],
  });
  charts.set("ppt/charts/chart2.xml", pie3D);
  const gf2 = graphicFrameXml(
    2,
    "3D Pie Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf2),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart2.xml" }]),
  });

  // Slide 3: line3DChart -> line fallback
  const line3D = chartXml("line3DChart", {
    title: "3D Line (fallback)",
    legendPos: "b",
    series: [
      {
        name: "Revenue",
        categories: ["Jan", "Feb", "Mar", "Apr", "May"],
        values: [100, 120, 90, 150, 130],
      },
    ],
  });
  charts.set("ppt/charts/chart3.xml", line3D);
  const gf3 = graphicFrameXml(
    2,
    "3D Line Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf3),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart3.xml" }]),
  });

  // Slide 4: area3DChart -> area fallback
  const area3D = chartXml("area3DChart", {
    title: "3D Area (fallback)",
    legendPos: "b",
    series: [
      {
        name: "Visitors",
        categories: ["Mon", "Tue", "Wed", "Thu", "Fri"],
        values: [200, 350, 280, 420, 380],
      },
      {
        name: "Page Views",
        categories: ["Mon", "Tue", "Wed", "Thu", "Fri"],
        values: [400, 500, 450, 600, 550],
      },
    ],
  });
  charts.set("ppt/charts/chart4.xml", area3D);
  const gf4 = graphicFrameXml(
    2,
    "3D Area Chart",
    margin,
    margin,
    SLIDE_W - margin * 2,
    SLIDE_H - margin * 2,
    "rId2",
  );
  slides.push({
    xml: wrapSlideXml(gf4),
    rels: slideRelsXml([{ id: "rId2", type: REL_TYPES.chart, target: "../charts/chart4.xml" }]),
  });

  const buffer = await buildPptx({
    slides,
    charts,
    contentTypesExtra: [
      `<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
      `<Override PartName="/ppt/charts/chart4.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
    ],
  });
  savePptx(buffer, "charts-3d-fallback.pptx");
}

// --- Color Transforms ---

export const chartFixtureCreators: FixtureCreatorMap = {
  "charts.pptx": createChartsFixture,
  "charts-3d-fallback.pptx": createCharts3dFallbackFixture,
};
