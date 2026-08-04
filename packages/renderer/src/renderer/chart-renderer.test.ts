import { describe, expect, it } from "vitest";

import type { ChartElement } from "../model/chart.js";
import { renderChart } from "./chart-renderer.js";

function createChartElement(overrides: Partial<ChartElement["chart"]>): ChartElement {
  return {
    type: "chart",
    transform: {
      offsetX: 914400,
      offsetY: 914400,
      extentWidth: 4572000,
      extentHeight: 2743200,
      rotation: 0,
      flipH: false,
      flipV: false,
    },
    chart: {
      chartType: "bar",
      title: null,
      series: [],
      categories: [],
      legend: null,
      ...overrides,
    },
  };
}

describe("renderChart", () => {
  describe("category combo chart", () => {
    it("renders both bar and line plot-group series", () => {
      const element = createChartElement({
        chartType: "combo",
        series: [
          {
            name: "Columns",
            values: [10, 20],
            color: { hex: "#4472C4", alpha: 1 },
            chartType: "bar",
          },
          {
            name: "Trend",
            values: [15, 25],
            color: { hex: "#ED7D31", alpha: 1 },
            chartType: "line",
          },
        ],
        categories: ["A", "B"],
        plotGroups: [
          {
            chartType: "bar",
            seriesIndexes: [0],
            axisIds: ["cat", "value"],
            valueAxisId: "value",
          },
          {
            chartType: "line",
            seriesIndexes: [1],
            axisIds: ["cat", "value"],
            valueAxisId: "value",
          },
        ],
        valueAxes: [{ id: "value", position: "l" }],
        legend: { position: "r" },
      });

      const result = renderChart(element);
      expect(result.content.match(/<rect[^>]*fill="#4472C4"[^>]*\/>/g)).toHaveLength(3);
      expect(result.content).toContain("<polyline");
      expect(result.content).toContain('stroke="#ED7D31"');
      expect(result.content).toContain("Columns");
      expect(result.content).toContain("Trend");
    });

    it("uses one shared scale and renders axes and categories once for a shared value axis", () => {
      const element = createChartElement({
        chartType: "combo",
        barDirection: "col",
        series: [
          {
            name: "Columns",
            values: [100, 80],
            color: { hex: "#4472C4", alpha: 1 },
            chartType: "bar",
          },
          {
            name: "Trend",
            values: [10, 20],
            color: { hex: "#ED7D31", alpha: 1 },
            chartType: "line",
          },
        ],
        categories: ["A", "B"],
        plotGroups: [
          {
            chartType: "bar",
            seriesIndexes: [0],
            axisIds: ["cat", "shared"],
            valueAxisId: "shared",
          },
          {
            chartType: "line",
            seriesIndexes: [1],
            axisIds: ["cat", "shared"],
            valueAxisId: "shared",
          },
        ],
        valueAxes: [{ id: "shared", position: "l" }],
      });

      const result = renderChart(element);
      expect(result.content.match(/>A<\/text>/g)).toHaveLength(1);
      expect(result.content.match(/>B<\/text>/g)).toHaveLength(1);
      expect(result.content).not.toContain('text-anchor="start"');
      expect(result.content).toContain(">100</text>");
      expect(result.content).toContain('points="152.5,234.2 357.5,210.4"');
      expect(result.content).toContain(
        '<line x1="50" y1="258" x2="460" y2="258" stroke="#D9D9D9" stroke-width="1"/>',
      );
    });

    it("uses a right-side secondary scale and paints plot groups in document order", () => {
      const element = createChartElement({
        chartType: "combo",
        barDirection: "col",
        series: [
          {
            name: "Trend",
            values: [500, 1000],
            color: { hex: "#ED7D31", alpha: 1 },
            chartType: "line",
          },
          {
            name: "Columns",
            values: [5, 10],
            color: { hex: "#4472C4", alpha: 1 },
            chartType: "bar",
          },
        ],
        categories: ["A", "B"],
        plotGroups: [
          {
            chartType: "line",
            seriesIndexes: [0],
            axisIds: ["cat", "secondary"],
            valueAxisId: "secondary",
          },
          {
            chartType: "bar",
            seriesIndexes: [1],
            axisIds: ["cat", "primary"],
            valueAxisId: "primary",
          },
        ],
        valueAxes: [
          { id: "primary", position: "l" },
          { id: "secondary", position: "r" },
        ],
      });

      const result = renderChart(element);
      expect(result.content).toMatch(/text-anchor="start"[^>]*>1000<\/text>/);
      expect(result.content.indexOf('stroke="#ED7D31"')).toBeLessThan(
        result.content.indexOf('fill="#4472C4"'),
      );
    });

    it("renders a zero-only shared axis instead of dropping the combo", () => {
      const element = createChartElement({
        chartType: "combo",
        series: [
          { name: "Columns", values: [0], color: { hex: "#4472C4", alpha: 1 } },
          { name: "Trend", values: [0], color: { hex: "#ED7D31", alpha: 1 } },
        ],
        categories: ["A"],
        plotGroups: [
          {
            chartType: "bar",
            seriesIndexes: [0],
            axisIds: ["cat", "shared"],
            valueAxisId: "shared",
          },
          {
            chartType: "line",
            seriesIndexes: [1],
            axisIds: ["cat", "shared"],
            valueAxisId: "shared",
          },
        ],
        valueAxes: [{ id: "shared", position: "l" }],
      });

      const result = renderChart(element);
      expect(result.content).toContain('fill="#4472C4"');
      expect(result.content).toContain('stroke="#ED7D31"');
      expect(result.content).toContain(">1</text>");
    });

    it("uses zero-inclusive negative domains on primary and secondary axes", () => {
      const element = createChartElement({
        chartType: "combo",
        series: [
          { name: "Columns", values: [-5, -10], color: { hex: "#4472C4", alpha: 1 } },
          { name: "Trend", values: [-50, -100], color: { hex: "#ED7D31", alpha: 1 } },
        ],
        categories: ["A", "B"],
        plotGroups: [
          {
            chartType: "bar",
            seriesIndexes: [0],
            axisIds: ["cat", "primary"],
            valueAxisId: "primary",
          },
          {
            chartType: "line",
            seriesIndexes: [1],
            axisIds: ["cat", "secondary"],
            valueAxisId: "secondary",
          },
        ],
        valueAxes: [
          { id: "primary", position: "l" },
          { id: "secondary", position: "r" },
        ],
      });

      const result = renderChart(element);
      expect(result.content).toContain(">-10</text>");
      expect(result.content).toMatch(/text-anchor="start"[^>]*>-100<\/text>/);
      expect(result.content).toContain('stroke="#ED7D31"');
      expect(result.content).toMatch(/<rect[^>]*height="[1-9][^"]*"[^>]*fill="#4472C4"/);
      expect(result.content).toContain(
        '<line x1="50" y1="20" x2="460" y2="20" stroke="#D9D9D9" stroke-width="1"/>',
      );
    });

    it("draws the category baseline at primary-axis zero for a mixed-sign domain", () => {
      const element = createChartElement({
        chartType: "combo",
        series: [
          { name: "Trend", values: [100, 200], color: { hex: "#ED7D31", alpha: 1 } },
          { name: "Columns", values: [-10, 10], color: { hex: "#4472C4", alpha: 1 } },
        ],
        categories: ["A", "B"],
        plotGroups: [
          {
            chartType: "line",
            seriesIndexes: [0],
            axisIds: ["cat", "secondary"],
            valueAxisId: "secondary",
          },
          {
            chartType: "bar",
            seriesIndexes: [1],
            axisIds: ["cat", "primary"],
            valueAxisId: "primary",
          },
        ],
        valueAxes: [
          { id: "primary", position: "l" },
          { id: "secondary", position: "r" },
        ],
      });

      const result = renderChart(element);
      expect(result.content).toContain(
        '<line x1="50" y1="139" x2="460" y2="139" stroke="#D9D9D9" stroke-width="1"/>',
      );
    });

    it("does not render a horizontal bar and line combo", () => {
      const element = createChartElement({
        chartType: "combo",
        barDirection: "bar",
        series: [
          {
            name: "Columns",
            values: [10],
            color: { hex: "#4472C4", alpha: 1 },
            chartType: "bar",
          },
          {
            name: "Trend",
            values: [20],
            color: { hex: "#ED7D31", alpha: 1 },
            chartType: "line",
          },
        ],
        categories: ["A"],
      });

      const result = renderChart(element);
      expect(result.content).not.toContain('fill="#4472C4"');
      expect(result.content).not.toContain('stroke="#ED7D31"');
    });
  });

  describe("bar chart", () => {
    it("renders rect elements for bars", () => {
      const element = createChartElement({
        chartType: "bar",
        series: [{ name: "S1", values: [10, 20, 30], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B", "C"],
      });

      const result = renderChart(element);
      // Should contain rect elements for the 3 bars
      const barRects = result.content.match(/<rect[^>]*fill="#4472C4"[^>]*\/>/g);
      expect(barRects).not.toBeNull();
      expect(barRects!.length).toBe(3);
    });

    it("renders category labels", () => {
      const element = createChartElement({
        chartType: "bar",
        series: [{ name: "S1", values: [10], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["Category 1"],
      });

      const result = renderChart(element);
      expect(result.content).toContain("Category 1");
    });

    it("renders multiple series", () => {
      const element = createChartElement({
        chartType: "bar",
        series: [
          { name: "S1", values: [10], color: { hex: "#4472C4", alpha: 1 } },
          { name: "S2", values: [20], color: { hex: "#ED7D31", alpha: 1 } },
        ],
        categories: ["A"],
      });

      const result = renderChart(element);
      expect(result.content).toContain('fill="#4472C4"');
      expect(result.content).toContain('fill="#ED7D31"');
    });
  });

  describe("line chart", () => {
    it("renders polyline elements", () => {
      const element = createChartElement({
        chartType: "line",
        series: [{ name: "L1", values: [10, 20, 15], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B", "C"],
      });

      const result = renderChart(element);
      expect(result.content).toContain("<polyline");
      expect(result.content).toContain('stroke="#4472C4"');
    });

    it("renders data point markers", () => {
      const element = createChartElement({
        chartType: "line",
        series: [{ name: "L1", values: [10, 20], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B"],
      });

      const result = renderChart(element);
      const circles = result.content.match(/<circle[^>]*fill="#4472C4"[^>]*\/>/g);
      expect(circles).not.toBeNull();
      expect(circles!.length).toBe(2);
    });
  });

  describe("pie chart", () => {
    it("renders path elements for slices", () => {
      const element = createChartElement({
        chartType: "pie",
        series: [{ name: "P1", values: [60, 40], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B"],
      });

      const result = renderChart(element);
      const paths = result.content.match(/<path[^>]*d="M[^"]*A[^"]*Z"[^>]*\/>/g);
      expect(paths).not.toBeNull();
      expect(paths!.length).toBe(2);
    });

    it("renders circle for single-value pie", () => {
      const element = createChartElement({
        chartType: "pie",
        series: [{ name: "P1", values: [100], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A"],
      });

      const result = renderChart(element);
      expect(result.content).toContain("<circle");
    });
  });

  describe("doughnut chart", () => {
    it("renders path elements with inner and outer arcs", () => {
      const element = createChartElement({
        chartType: "doughnut",
        holeSize: 50,
        series: [{ name: "D1", values: [60, 40], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B"],
      });

      const result = renderChart(element);
      const paths = result.content.match(/<path[^>]*d="M[^"]*A[^"]*A[^"]*Z"[^>]*\/>/g);
      expect(paths).not.toBeNull();
      expect(paths!.length).toBe(2);
    });

    it("renders circles for single-value doughnut", () => {
      const element = createChartElement({
        chartType: "doughnut",
        holeSize: 50,
        series: [{ name: "D1", values: [100], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A"],
      });

      const result = renderChart(element);
      const circles = result.content.match(/<circle[^>]*\/>/g);
      expect(circles).not.toBeNull();
      expect(circles!.length).toBeGreaterThanOrEqual(2);
    });

    it("uses default holeSize of 50 when not specified", () => {
      const element = createChartElement({
        chartType: "doughnut",
        series: [{ name: "D1", values: [60, 40], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B"],
      });

      const result = renderChart(element);
      expect(result.content).toContain("<path");
    });
  });

  describe("scatter chart", () => {
    it("renders circle elements for data points", () => {
      const element = createChartElement({
        chartType: "scatter",
        series: [
          {
            name: "Scatter1",
            values: [10, 20, 15],
            xValues: [1, 2, 3],
            color: { hex: "#4472C4", alpha: 1 },
          },
        ],
        categories: [],
      });

      const result = renderChart(element);
      const circles = result.content.match(/<circle[^>]*fill="#4472C4"[^>]*\/>/g);
      expect(circles).not.toBeNull();
      expect(circles!.length).toBe(3);
    });
  });

  describe("bubble chart", () => {
    it("renders circle elements for data points", () => {
      const element = createChartElement({
        chartType: "bubble",
        series: [
          {
            name: "Bubble1",
            values: [10, 20, 15],
            xValues: [1, 2, 3],
            bubbleSizes: [5, 10, 8],
            color: { hex: "#4472C4", alpha: 1 },
          },
        ],
        categories: [],
      });

      const result = renderChart(element);
      const circles = result.content.match(/<circle[^>]*fill="#4472C4"[^>]*\/>/g);
      expect(circles).not.toBeNull();
      expect(circles!.length).toBe(3);
    });

    it("renders circles with varying radii based on bubble sizes", () => {
      const element = createChartElement({
        chartType: "bubble",
        series: [
          {
            name: "Bubble1",
            values: [10, 20],
            xValues: [1, 2],
            bubbleSizes: [4, 16],
            color: { hex: "#4472C4", alpha: 1 },
          },
        ],
        categories: [],
      });

      const result = renderChart(element);
      const circles = result.content.match(/<circle[^>]*r="([^"]*)"[^>]*\/>/g);
      expect(circles).not.toBeNull();
      expect(circles!.length).toBe(2);
      // Extract radii - larger bubble size should produce a larger radius
      const radii = circles!.map((c) => {
        const match = c.match(/r="([^"]*)"/);
        return Number(match![1]);
      });
      expect(radii[1]).toBeGreaterThan(radii[0]);
    });

    it("renders bubbles with fill-opacity", () => {
      const element = createChartElement({
        chartType: "bubble",
        series: [
          {
            name: "Bubble1",
            values: [10],
            xValues: [1],
            bubbleSizes: [5],
            color: { hex: "#4472C4", alpha: 1 },
          },
        ],
        categories: [],
      });

      const result = renderChart(element);
      expect(result.content).toContain('fill-opacity="0.6"');
    });
  });

  describe("radar chart", () => {
    it("renders grid circles and data polygon", () => {
      const element = createChartElement({
        chartType: "radar",
        radarStyle: "standard",
        series: [{ name: "S1", values: [8, 6, 9, 7], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B", "C", "D"],
      });

      const result = renderChart(element);
      // Grid circles (5 levels)
      const circles = result.content.match(/<circle[^>]*stroke="#D9D9D9"[^>]*\/>/g);
      expect(circles).not.toBeNull();
      expect(circles!.length).toBe(5);
      // Data polygon
      expect(result.content).toContain("<polygon");
      expect(result.content).toContain('fill="none"');
      expect(result.content).toContain('stroke="#4472C4"');
    });

    it("renders filled polygon for filled radarStyle", () => {
      const element = createChartElement({
        chartType: "radar",
        radarStyle: "filled",
        series: [{ name: "S1", values: [5, 3, 7], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B", "C"],
      });

      const result = renderChart(element);
      expect(result.content).toContain('fill="#4472C4"');
      expect(result.content).toContain('fill-opacity="0.3"');
    });

    it("renders markers for marker radarStyle", () => {
      const element = createChartElement({
        chartType: "radar",
        radarStyle: "marker",
        series: [{ name: "S1", values: [5, 3, 7], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B", "C"],
      });

      const result = renderChart(element);
      // 3 data point markers
      const markers = result.content.match(/<circle[^>]*r="3"[^>]*\/>/g);
      expect(markers).not.toBeNull();
      expect(markers!.length).toBe(3);
    });

    it("renders category labels", () => {
      const element = createChartElement({
        chartType: "radar",
        radarStyle: "standard",
        series: [{ name: "S1", values: [5, 3], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["Speed", "Power"],
      });

      const result = renderChart(element);
      expect(result.content).toContain("Speed");
      expect(result.content).toContain("Power");
    });
  });

  describe("title", () => {
    it("renders title text", () => {
      const element = createChartElement({
        title: "My Chart",
        series: [{ name: "S1", values: [10], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A"],
      });

      const result = renderChart(element);
      expect(result.content).toContain("My Chart");
      expect(result.content).toContain('font-weight="bold"');
    });

    it("escapes XML characters in title", () => {
      const element = createChartElement({
        title: "A & B <C>",
        series: [{ name: "S1", values: [10], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A"],
      });

      const result = renderChart(element);
      expect(result.content).toContain("A &amp; B &lt;C&gt;");
    });
  });

  describe("legend", () => {
    it("renders legend entries", () => {
      const element = createChartElement({
        series: [
          { name: "Series A", values: [10], color: { hex: "#4472C4", alpha: 1 } },
          { name: "Series B", values: [20], color: { hex: "#ED7D31", alpha: 1 } },
        ],
        categories: ["A"],
        legend: { position: "b" },
      });

      const result = renderChart(element);
      expect(result.content).toContain("Series A");
      expect(result.content).toContain("Series B");
    });

    it("does not render legend when not present", () => {
      const element = createChartElement({
        series: [{ name: "S1", values: [10], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A"],
        legend: null,
      });

      const result = renderChart(element);
      // Should not contain legend rect elements (aside from chart background and bar)
      expect(result.content).not.toContain("Series 1");
    });

    it("renders right legend vertically on right side and reserves right margin", () => {
      const element = createChartElement({
        chartType: "bar",
        series: [
          { name: "Alpha", values: [10, 20], color: { hex: "#4472C4", alpha: 1 } },
          { name: "Beta", values: [15, 25], color: { hex: "#ED7D31", alpha: 1 } },
        ],
        categories: ["A", "B"],
        legend: { position: "r" },
      });

      const result = renderChart(element);
      expect(result.content).toContain("Alpha");
      expect(result.content).toContain("Beta");

      // Chart: 4572000 EMU -> 480px; margin.right=100 -> plotW=330, x-axis x2=380
      // x-axis line x2 must be 380 (plot right edge), not 480 (chart right edge)
      expect(result.content).toContain('x2="380"');
      // Legend rect starts at chartW - 100 + 5 = 385
      expect(result.content).toContain('x="385"');
    });

    it("renders left legend vertically on left side and shifts plot area right", () => {
      const element = createChartElement({
        chartType: "line",
        series: [{ name: "Gamma", values: [5, 10], color: { hex: "#A5A5A5", alpha: 1 } }],
        categories: ["A", "B"],
        legend: { position: "l" },
      });

      const result = renderChart(element);
      expect(result.content).toContain("Gamma");

      // margin.left = 50 + 100 = 150; x-axis x1 must be 150 (plot left edge)
      expect(result.content).toContain('x1="150"');
      // Legend rect starts at x=5
      expect(result.content).toContain('x="5"');
    });
  });

  describe("empty data", () => {
    it("renders empty chart when no series", () => {
      const element = createChartElement({
        series: [],
        categories: [],
      });

      const result = renderChart(element);
      // Should still return valid SVG with just the background
      expect(result.content).toContain("<g");
      expect(result.content).toContain("</g>");
    });
  });

  describe("stock chart", () => {
    it("renders hi-lo vertical lines and close tick marks", () => {
      const element = createChartElement({
        chartType: "stock",
        series: [
          { name: "High", values: [150, 160], color: { hex: "#4472C4", alpha: 1 } },
          { name: "Low", values: [100, 110], color: { hex: "#ED7D31", alpha: 1 } },
          { name: "Close", values: [130, 140], color: { hex: "#A5A5A5", alpha: 1 } },
        ],
        categories: ["Day 1", "Day 2"],
      });

      const result = renderChart(element);
      // Hi-lo vertical lines (2 categories)
      const hiLoLines = result.content.match(
        /<line[^>]*stroke="#404040"[^>]*stroke-width="2"[^>]*\/>/g,
      );
      expect(hiLoLines).not.toBeNull();
      // 2 hi-lo lines + 2 close tick marks = 4 lines
      expect(hiLoLines!.length).toBe(4);
    });

    it("renders category labels", () => {
      const element = createChartElement({
        chartType: "stock",
        series: [
          { name: "High", values: [150], color: { hex: "#4472C4", alpha: 1 } },
          { name: "Low", values: [100], color: { hex: "#ED7D31", alpha: 1 } },
          { name: "Close", values: [130], color: { hex: "#A5A5A5", alpha: 1 } },
        ],
        categories: ["Day 1"],
      });

      const result = renderChart(element);
      expect(result.content).toContain("Day 1");
    });

    it("returns empty string when fewer than 3 series", () => {
      const element = createChartElement({
        chartType: "stock",
        series: [
          { name: "High", values: [150], color: { hex: "#4472C4", alpha: 1 } },
          { name: "Low", values: [100], color: { hex: "#ED7D31", alpha: 1 } },
        ],
        categories: ["Day 1"],
      });

      const result = renderChart(element);
      // Should have no hi-lo lines since there are not enough series
      expect(result.content).not.toContain('stroke="#404040"');
    });
  });

  describe("surface chart", () => {
    it("renders heatmap cells as rect elements", () => {
      const element = createChartElement({
        chartType: "surface",
        series: [
          { name: "Row 1", values: [10, 20, 30], color: { hex: "#4472C4", alpha: 1 } },
          { name: "Row 2", values: [15, 25, 35], color: { hex: "#ED7D31", alpha: 1 } },
        ],
        categories: ["A", "B", "C"],
      });

      const result = renderChart(element);
      // 2 rows × 3 cols = 6 heatmap cells + 1 background rect = 7 total rects
      const rects = result.content.match(/<rect[^>]*\/>/g);
      expect(rects).not.toBeNull();
      expect(rects!.length).toBe(7); // 6 cells + 1 background
    });

    it("renders category labels", () => {
      const element = createChartElement({
        chartType: "surface",
        series: [{ name: "Row 1", values: [10, 20], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["Col A", "Col B"],
      });

      const result = renderChart(element);
      expect(result.content).toContain("Col A");
      expect(result.content).toContain("Col B");
    });

    it("renders series labels", () => {
      const element = createChartElement({
        chartType: "surface",
        series: [
          { name: "Row 1", values: [10], color: { hex: "#4472C4", alpha: 1 } },
          { name: "Row 2", values: [20], color: { hex: "#ED7D31", alpha: 1 } },
        ],
        categories: ["A"],
      });

      const result = renderChart(element);
      expect(result.content).toContain("Row 1");
      expect(result.content).toContain("Row 2");
    });
  });

  describe("ofPie chart", () => {
    it("renders main pie slices and secondary pie", () => {
      const element = createChartElement({
        chartType: "ofPie",
        ofPieType: "pie",
        splitPos: 2,
        secondPieSize: 75,
        series: [
          {
            name: "Data",
            values: [40, 30, 20, 10],
            color: { hex: "#4472C4", alpha: 1 },
          },
        ],
        categories: ["A", "B", "C", "D"],
      });

      const result = renderChart(element);
      // Should contain path elements for pie slices
      const paths = result.content.match(/<path[^>]*d="M[^"]*"[^>]*\/>/g);
      expect(paths).not.toBeNull();
      expect(paths!.length).toBeGreaterThanOrEqual(3);
      // Should contain connection lines
      const lines = result.content.match(/<line[^>]*stroke="#A6A6A6"[^>]*\/>/g);
      expect(lines).not.toBeNull();
      expect(lines!.length).toBe(2);
    });

    it("renders bar-of-pie with rect elements for secondary chart", () => {
      const element = createChartElement({
        chartType: "ofPie",
        ofPieType: "bar",
        splitPos: 2,
        secondPieSize: 75,
        series: [
          {
            name: "Data",
            values: [50, 30, 15, 5],
            color: { hex: "#4472C4", alpha: 1 },
          },
        ],
        categories: ["A", "B", "C", "D"],
      });

      const result = renderChart(element);
      // Background rect + 2 bar segments = 3 rects
      const rects = result.content.match(/<rect[^>]*\/>/g);
      expect(rects).not.toBeNull();
      expect(rects!.length).toBe(3); // 1 background + 2 bar segments
    });
  });

  describe("fill opacity", () => {
    it("applies fill-opacity for transparent colors", () => {
      const element = createChartElement({
        chartType: "bar",
        series: [{ name: "S1", values: [10], color: { hex: "#4472C4", alpha: 0.5 } }],
        categories: ["A"],
      });

      const result = renderChart(element);
      expect(result.content).toContain('fill-opacity="0.5"');
    });
  });

  describe("value axis labels", () => {
    it("renders Y-axis numeric labels for bar chart", () => {
      const element = createChartElement({
        chartType: "bar",
        series: [{ name: "S1", values: [10, 20, 30], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B", "C"],
      });

      const result = renderChart(element);
      // Should contain numeric tick labels on the Y-axis
      const tickLabels = result.content.match(/text-anchor="end"[^>]*>\d+</g);
      expect(tickLabels).not.toBeNull();
      expect(tickLabels!.length).toBeGreaterThanOrEqual(2);
    });

    it("renders Y-axis numeric labels for line chart", () => {
      const element = createChartElement({
        chartType: "line",
        series: [{ name: "L1", values: [25, 50, 75], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B", "C"],
      });

      const result = renderChart(element);
      const tickLabels = result.content.match(/text-anchor="end"[^>]*>\d+</g);
      expect(tickLabels).not.toBeNull();
      expect(tickLabels!.length).toBeGreaterThanOrEqual(2);
    });

    it("uses nice scale max (>= data max) for bar chart", () => {
      const element = createChartElement({
        chartType: "bar",
        series: [{ name: "S1", values: [7, 13, 18], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B", "C"],
      });

      const result = renderChart(element);
      // The highest tick should be >= 18 (the data max), e.g. 20
      expect(result.content).toContain(">20<");
    });
  });
});
