import { describe, it, expect } from "vitest";
import { renderChart } from "./chart-renderer.js";
import type { ChartElement } from "../model/chart.js";

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
      // Extract radii — larger bubble size should produce a larger radius
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
});
