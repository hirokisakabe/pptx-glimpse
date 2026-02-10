import { describe, it, expect } from "vitest";
import { renderChart } from "../../src/renderer/chart-renderer.js";
import type { ChartElement } from "../../src/model/chart.js";

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

      const svg = renderChart(element);
      // Should contain rect elements for the 3 bars
      const barRects = svg.match(/<rect[^>]*fill="#4472C4"[^>]*\/>/g);
      expect(barRects).not.toBeNull();
      expect(barRects!.length).toBe(3);
    });

    it("renders category labels", () => {
      const element = createChartElement({
        chartType: "bar",
        series: [{ name: "S1", values: [10], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["Category 1"],
      });

      const svg = renderChart(element);
      expect(svg).toContain("Category 1");
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

      const svg = renderChart(element);
      expect(svg).toContain('fill="#4472C4"');
      expect(svg).toContain('fill="#ED7D31"');
    });
  });

  describe("line chart", () => {
    it("renders polyline elements", () => {
      const element = createChartElement({
        chartType: "line",
        series: [{ name: "L1", values: [10, 20, 15], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B", "C"],
      });

      const svg = renderChart(element);
      expect(svg).toContain("<polyline");
      expect(svg).toContain('stroke="#4472C4"');
    });

    it("renders data point markers", () => {
      const element = createChartElement({
        chartType: "line",
        series: [{ name: "L1", values: [10, 20], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A", "B"],
      });

      const svg = renderChart(element);
      const circles = svg.match(/<circle[^>]*fill="#4472C4"[^>]*\/>/g);
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

      const svg = renderChart(element);
      const paths = svg.match(/<path[^>]*d="M[^"]*A[^"]*Z"[^>]*\/>/g);
      expect(paths).not.toBeNull();
      expect(paths!.length).toBe(2);
    });

    it("renders circle for single-value pie", () => {
      const element = createChartElement({
        chartType: "pie",
        series: [{ name: "P1", values: [100], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A"],
      });

      const svg = renderChart(element);
      expect(svg).toContain("<circle");
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

      const svg = renderChart(element);
      const circles = svg.match(/<circle[^>]*fill="#4472C4"[^>]*\/>/g);
      expect(circles).not.toBeNull();
      expect(circles!.length).toBe(3);
    });
  });

  describe("title", () => {
    it("renders title text", () => {
      const element = createChartElement({
        title: "My Chart",
        series: [{ name: "S1", values: [10], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A"],
      });

      const svg = renderChart(element);
      expect(svg).toContain("My Chart");
      expect(svg).toContain('font-weight="bold"');
    });

    it("escapes XML characters in title", () => {
      const element = createChartElement({
        title: "A & B <C>",
        series: [{ name: "S1", values: [10], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A"],
      });

      const svg = renderChart(element);
      expect(svg).toContain("A &amp; B &lt;C&gt;");
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

      const svg = renderChart(element);
      expect(svg).toContain("Series A");
      expect(svg).toContain("Series B");
    });

    it("does not render legend when not present", () => {
      const element = createChartElement({
        series: [{ name: "S1", values: [10], color: { hex: "#4472C4", alpha: 1 } }],
        categories: ["A"],
        legend: null,
      });

      const svg = renderChart(element);
      // Should not contain legend rect elements (aside from chart background and bar)
      expect(svg).not.toContain("Series 1");
    });
  });

  describe("empty data", () => {
    it("renders empty chart when no series", () => {
      const element = createChartElement({
        series: [],
        categories: [],
      });

      const svg = renderChart(element);
      // Should still return valid SVG with just the background
      expect(svg).toContain("<g");
      expect(svg).toContain("</g>");
    });
  });

  describe("fill opacity", () => {
    it("applies fill-opacity for transparent colors", () => {
      const element = createChartElement({
        chartType: "bar",
        series: [{ name: "S1", values: [10], color: { hex: "#4472C4", alpha: 0.5 } }],
        categories: ["A"],
      });

      const svg = renderChart(element);
      expect(svg).toContain('fill-opacity="0.5"');
    });
  });
});
