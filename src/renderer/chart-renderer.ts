import type { ChartElement, ChartData, ChartSeries } from "../model/chart.js";
import type { ResolvedColor } from "../model/fill.js";
import { emuToPixels } from "../utils/emu.js";
import { buildTransformAttr } from "./transform.js";

const DEFAULT_SERIES_COLORS: ResolvedColor[] = [
  { hex: "#4472C4", alpha: 1 },
  { hex: "#ED7D31", alpha: 1 },
  { hex: "#A5A5A5", alpha: 1 },
  { hex: "#FFC000", alpha: 1 },
  { hex: "#5B9BD5", alpha: 1 },
  { hex: "#70AD47", alpha: 1 },
];

export function renderChart(element: ChartElement): string {
  const { transform, chart } = element;
  const w = emuToPixels(transform.extentWidth);
  const h = emuToPixels(transform.extentHeight);
  const transformAttr = buildTransformAttr(transform);

  const parts: string[] = [];
  parts.push(`<g transform="${transformAttr}">`);
  parts.push(
    `<rect width="${w}" height="${h}" fill="#FFFFFF" stroke="#D9D9D9" stroke-width="0.5"/>`,
  );

  const margin = { top: 20, right: 20, bottom: 30, left: 50 };

  if (chart.title) {
    parts.push(renderChartTitle(chart.title, w));
    margin.top = 40;
  }

  if (chart.legend) {
    if (chart.legend.position === "b") margin.bottom = 50;
    else if (chart.legend.position === "t") margin.top += 20;
  }

  const plotX = margin.left;
  const plotY = margin.top;
  const plotW = Math.max(w - margin.left - margin.right, 0);
  const plotH = Math.max(h - margin.top - margin.bottom, 0);

  if (plotW > 0 && plotH > 0) {
    switch (chart.chartType) {
      case "bar":
        parts.push(renderBarChart(chart, plotX, plotY, plotW, plotH));
        break;
      case "line":
        parts.push(renderLineChart(chart, plotX, plotY, plotW, plotH));
        break;
      case "pie":
        parts.push(renderPieChart(chart, plotX, plotY, plotW, plotH));
        break;
      case "doughnut":
        parts.push(renderDoughnutChart(chart, plotX, plotY, plotW, plotH));
        break;
      case "scatter":
        parts.push(renderScatterChart(chart, plotX, plotY, plotW, plotH));
        break;
      case "bubble":
        parts.push(renderBubbleChart(chart, plotX, plotY, plotW, plotH));
        break;
      case "area":
        parts.push(renderAreaChart(chart, plotX, plotY, plotW, plotH));
        break;
    }
  }

  if (chart.legend && chart.series.length > 0) {
    parts.push(renderLegend(chart, w, h, chart.legend.position));
  }

  parts.push("</g>");
  return parts.join("");
}

function renderChartTitle(title: string, chartWidth: number): string {
  return `<text x="${round(chartWidth / 2)}" y="20" text-anchor="middle" font-size="14" font-weight="bold" fill="#404040">${escapeXml(title)}</text>`;
}

function renderBarChart(chart: ChartData, x: number, y: number, w: number, h: number): string {
  const parts: string[] = [];
  const { series, categories } = chart;
  if (series.length === 0) return "";

  const maxVal = getMaxValue(series);
  if (maxVal === 0) return "";

  const catCount = categories.length || Math.max(...series.map((s) => s.values.length));
  if (catCount === 0) return "";

  const isHorizontal = chart.barDirection === "bar";

  // Axes
  parts.push(
    `<line x1="${round(x)}" y1="${round(y + h)}" x2="${round(x + w)}" y2="${round(y + h)}" stroke="#D9D9D9" stroke-width="1"/>`,
  );
  parts.push(
    `<line x1="${round(x)}" y1="${round(y)}" x2="${round(x)}" y2="${round(y + h)}" stroke="#D9D9D9" stroke-width="1"/>`,
  );

  if (isHorizontal) {
    const groupHeight = h / catCount;
    const barHeight = (groupHeight * 0.7) / series.length;
    const groupPadding = groupHeight * 0.15;

    for (let c = 0; c < catCount; c++) {
      const label = categories[c] ?? "";
      const labelY = y + c * groupHeight + groupHeight / 2;
      parts.push(
        `<text x="${round(x - 5)}" y="${round(labelY + 4)}" text-anchor="end" font-size="10" fill="#595959">${escapeXml(label)}</text>`,
      );
    }

    for (let s = 0; s < series.length; s++) {
      const color = series[s].color;
      for (let c = 0; c < series[s].values.length; c++) {
        const val = series[s].values[c];
        const barW = (val / maxVal) * w;
        const barX = x;
        const barY = y + c * groupHeight + groupPadding + s * barHeight;
        parts.push(
          `<rect x="${round(barX)}" y="${round(barY)}" width="${round(barW)}" height="${round(barHeight)}" ${fillAttr(color)}/>`,
        );
      }
    }
  } else {
    const groupWidth = w / catCount;
    const barWidth = (groupWidth * 0.7) / series.length;
    const groupPadding = groupWidth * 0.15;

    for (let c = 0; c < catCount; c++) {
      const label = categories[c] ?? "";
      const labelX = x + c * groupWidth + groupWidth / 2;
      parts.push(
        `<text x="${round(labelX)}" y="${round(y + h + 15)}" text-anchor="middle" font-size="10" fill="#595959">${escapeXml(label)}</text>`,
      );
    }

    for (let s = 0; s < series.length; s++) {
      const color = series[s].color;
      for (let c = 0; c < series[s].values.length; c++) {
        const val = series[s].values[c];
        const barH = (val / maxVal) * h;
        const barX = x + c * groupWidth + groupPadding + s * barWidth;
        const barY = y + h - barH;
        parts.push(
          `<rect x="${round(barX)}" y="${round(barY)}" width="${round(barWidth)}" height="${round(barH)}" ${fillAttr(color)}/>`,
        );
      }
    }
  }

  return parts.join("");
}

function renderLineChart(chart: ChartData, x: number, y: number, w: number, h: number): string {
  const parts: string[] = [];
  const { series, categories } = chart;
  if (series.length === 0) return "";

  const maxVal = getMaxValue(series);
  if (maxVal === 0) return "";

  const catCount = categories.length || Math.max(...series.map((s) => s.values.length));
  if (catCount === 0) return "";

  // Axes
  parts.push(
    `<line x1="${round(x)}" y1="${round(y + h)}" x2="${round(x + w)}" y2="${round(y + h)}" stroke="#D9D9D9" stroke-width="1"/>`,
  );
  parts.push(
    `<line x1="${round(x)}" y1="${round(y)}" x2="${round(x)}" y2="${round(y + h)}" stroke="#D9D9D9" stroke-width="1"/>`,
  );

  // Category labels
  for (let c = 0; c < catCount; c++) {
    const label = categories[c] ?? "";
    const divisor = catCount > 1 ? catCount - 1 : 1;
    const labelX = x + (c / divisor) * w;
    parts.push(
      `<text x="${round(labelX)}" y="${round(y + h + 15)}" text-anchor="middle" font-size="10" fill="#595959">${escapeXml(label)}</text>`,
    );
  }

  // Lines
  for (let s = 0; s < series.length; s++) {
    const color = series[s].color;
    const divisor = catCount > 1 ? catCount - 1 : 1;
    const points = series[s].values.map((val, i) => {
      const px = round(x + (i / divisor) * w);
      const py = round(y + h - (val / maxVal) * h);
      return `${px},${py}`;
    });

    parts.push(
      `<polyline points="${points.join(" ")}" fill="none" stroke="${color.hex}" stroke-width="2"${color.alpha < 1 ? ` stroke-opacity="${color.alpha}"` : ""}/>`,
    );

    // Data point markers
    for (const point of points) {
      const [px, py] = point.split(",");
      parts.push(`<circle cx="${px}" cy="${py}" r="3" ${fillAttr(color)}/>`);
    }
  }

  return parts.join("");
}

function renderAreaChart(chart: ChartData, x: number, y: number, w: number, h: number): string {
  const parts: string[] = [];
  const { series, categories } = chart;
  if (series.length === 0) return "";

  const maxVal = getMaxValue(series);
  if (maxVal === 0) return "";

  const catCount = categories.length || Math.max(...series.map((s) => s.values.length));
  if (catCount === 0) return "";

  // Axes
  parts.push(
    `<line x1="${round(x)}" y1="${round(y + h)}" x2="${round(x + w)}" y2="${round(y + h)}" stroke="#D9D9D9" stroke-width="1"/>`,
  );
  parts.push(
    `<line x1="${round(x)}" y1="${round(y)}" x2="${round(x)}" y2="${round(y + h)}" stroke="#D9D9D9" stroke-width="1"/>`,
  );

  // Category labels
  for (let c = 0; c < catCount; c++) {
    const label = categories[c] ?? "";
    const divisor = catCount > 1 ? catCount - 1 : 1;
    const labelX = x + (c / divisor) * w;
    parts.push(
      `<text x="${round(labelX)}" y="${round(y + h + 15)}" text-anchor="middle" font-size="10" fill="#595959">${escapeXml(label)}</text>`,
    );
  }

  // Areas (filled polygons)
  const baseline = round(y + h);
  for (let s = 0; s < series.length; s++) {
    const color = series[s].color;
    const divisor = catCount > 1 ? catCount - 1 : 1;
    const dataPoints = series[s].values.map((val, i) => {
      const px = round(x + (i / divisor) * w);
      const py = round(y + h - (val / maxVal) * h);
      return { px, py };
    });

    // Build polygon points: data points + bottom-right + bottom-left
    const topPoints = dataPoints.map((p) => `${p.px},${p.py}`).join(" ");
    const lastX = dataPoints[dataPoints.length - 1].px;
    const firstX = dataPoints[0].px;
    const polygonPoints = `${topPoints} ${lastX},${baseline} ${firstX},${baseline}`;

    parts.push(
      `<polygon points="${polygonPoints}" fill="${color.hex}" fill-opacity="${color.alpha < 1 ? color.alpha : 0.5}" stroke="${color.hex}" stroke-width="2"${color.alpha < 1 ? ` stroke-opacity="${color.alpha}"` : ""}/>`,
    );
  }

  return parts.join("");
}

function renderPieChart(chart: ChartData, x: number, y: number, w: number, h: number): string {
  const parts: string[] = [];
  const series = chart.series[0];
  if (!series || series.values.length === 0) return "";

  const total = series.values.reduce((sum, v) => sum + v, 0);
  if (total === 0) return "";

  const cx = x + w / 2;
  const cy = y + h / 2;
  const r = (Math.min(w, h) / 2) * 0.85;

  let currentAngle = -Math.PI / 2;

  for (let i = 0; i < series.values.length; i++) {
    const val = series.values[i];
    const sliceAngle = (val / total) * 2 * Math.PI;
    const color = getPieSliceColor(i, chart);

    if (series.values.length === 1) {
      parts.push(
        `<circle cx="${round(cx)}" cy="${round(cy)}" r="${round(r)}" ${fillAttr(color)}/>`,
      );
    } else {
      const x1 = cx + r * Math.cos(currentAngle);
      const y1 = cy + r * Math.sin(currentAngle);
      const x2 = cx + r * Math.cos(currentAngle + sliceAngle);
      const y2 = cy + r * Math.sin(currentAngle + sliceAngle);
      const largeArc = sliceAngle > Math.PI ? 1 : 0;

      parts.push(
        `<path d="M${round(cx)},${round(cy)} L${round(x1)},${round(y1)} A${round(r)},${round(r)} 0 ${largeArc},1 ${round(x2)},${round(y2)} Z" ${fillAttr(color)}/>`,
      );
    }

    currentAngle += sliceAngle;
  }

  return parts.join("");
}

function renderDoughnutChart(chart: ChartData, x: number, y: number, w: number, h: number): string {
  const parts: string[] = [];
  const series = chart.series[0];
  if (!series || series.values.length === 0) return "";

  const total = series.values.reduce((sum, v) => sum + v, 0);
  if (total === 0) return "";

  const cx = x + w / 2;
  const cy = y + h / 2;
  const outerR = (Math.min(w, h) / 2) * 0.85;
  const holeSize = chart.holeSize ?? 50;
  const innerR = outerR * (holeSize / 100);

  let currentAngle = -Math.PI / 2;

  for (let i = 0; i < series.values.length; i++) {
    const val = series.values[i];
    const sliceAngle = (val / total) * 2 * Math.PI;
    const color = getPieSliceColor(i, chart);

    if (series.values.length === 1) {
      parts.push(
        `<circle cx="${round(cx)}" cy="${round(cy)}" r="${round(outerR)}" ${fillAttr(color)}/>`,
      );
      parts.push(
        `<circle cx="${round(cx)}" cy="${round(cy)}" r="${round(innerR)}" fill="#FFFFFF"/>`,
      );
    } else {
      const ox1 = cx + outerR * Math.cos(currentAngle);
      const oy1 = cy + outerR * Math.sin(currentAngle);
      const ox2 = cx + outerR * Math.cos(currentAngle + sliceAngle);
      const oy2 = cy + outerR * Math.sin(currentAngle + sliceAngle);
      const ix1 = cx + innerR * Math.cos(currentAngle + sliceAngle);
      const iy1 = cy + innerR * Math.sin(currentAngle + sliceAngle);
      const ix2 = cx + innerR * Math.cos(currentAngle);
      const iy2 = cy + innerR * Math.sin(currentAngle);
      const largeArc = sliceAngle > Math.PI ? 1 : 0;

      parts.push(
        `<path d="M${round(ox1)},${round(oy1)} A${round(outerR)},${round(outerR)} 0 ${largeArc},1 ${round(ox2)},${round(oy2)} L${round(ix1)},${round(iy1)} A${round(innerR)},${round(innerR)} 0 ${largeArc},0 ${round(ix2)},${round(iy2)} Z" ${fillAttr(color)}/>`,
      );
    }

    currentAngle += sliceAngle;
  }

  return parts.join("");
}

function renderScatterChart(chart: ChartData, x: number, y: number, w: number, h: number): string {
  const parts: string[] = [];
  const { series } = chart;
  if (series.length === 0) return "";

  let maxX = 0;
  let maxY = 0;
  for (const s of series) {
    const xVals = s.xValues ?? [];
    for (const v of xVals) maxX = Math.max(maxX, v);
    for (const v of s.values) maxY = Math.max(maxY, v);
  }
  if (maxX === 0) maxX = 1;
  if (maxY === 0) maxY = 1;

  // Axes
  parts.push(
    `<line x1="${round(x)}" y1="${round(y + h)}" x2="${round(x + w)}" y2="${round(y + h)}" stroke="#D9D9D9" stroke-width="1"/>`,
  );
  parts.push(
    `<line x1="${round(x)}" y1="${round(y)}" x2="${round(x)}" y2="${round(y + h)}" stroke="#D9D9D9" stroke-width="1"/>`,
  );

  // Points
  for (let s = 0; s < series.length; s++) {
    const color = series[s].color;
    const xVals = series[s].xValues ?? [];
    for (let i = 0; i < series[s].values.length; i++) {
      const xVal = xVals[i] ?? i;
      const yVal = series[s].values[i];
      const px = x + (xVal / maxX) * w;
      const py = y + h - (yVal / maxY) * h;
      parts.push(`<circle cx="${round(px)}" cy="${round(py)}" r="4" ${fillAttr(color)}/>`);
    }
  }

  return parts.join("");
}

function renderBubbleChart(chart: ChartData, x: number, y: number, w: number, h: number): string {
  const parts: string[] = [];
  const { series } = chart;
  if (series.length === 0) return "";

  let maxX = 0;
  let maxY = 0;
  let maxBubble = 0;
  for (const s of series) {
    const xVals = s.xValues ?? [];
    for (const v of xVals) maxX = Math.max(maxX, v);
    for (const v of s.values) maxY = Math.max(maxY, v);
    const sizes = s.bubbleSizes ?? [];
    for (const v of sizes) maxBubble = Math.max(maxBubble, v);
  }
  if (maxX === 0) maxX = 1;
  if (maxY === 0) maxY = 1;
  if (maxBubble === 0) maxBubble = 1;

  const maxRadius = Math.min(w, h) * 0.08;

  // Axes
  parts.push(
    `<line x1="${round(x)}" y1="${round(y + h)}" x2="${round(x + w)}" y2="${round(y + h)}" stroke="#D9D9D9" stroke-width="1"/>`,
  );
  parts.push(
    `<line x1="${round(x)}" y1="${round(y)}" x2="${round(x)}" y2="${round(y + h)}" stroke="#D9D9D9" stroke-width="1"/>`,
  );

  // Bubbles
  for (let s = 0; s < series.length; s++) {
    const color = series[s].color;
    const xVals = series[s].xValues ?? [];
    const sizes = series[s].bubbleSizes ?? [];
    for (let i = 0; i < series[s].values.length; i++) {
      const xVal = xVals[i] ?? i;
      const yVal = series[s].values[i];
      const size = sizes[i] ?? 1;
      const px = x + (xVal / maxX) * w;
      const py = y + h - (yVal / maxY) * h;
      const r = Math.max(2, Math.sqrt(size / maxBubble) * maxRadius);
      parts.push(
        `<circle cx="${round(px)}" cy="${round(py)}" r="${round(r)}" ${fillAttr(color)} fill-opacity="0.6"/>`,
      );
    }
  }

  return parts.join("");
}

function renderLegend(chart: ChartData, chartW: number, chartH: number, position: string): string {
  const parts: string[] = [];

  const entries =
    chart.chartType === "pie" || chart.chartType === "doughnut"
      ? chart.categories.map((cat, i) => ({
          label: cat,
          color: getPieSliceColor(i, chart),
        }))
      : chart.series.map((s, i) => ({
          label: s.name ?? `Series ${i + 1}`,
          color: s.color,
        }));

  if (entries.length === 0) return "";

  const entryWidth = 80;
  const totalWidth = entries.length * entryWidth;
  const startX = Math.max((chartW - totalWidth) / 2, 5);
  const legendY = position === "t" ? 25 : chartH - 15;

  for (let i = 0; i < entries.length; i++) {
    const ex = startX + i * entryWidth;
    parts.push(
      `<rect x="${round(ex)}" y="${round(legendY - 6)}" width="12" height="12" ${fillAttr(entries[i].color)}/>`,
    );
    parts.push(
      `<text x="${round(ex + 16)}" y="${round(legendY + 4)}" font-size="10" fill="#595959">${escapeXml(entries[i].label)}</text>`,
    );
  }

  return parts.join("");
}

function getPieSliceColor(index: number, chart: ChartData): ResolvedColor {
  // For pie charts, use a color per slice (not per series)
  const series = chart.series[0];
  if (series) {
    // If the series has explicit color, use default palette for slices
    return DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
  }
  return DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
}

function getMaxValue(series: ChartSeries[]): number {
  let max = 0;
  for (const s of series) {
    for (const v of s.values) {
      max = Math.max(max, v);
    }
  }
  return max;
}

function fillAttr(color: ResolvedColor): string {
  const alpha = color.alpha < 1 ? ` fill-opacity="${color.alpha}"` : "";
  return `fill="${color.hex}"${alpha}`;
}

function round(n: number): number {
  return Math.round(n * 100) / 100;
}

function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
