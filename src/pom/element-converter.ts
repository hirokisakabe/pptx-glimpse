import type { ChartElement } from "../model/chart.js";
import type { ImageElement } from "../model/image.js";
import type { Outline } from "../model/line.js";
import type { ConnectorElement, GroupElement, ShapeElement, SlideElement } from "../model/shape.js";
import type { TableElement } from "../model/table.js";
import type { SlideScaleContext } from "./pom-converter.js";
import { stripHash } from "./pom-converter.js";
import type {
  PomChartData,
  PomChartNode,
  PomChartType,
  PomDashType,
  PomImageNode,
  PomLayerChild,
  PomLayerNode,
  PomLineArrow,
  PomLineNode,
  PomNode,
  PomShapeNode,
  PomTableCell,
  PomTableNode,
} from "./pom-types.js";
import { convertTextBodyToNodes } from "./text-converter.js";

export function convertElement(element: SlideElement, ctx: SlideScaleContext): PomNode | null {
  switch (element.type) {
    case "shape":
      return convertShape(element, ctx);
    case "image":
      return convertImage(element);
    case "table":
      return convertTable(element, ctx);
    case "chart":
      return convertChart(element);
    case "connector":
      return convertConnector(element, ctx);
    case "group":
      return convertGroup(element, ctx);
  }
}

// ---------------------------------------------------------------------------
// Shape
// ---------------------------------------------------------------------------

const POM_SUPPORTED_SHAPES = new Set([
  "accentBorderCallout1",
  "accentBorderCallout2",
  "accentBorderCallout3",
  "accentCallout1",
  "accentCallout2",
  "accentCallout3",
  "actionButtonBackPrevious",
  "actionButtonBeginning",
  "actionButtonBlank",
  "actionButtonDocument",
  "actionButtonEnd",
  "actionButtonForwardNext",
  "actionButtonHelp",
  "actionButtonHome",
  "actionButtonInformation",
  "actionButtonMovie",
  "actionButtonReturn",
  "actionButtonSound",
  "arc",
  "bentArrow",
  "bentUpArrow",
  "bevel",
  "blockArc",
  "borderCallout1",
  "borderCallout2",
  "borderCallout3",
  "bracePair",
  "bracketPair",
  "callout1",
  "callout2",
  "callout3",
  "can",
  "chartPlus",
  "chartStar",
  "chartX",
  "chevron",
  "chord",
  "circularArrow",
  "cloud",
  "cloudCallout",
  "corner",
  "cornerTabs",
  "cube",
  "curvedDownArrow",
  "curvedLeftArrow",
  "curvedRightArrow",
  "curvedUpArrow",
  "decagon",
  "diagStripe",
  "diamond",
  "dodecagon",
  "donut",
  "doubleWave",
  "downArrow",
  "downArrowCallout",
  "ellipse",
  "ellipseRibbon",
  "ellipseRibbon2",
  "flowChartAlternateProcess",
  "flowChartCollate",
  "flowChartConnector",
  "flowChartDecision",
  "flowChartDelay",
  "flowChartDisplay",
  "flowChartDocument",
  "flowChartExtract",
  "flowChartInputOutput",
  "flowChartInternalStorage",
  "flowChartMagneticDisk",
  "flowChartMagneticDrum",
  "flowChartMagneticTape",
  "flowChartManualInput",
  "flowChartManualOperation",
  "flowChartMerge",
  "flowChartMultidocument",
  "flowChartOfflineStorage",
  "flowChartOffpageConnector",
  "flowChartOnlineStorage",
  "flowChartOr",
  "flowChartPredefinedProcess",
  "flowChartPreparation",
  "flowChartProcess",
  "flowChartPunchedCard",
  "flowChartPunchedTape",
  "flowChartSort",
  "flowChartSummingJunction",
  "flowChartTerminator",
  "folderCorner",
  "frame",
  "funnel",
  "gear6",
  "gear9",
  "halfFrame",
  "heart",
  "heptagon",
  "hexagon",
  "homePlate",
  "horizontalScroll",
  "irregularSeal1",
  "irregularSeal2",
  "leftArrow",
  "leftArrowCallout",
  "leftBrace",
  "leftBracket",
  "leftCircularArrow",
  "leftRightArrow",
  "leftRightArrowCallout",
  "leftRightCircularArrow",
  "leftRightRibbon",
  "leftRightUpArrow",
  "leftUpArrow",
  "lightningBolt",
  "line",
  "lineInv",
  "mathDivide",
  "mathEqual",
  "mathMinus",
  "mathMultiply",
  "mathNotEqual",
  "mathPlus",
  "moon",
  "noSmoking",
  "nonIsoscelesTrapezoid",
  "notchedRightArrow",
  "octagon",
  "parallelogram",
  "pentagon",
  "pie",
  "pieWedge",
  "plaque",
  "plaqueTabs",
  "plus",
  "quadArrow",
  "quadArrowCallout",
  "rect",
  "ribbon",
  "ribbon2",
  "rightArrow",
  "rightArrowCallout",
  "rightBrace",
  "rightBracket",
  "round1Rect",
  "round2DiagRect",
  "round2SameRect",
  "roundRect",
  "rtTriangle",
  "smileyFace",
  "snip1Rect",
  "snip2DiagRect",
  "snip2SameRect",
  "snipRoundRect",
  "squareTabs",
  "star10",
  "star12",
  "star16",
  "star24",
  "star32",
  "star4",
  "star5",
  "star6",
  "star7",
  "star8",
  "stripedRightArrow",
  "sun",
  "swooshArrow",
  "teardrop",
  "trapezoid",
  "triangle",
  "upArrow",
  "upArrowCallout",
  "upDownArrow",
  "upDownArrowCallout",
  "uturnArrow",
  "verticalScroll",
  "wave",
  "wedgeEllipseCallout",
  "wedgeRectCallout",
  "wedgeRoundRectCallout",
]);

function convertShape(shape: ShapeElement, ctx: SlideScaleContext): PomNode | null {
  // If the shape has text but no recognized geometry, convert as text node(s)
  if (shape.textBody && shape.textBody.paragraphs.length > 0) {
    const hasVisualShape = shape.geometry.type === "preset" && shape.geometry.preset !== "rect";
    const hasFill = shape.fill !== null && shape.fill.type !== "none";
    const hasOutline = shape.outline !== null && (shape.outline.width as number) > 0;

    // Text-only shape (plain rect with no fill/outline) → convert as text/list nodes
    if (!hasVisualShape && !hasFill && !hasOutline) {
      return convertTextBodyToNodes(shape.textBody, ctx);
    }
  }

  // Convert as pom Shape node
  const preset = shape.geometry.type === "preset" ? shape.geometry.preset : "rect";
  const shapeType = POM_SUPPORTED_SHAPES.has(preset) ? preset : "rect";

  const node: PomShapeNode = {
    type: "shape",
    shapeType,
  };

  // Fill
  if (shape.fill && shape.fill.type === "solid") {
    node.fill = {
      color: stripHash(shape.fill.color.hex),
    };
    if (shape.fill.color.alpha < 1) {
      node.fill.transparency = Math.round((1 - shape.fill.color.alpha) * 100);
    }
  }

  // Outline → line
  if (shape.outline && (shape.outline.width as number) > 0) {
    node.line = convertOutlineToLine(shape.outline, ctx);
  }

  // Shadow
  if (shape.effects?.outerShadow) {
    const s = shape.effects.outerShadow;
    node.shadow = {
      type: "outer",
      color: stripHash(s.color.hex),
      opacity: s.color.alpha,
      blur: round((s.blurRadius as number) * ctx.scaleX),
      offset: round((s.distance as number) * ctx.scaleX),
      angle: s.direction,
    };
  }

  // Text inside shape
  if (shape.textBody && shape.textBody.paragraphs.length > 0) {
    const textParts: string[] = [];
    let firstFontPx: number | undefined;
    let firstColor: string | undefined;
    let firstBold: boolean | undefined;
    let firstItalic: boolean | undefined;
    let firstAlign: "left" | "center" | "right" | undefined;

    for (const para of shape.textBody.paragraphs) {
      const paraText = para.runs.map((r) => r.text).join("");
      if (paraText) textParts.push(paraText);

      if (firstFontPx === undefined) {
        for (const run of para.runs) {
          if (run.properties.fontSize !== null) {
            firstFontPx = ptToPx(run.properties.fontSize as number);
            break;
          }
        }
      }
      if (firstColor === undefined) {
        for (const run of para.runs) {
          if (run.properties.color) {
            firstColor = stripHash(run.properties.color.hex);
            break;
          }
        }
      }
      if (firstBold === undefined) {
        for (const run of para.runs) {
          if (run.properties.bold) {
            firstBold = true;
            break;
          }
        }
      }
      if (firstItalic === undefined) {
        for (const run of para.runs) {
          if (run.properties.italic) {
            firstItalic = true;
            break;
          }
        }
      }
      if (firstAlign === undefined) {
        firstAlign = convertAlignment(para.properties.alignment);
      }
    }

    if (textParts.length > 0) {
      node.text = textParts.join("\n");
    }
    if (firstFontPx !== undefined) node.fontPx = firstFontPx;
    if (firstColor !== undefined) node.color = firstColor;
    if (firstBold) node.bold = true;
    if (firstItalic) node.italic = true;
    if (firstAlign && firstAlign !== "left") node.alignText = firstAlign;
  }

  return node;
}

// ---------------------------------------------------------------------------
// Image
// ---------------------------------------------------------------------------

function convertImage(image: ImageElement): PomImageNode {
  const src = `data:${image.mimeType};base64,${image.imageData}`;

  const node: PomImageNode = {
    type: "image",
    src,
  };

  // Crop
  if (image.srcRect) {
    const { left, top, right, bottom } = image.srcRect;
    if (left > 0 || top > 0 || right > 0 || bottom > 0) {
      node.sizing = {
        type: "crop",
        x: left,
        y: top,
      };
    }
  }

  return node;
}

// ---------------------------------------------------------------------------
// Table
// ---------------------------------------------------------------------------

function convertTable(table: TableElement, ctx: SlideScaleContext): PomTableNode {
  const columns = table.table.columns.map((col) => ({
    width: round((col.width as number) * ctx.scaleX),
  }));

  const rows = table.table.rows.map((row) => {
    const cells: PomTableCell[] = row.cells
      .filter((cell) => !cell.hMerge && !cell.vMerge)
      .map((cell) => {
        const textParts: string[] = [];
        let fontPx: number | undefined;
        let color: string | undefined;
        let bold: boolean | undefined;
        let italic: boolean | undefined;
        let alignText: "left" | "center" | "right" | undefined;

        if (cell.textBody) {
          for (const para of cell.textBody.paragraphs) {
            const paraText = para.runs.map((r) => r.text).join("");
            if (paraText) textParts.push(paraText);

            if (fontPx === undefined) {
              for (const run of para.runs) {
                if (run.properties.fontSize !== null) {
                  fontPx = ptToPx(run.properties.fontSize as number);
                  break;
                }
              }
            }
            if (color === undefined) {
              for (const run of para.runs) {
                if (run.properties.color) {
                  color = stripHash(run.properties.color.hex);
                  break;
                }
              }
            }
            if (bold === undefined && para.runs.some((r) => r.properties.bold)) {
              bold = true;
            }
            if (italic === undefined && para.runs.some((r) => r.properties.italic)) {
              italic = true;
            }
            if (alignText === undefined) {
              alignText = convertAlignment(para.properties.alignment);
            }
          }
        }

        const pomCell: PomTableCell = {
          text: textParts.join("\n"),
        };
        if (fontPx !== undefined) pomCell.fontPx = fontPx;
        if (color !== undefined) pomCell.color = color;
        if (bold) pomCell.bold = true;
        if (italic) pomCell.italic = true;
        if (alignText && alignText !== "left") pomCell.alignText = alignText;
        if (cell.fill && cell.fill.type === "solid") {
          pomCell.backgroundColor = stripHash(cell.fill.color.hex);
        }
        if (cell.gridSpan > 1) pomCell.colspan = cell.gridSpan;
        if (cell.rowSpan > 1) pomCell.rowspan = cell.rowSpan;

        return pomCell;
      });

    return {
      cells,
      height: round((row.height as number) * ctx.scaleY),
    };
  });

  return {
    type: "table",
    columns,
    rows,
  } as PomTableNode;
}

// ---------------------------------------------------------------------------
// Chart
// ---------------------------------------------------------------------------

const POM_CHART_TYPES = new Set(["bar", "line", "pie", "area", "doughnut", "radar"]);

function convertChart(chart: ChartElement): PomChartNode | null {
  const chartType = chart.chart.chartType;
  if (!POM_CHART_TYPES.has(chartType)) return null;

  const data: PomChartData[] = chart.chart.series.map((s) => ({
    name: s.name ?? undefined,
    labels: chart.chart.categories,
    values: s.values,
  }));

  const chartColors = chart.chart.series.map((s) => stripHash(s.color.hex));

  const node: PomChartNode = {
    type: "chart",
    chartType: chartType as PomChartType,
    data,
  };

  if (chart.chart.title) {
    node.title = chart.chart.title;
    node.showTitle = true;
  }
  if (chart.chart.legend) {
    node.showLegend = true;
  }
  if (chartColors.length > 0) {
    node.chartColors = chartColors;
  }
  if (chartType === "radar" && chart.chart.radarStyle) {
    node.radarStyle = chart.chart.radarStyle;
  }

  return node;
}

// ---------------------------------------------------------------------------
// Connector → Line
// ---------------------------------------------------------------------------

function convertConnector(connector: ConnectorElement, ctx: SlideScaleContext): PomLineNode {
  const t = connector.transform;
  const x = (t.offsetX as number) * ctx.scaleX;
  const y = (t.offsetY as number) * ctx.scaleY;
  const w = (t.extentWidth as number) * ctx.scaleX;
  const h = (t.extentHeight as number) * ctx.scaleY;

  const flipH = t.flipH;
  const flipV = t.flipV;

  const node: PomLineNode = {
    type: "line",
    x1: round(flipH ? x + w : x),
    y1: round(flipV ? y + h : y),
    x2: round(flipH ? x : x + w),
    y2: round(flipV ? y : y + h),
  };

  if (connector.outline) {
    const outline = connector.outline;
    if (outline.fill && outline.fill.type === "solid") {
      node.color = stripHash(outline.fill.color.hex);
    }
    if ((outline.width as number) > 0) {
      node.lineWidth = round(emuToPoints(outline.width as number));
    }
    if (outline.dashStyle !== "solid") {
      node.dashType = outline.dashStyle as PomDashType;
    }
    if (outline.headEnd && outline.headEnd.type !== "none") {
      node.beginArrow = convertArrow(outline.headEnd.type);
    }
    if (outline.tailEnd && outline.tailEnd.type !== "none") {
      node.endArrow = convertArrow(outline.tailEnd.type);
    }
  }

  return node;
}

function convertArrow(arrowType: string): PomLineArrow {
  if (
    arrowType === "triangle" ||
    arrowType === "stealth" ||
    arrowType === "diamond" ||
    arrowType === "oval" ||
    arrowType === "arrow"
  ) {
    return { type: arrowType };
  }
  return true;
}

// ---------------------------------------------------------------------------
// Group → Layer
// ---------------------------------------------------------------------------

function convertGroup(group: GroupElement, ctx: SlideScaleContext): PomLayerNode {
  const children: PomLayerChild[] = [];

  // Group child coordinates are in the group's local coordinate system (chOff/chExt).
  // We need to transform: (child - chOff) * (group.extent / chExt) + group.offset
  // then apply the slide scale.
  const chOff = group.childTransform;
  const grp = group.transform;
  const chExtW = chOff.extentWidth as number;
  const chExtH = chOff.extentHeight as number;
  const grpExtW = grp.extentWidth as number;
  const grpExtH = grp.extentHeight as number;
  const grpOffX = grp.offsetX as number;
  const grpOffY = grp.offsetY as number;
  const chOffX = chOff.offsetX as number;
  const chOffY = chOff.offsetY as number;

  for (const child of group.children) {
    const pomNode = convertElement(child, ctx);
    if (!pomNode) continue;

    const t = child.transform;
    // Map from group-local coords to slide coords
    const slideX =
      chExtW > 0
        ? grpOffX + ((t.offsetX as number) - chOffX) * (grpExtW / chExtW)
        : (t.offsetX as number);
    const slideY =
      chExtH > 0
        ? grpOffY + ((t.offsetY as number) - chOffY) * (grpExtH / chExtH)
        : (t.offsetY as number);
    const slideW =
      chExtW > 0 ? (t.extentWidth as number) * (grpExtW / chExtW) : (t.extentWidth as number);
    const slideH =
      chExtH > 0 ? (t.extentHeight as number) * (grpExtH / chExtH) : (t.extentHeight as number);

    const x = round(slideX * ctx.scaleX);
    const y = round(slideY * ctx.scaleY);
    const w = round(slideW * ctx.scaleX);
    const h = round(slideH * ctx.scaleY);

    children.push({ ...pomNode, x, y, w, h } as PomLayerChild);
  }

  return {
    type: "layer",
    children,
  };
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function convertOutlineToLine(
  outline: Outline,
  _ctx: SlideScaleContext,
): { color?: string; width?: number; dashType?: PomDashType } {
  const result: { color?: string; width?: number; dashType?: PomDashType } = {};
  if (outline.fill && outline.fill.type === "solid") {
    result.color = stripHash(outline.fill.color.hex);
  }
  if ((outline.width as number) > 0) {
    result.width = round(emuToPoints(outline.width as number));
  }
  if (outline.dashStyle !== "solid") {
    result.dashType = outline.dashStyle as PomDashType;
  }
  return result;
}

function convertAlignment(alignment: "l" | "ctr" | "r" | "just"): "left" | "center" | "right" {
  switch (alignment) {
    case "ctr":
      return "center";
    case "r":
      return "right";
    default:
      return "left";
  }
}

function ptToPx(pt: number): number {
  // 1pt = 4/3 px at 96 DPI
  return Math.round((pt * 4) / 3);
}

function emuToPoints(emu: number): number {
  return emu / 12700;
}

function round(n: number): number {
  return Math.round(n * 100) / 100;
}
