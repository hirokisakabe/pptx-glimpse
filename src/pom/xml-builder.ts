import type {
  PomChartNode,
  PomHStackNode,
  PomImageNode,
  PomLayerChild,
  PomLayerNode,
  PomLineNode,
  PomLiNode,
  PomNode,
  PomOlNode,
  PomShapeNode,
  PomTableNode,
  PomTextNode,
  PomUlNode,
  PomVStackNode,
} from "./pom-types.js";

export function pomLayerToXml(layer: PomLayerNode): string {
  return renderLayerNode(layer, 0);
}

function renderNode(node: PomNode, indent: number): string {
  switch (node.type) {
    case "text":
      return renderTextNode(node, indent);
    case "ul":
      return renderUlNode(node, indent);
    case "ol":
      return renderOlNode(node, indent);
    case "image":
      return renderImageNode(node, indent);
    case "table":
      return renderTableNode(node, indent);
    case "shape":
      return renderShapeNode(node, indent);
    case "chart":
      return renderChartNode(node, indent);
    case "line":
      return renderLineNode(node, indent);
    case "layer":
      return renderLayerNode(node, indent);
    case "vstack":
      return renderStackNode(node, indent);
    case "hstack":
      return renderStackNode(node, indent);
    case "box":
      return renderBoxNode(node, indent);
  }
}

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

function renderTextNode(node: PomTextNode, indent: number): string {
  const attrs: string[] = [];
  addBaseAttrs(attrs, node);
  addTextStyleAttrs(attrs, node);

  const pad = "  ".repeat(indent);
  return `${pad}<Text${attrStr(attrs)}>${escapeXml(node.text)}</Text>`;
}

// ---------------------------------------------------------------------------
// Lists
// ---------------------------------------------------------------------------

function renderUlNode(node: PomUlNode, indent: number): string {
  const attrs: string[] = [];
  addBaseAttrs(attrs, node);
  addTextStyleAttrs(attrs, node);

  const pad = "  ".repeat(indent);
  const children = node.items.map((item) => renderLiNode(item, indent + 1)).join("\n");
  return `${pad}<Ul${attrStr(attrs)}>\n${children}\n${pad}</Ul>`;
}

function renderOlNode(node: PomOlNode, indent: number): string {
  const attrs: string[] = [];
  addBaseAttrs(attrs, node);
  addTextStyleAttrs(attrs, node);
  if (node.numberType) attrs.push(a("numberType", node.numberType));
  if (node.numberStartAt !== undefined) attrs.push(a("numberStartAt", node.numberStartAt));

  const pad = "  ".repeat(indent);
  const children = node.items.map((item) => renderLiNode(item, indent + 1)).join("\n");
  return `${pad}<Ol${attrStr(attrs)}>\n${children}\n${pad}</Ol>`;
}

function renderLiNode(item: PomLiNode, indent: number): string {
  const attrs: string[] = [];
  if (item.bold) attrs.push(a("bold", true));
  if (item.italic) attrs.push(a("italic", true));
  if (item.underline) attrs.push(a("underline", true));
  if (item.strike) attrs.push(a("strike", true));
  if (item.highlight) attrs.push(a("highlight", item.highlight));
  if (item.fontPx !== undefined) attrs.push(a("fontPx", item.fontPx));
  if (item.color) attrs.push(a("color", item.color));
  if (item.fontFamily) attrs.push(a("fontFamily", item.fontFamily));

  const pad = "  ".repeat(indent);
  return `${pad}<Li${attrStr(attrs)}>${escapeXml(item.text)}</Li>`;
}

// ---------------------------------------------------------------------------
// Image
// ---------------------------------------------------------------------------

function renderImageNode(node: PomImageNode, indent: number): string {
  const attrs: string[] = [];
  addBaseAttrs(attrs, node);
  attrs.push(a("src", node.src));

  const pad = "  ".repeat(indent);
  return `${pad}<Image${attrStr(attrs)} />`;
}

// ---------------------------------------------------------------------------
// Table
// ---------------------------------------------------------------------------

function renderTableNode(node: PomTableNode, indent: number): string {
  const attrs: string[] = [];
  addBaseAttrs(attrs, node);
  if (node.defaultRowHeight !== undefined) attrs.push(a("defaultRowHeight", node.defaultRowHeight));

  const pad = "  ".repeat(indent);
  const lines: string[] = [`${pad}<Table${attrStr(attrs)}>`];

  for (const col of node.columns) {
    const colAttrs: string[] = [];
    if (col.width !== undefined) colAttrs.push(a("width", col.width));
    lines.push(`${pad}  <TableColumn${attrStr(colAttrs)} />`);
  }

  for (const row of node.rows) {
    const rowAttrs: string[] = [];
    if (row.height !== undefined) rowAttrs.push(a("height", row.height));
    lines.push(`${pad}  <TableRow${attrStr(rowAttrs)}>`);
    for (const cell of row.cells) {
      const cellAttrs: string[] = [];
      if (cell.bold) cellAttrs.push(a("bold", true));
      if (cell.italic) cellAttrs.push(a("italic", true));
      if (cell.fontPx !== undefined) cellAttrs.push(a("fontPx", cell.fontPx));
      if (cell.color) cellAttrs.push(a("color", cell.color));
      if (cell.alignText) cellAttrs.push(a("alignText", cell.alignText));
      if (cell.backgroundColor) cellAttrs.push(a("backgroundColor", cell.backgroundColor));
      if (cell.colspan !== undefined && cell.colspan > 1)
        cellAttrs.push(a("colspan", cell.colspan));
      if (cell.rowspan !== undefined && cell.rowspan > 1)
        cellAttrs.push(a("rowspan", cell.rowspan));
      lines.push(`${pad}    <TableCell${attrStr(cellAttrs)}>${escapeXml(cell.text)}</TableCell>`);
    }
    lines.push(`${pad}  </TableRow>`);
  }

  lines.push(`${pad}</Table>`);
  return lines.join("\n");
}

// ---------------------------------------------------------------------------
// Shape
// ---------------------------------------------------------------------------

function renderShapeNode(node: PomShapeNode, indent: number): string {
  const attrs: string[] = [];
  addBaseAttrs(attrs, node);
  attrs.push(a("shapeType", node.shapeType));
  addJsonAttr(attrs, "fill", node.fill);
  addJsonAttr(attrs, "line", node.line);
  addJsonAttr(attrs, "shadow", node.shadow);
  addTextStyleAttrs(attrs, node);

  const pad = "  ".repeat(indent);
  if (node.text) {
    return `${pad}<Shape${attrStr(attrs)}>${escapeXml(node.text)}</Shape>`;
  }
  return `${pad}<Shape${attrStr(attrs)} />`;
}

// ---------------------------------------------------------------------------
// Chart
// ---------------------------------------------------------------------------

function renderChartNode(node: PomChartNode, indent: number): string {
  const attrs: string[] = [];
  addBaseAttrs(attrs, node);
  attrs.push(a("chartType", node.chartType));
  if (node.showTitle) attrs.push(a("showTitle", true));
  if (node.title) attrs.push(a("title", node.title));
  if (node.showLegend) attrs.push(a("showLegend", true));
  if (node.chartColors && node.chartColors.length > 0) {
    attrs.push(`chartColors="${escapeXmlAttr(JSON.stringify(node.chartColors))}"`);
  }
  if (node.radarStyle) attrs.push(a("radarStyle", node.radarStyle));

  const pad = "  ".repeat(indent);
  const lines: string[] = [`${pad}<Chart${attrStr(attrs)}>`];

  for (const series of node.data) {
    const seriesAttrs: string[] = [];
    if (series.name) seriesAttrs.push(a("name", series.name));
    lines.push(`${pad}  <ChartSeries${attrStr(seriesAttrs)}>`);
    for (let i = 0; i < series.labels.length; i++) {
      lines.push(
        `${pad}    <ChartDataPoint ${a("label", series.labels[i])} ${a("value", series.values[i])} />`,
      );
    }
    lines.push(`${pad}  </ChartSeries>`);
  }

  lines.push(`${pad}</Chart>`);
  return lines.join("\n");
}

// ---------------------------------------------------------------------------
// Line
// ---------------------------------------------------------------------------

function renderLineNode(node: PomLineNode, indent: number): string {
  const attrs: string[] = [];
  addBaseAttrs(attrs, node);
  attrs.push(a("x1", node.x1));
  attrs.push(a("y1", node.y1));
  attrs.push(a("x2", node.x2));
  attrs.push(a("y2", node.y2));
  if (node.color) attrs.push(a("color", node.color));
  if (node.lineWidth !== undefined) attrs.push(a("lineWidth", node.lineWidth));
  if (node.dashType) attrs.push(a("dashType", node.dashType));
  addArrowAttr(attrs, "beginArrow", node.beginArrow);
  addArrowAttr(attrs, "endArrow", node.endArrow);

  const pad = "  ".repeat(indent);
  return `${pad}<Line${attrStr(attrs)} />`;
}

// ---------------------------------------------------------------------------
// Layout containers
// ---------------------------------------------------------------------------

function renderLayerNode(node: PomLayerNode, indent: number): string {
  const attrs: string[] = [];
  addBaseAttrs(attrs, node);

  const pad = "  ".repeat(indent);
  if (node.children.length === 0) {
    return `${pad}<Layer${attrStr(attrs)} />`;
  }

  const children = node.children.map((child) => renderLayerChild(child, indent + 1)).join("\n");
  return `${pad}<Layer${attrStr(attrs)}>\n${children}\n${pad}</Layer>`;
}

function renderLayerChild(child: PomLayerChild, indent: number): string {
  const { x, y, ...nodeRest } = child;
  const xml = renderNode(nodeRest as PomNode, indent);

  // Insert x and y attributes right after the tag name
  const openBracket = xml.indexOf("<");
  const firstSpace = xml.indexOf(" ", openBracket + 1);
  const firstClose = xml.indexOf(">", openBracket + 1);
  const firstSelfClose = xml.indexOf("/>", openBracket + 1);

  const insertPos = Math.min(
    firstSpace >= 0 ? firstSpace : Infinity,
    firstClose >= 0 ? firstClose : Infinity,
    firstSelfClose >= 0 ? firstSelfClose : Infinity,
  );

  if (insertPos === Infinity) return xml;

  const posAttrs = ` x="${x}" y="${y}"`;
  return xml.slice(0, insertPos) + posAttrs + xml.slice(insertPos);
}

function renderStackNode(node: PomVStackNode | PomHStackNode, indent: number): string {
  const tagName = node.type === "vstack" ? "VStack" : "HStack";
  const attrs: string[] = [];
  addBaseAttrs(attrs, node);
  if (node.gap !== undefined) attrs.push(a("gap", node.gap));
  if (node.alignItems) attrs.push(a("alignItems", node.alignItems));
  if (node.justifyContent) attrs.push(a("justifyContent", node.justifyContent));

  const pad = "  ".repeat(indent);
  if (node.children.length === 0) {
    return `${pad}<${tagName}${attrStr(attrs)} />`;
  }

  const children = node.children.map((c) => renderNode(c, indent + 1)).join("\n");
  return `${pad}<${tagName}${attrStr(attrs)}>\n${children}\n${pad}</${tagName}>`;
}

function renderBoxNode(
  node: { type: "box"; children: PomNode; shadow?: unknown } & BaseNodeLike,
  indent: number,
): string {
  const attrs: string[] = [];
  addBaseAttrs(attrs, node);

  const pad = "  ".repeat(indent);
  const child = renderNode(node.children, indent + 1);
  return `${pad}<Box${attrStr(attrs)}>\n${child}\n${pad}</Box>`;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

interface BaseNodeLike {
  w?: number | string;
  h?: number | string;
  backgroundColor?: string;
  border?: { color?: string; width?: number; dashType?: string };
  borderRadius?: number;
  opacity?: number;
  padding?: number | { top?: number; right?: number; bottom?: number; left?: number };
}

function addBaseAttrs(attrs: string[], node: BaseNodeLike): void {
  if (node.w !== undefined) attrs.push(a("w", node.w));
  if (node.h !== undefined) attrs.push(a("h", node.h));
  if (node.backgroundColor) attrs.push(a("backgroundColor", node.backgroundColor));
  addJsonAttr(attrs, "border", node.border);
  if (node.borderRadius !== undefined) attrs.push(a("borderRadius", node.borderRadius));
  if (node.opacity !== undefined && node.opacity !== 1) attrs.push(a("opacity", node.opacity));
  if (node.padding !== undefined) {
    if (typeof node.padding === "number") {
      attrs.push(a("padding", node.padding));
    }
  }
}

function addTextStyleAttrs(
  attrs: string[],
  node: {
    fontPx?: number;
    color?: string;
    alignText?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    highlight?: string;
    fontFamily?: string;
    lineSpacingMultiple?: number;
  },
): void {
  if (node.fontPx !== undefined) attrs.push(a("fontPx", node.fontPx));
  if (node.color) attrs.push(a("color", node.color));
  if (node.alignText) attrs.push(a("alignText", node.alignText));
  if (node.bold) attrs.push(a("bold", true));
  if (node.italic) attrs.push(a("italic", true));
  if (node.underline) attrs.push(a("underline", true));
  if (node.strike) attrs.push(a("strike", true));
  if (node.highlight) attrs.push(a("highlight", node.highlight));
  if (node.fontFamily) attrs.push(a("fontFamily", node.fontFamily));
  if (node.lineSpacingMultiple !== undefined)
    attrs.push(a("lineSpacingMultiple", node.lineSpacingMultiple));
}

/** Serialize an object value as a JSON attribute, omitting undefined fields */
function addJsonAttr(
  attrs: string[],
  name: string,
  obj: object | undefined,
): void {
  if (!obj) return;
  const cleaned: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(obj)) {
    if (v !== undefined) cleaned[k] = v;
  }
  if (Object.keys(cleaned).length === 0) return;
  attrs.push(`${name}="${escapeXmlAttr(JSON.stringify(cleaned))}"`);
}

/** Serialize a line arrow attribute (boolean | {type}) */
function addArrowAttr(
  attrs: string[],
  name: string,
  arrow: boolean | { type?: string } | undefined,
): void {
  if (arrow === undefined) return;
  if (typeof arrow === "boolean") {
    attrs.push(a(name, true));
  } else {
    addJsonAttr(attrs, name, arrow as Record<string, unknown>);
  }
}

/** Build a single XML attribute string with proper escaping */
function a(name: string, value: string | number | boolean): string {
  return `${name}="${escapeXmlAttr(String(value))}"`;
}

function attrStr(attrs: string[]): string {
  return attrs.length > 0 ? " " + attrs.join(" ") : "";
}

function escapeXml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeXmlAttr(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
