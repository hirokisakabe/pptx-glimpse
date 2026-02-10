import type { Slide, Background } from "../model/slide.js";
import type {
  SlideElement,
  ShapeElement,
  ConnectorElement,
  GroupElement,
  Transform,
  Geometry,
} from "../model/shape.js";
import type { ImageElement } from "../model/image.js";
import type { ChartElement } from "../model/chart.js";
import type { TextBody, BodyProperties, Paragraph, TextRun, RunProperties } from "../model/text.js";
import type { PptxArchive } from "./pptx-reader.js";
import type { Relationship } from "./relationship-parser.js";
import type { ColorResolver } from "../color/color-resolver.js";
import { parseXml } from "./xml-parser.js";
import { parseFillFromNode, parseOutline } from "./fill-parser.js";
import { parseChart } from "./chart-parser.js";
import { parseRelationships, resolveRelationshipTarget } from "./relationship-parser.js";
import { hundredthPointToPoint } from "../utils/emu.js";

export function parseSlide(
  slideXml: string,
  slidePath: string,
  slideNumber: number,
  archive: PptxArchive,
  colorResolver: ColorResolver,
): Slide {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(slideXml) as any;
  const sld = parsed.sld;

  const relsPath = slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
  const relsXml = archive.files.get(relsPath);
  const rels = relsXml ? parseRelationships(relsXml) : new Map<string, Relationship>();

  const background = parseBackground(sld?.cSld?.bg, colorResolver);
  const elements = parseShapeTree(sld?.cSld?.spTree, rels, slidePath, archive, colorResolver);

  return { slideNumber, background, elements };
}

function parseBackground(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  bgNode: any,
  colorResolver: ColorResolver,
): Background | null {
  if (!bgNode) return null;

  const bgPr = bgNode.bgPr;
  if (!bgPr) return null;

  const fill = parseFillFromNode(bgPr, colorResolver);
  return { fill };
}

export function parseShapeTree(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  spTree: any,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
): SlideElement[] {
  if (!spTree) return [];

  const elements: SlideElement[] = [];

  const shapes = spTree.sp ?? [];
  for (const sp of shapes) {
    const shape = parseShape(sp, colorResolver);
    if (shape) elements.push(shape);
  }

  const pics = spTree.pic ?? [];
  for (const pic of pics) {
    const img = parseImage(pic, rels, slidePath, archive);
    if (img) elements.push(img);
  }

  const cxnSps = spTree.cxnSp ?? [];
  for (const cxn of cxnSps) {
    const connector = parseConnector(cxn, colorResolver);
    if (connector) elements.push(connector);
  }

  const grpSps = spTree.grpSp ?? [];
  for (const grp of grpSps) {
    const group = parseGroup(grp, rels, slidePath, archive, colorResolver);
    if (group) elements.push(group);
  }

  const graphicFrames = spTree.graphicFrame ?? [];
  for (const gf of graphicFrames) {
    const chart = parseGraphicFrame(gf, rels, slidePath, archive, colorResolver);
    if (chart) elements.push(chart);
  }

  return elements;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseShape(sp: any, colorResolver: ColorResolver): ShapeElement | null {
  const spPr = sp.spPr;
  if (!spPr) return null;

  const transform = parseTransform(spPr.xfrm);
  if (!transform) return null;

  const geometry = parseGeometry(spPr);
  const fill = parseFillFromNode(spPr, colorResolver);
  const outline = parseOutline(spPr.ln, colorResolver);
  const textBody = parseTextBody(sp.txBody, colorResolver);

  const ph = sp.nvSpPr?.nvPr?.ph;
  const placeholderType = ph ? (ph["@_type"] ?? "body") : undefined;
  const placeholderIdx = ph?.["@_idx"] !== undefined ? Number(ph["@_idx"]) : undefined;

  return {
    type: "shape",
    transform,
    geometry,
    fill,
    outline,
    textBody,
    ...(placeholderType !== undefined && { placeholderType }),
    ...(placeholderIdx !== undefined && { placeholderIdx }),
  };
}

function parseImage(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  pic: any,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
): ImageElement | null {
  const spPr = pic.spPr;
  if (!spPr) return null;

  const transform = parseTransform(spPr.xfrm);
  if (!transform) return null;

  const blipFill = pic.blipFill;
  const rId = blipFill?.blip?.["@_r:embed"] ?? blipFill?.blip?.["@_embed"];
  if (!rId) return null;

  const rel = rels.get(rId);
  if (!rel) return null;

  const mediaPath = resolveRelationshipTarget(slidePath, rel.target);
  const mediaData = archive.media.get(mediaPath);
  if (!mediaData) return null;

  const ext = mediaPath.split(".").pop()?.toLowerCase() ?? "png";
  const mimeMap: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    svg: "image/svg+xml",
    emf: "image/emf",
    wmf: "image/wmf",
  };
  const mimeType = mimeMap[ext] ?? "image/png";
  const imageData = mediaData.toString("base64");

  return {
    type: "image",
    transform,
    imageData,
    mimeType,
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseConnector(cxn: any, colorResolver: ColorResolver): ConnectorElement | null {
  const spPr = cxn.spPr;
  if (!spPr) return null;

  const transform = parseTransform(spPr.xfrm);
  if (!transform) return null;

  const outline = parseOutline(spPr.ln, colorResolver);

  return { type: "connector", transform, outline };
}

function parseGroup(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  grp: any,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
): GroupElement | null {
  const grpSpPr = grp.grpSpPr;
  if (!grpSpPr) return null;

  const transform = parseTransform(grpSpPr.xfrm);
  if (!transform) return null;

  const childOff = grpSpPr.xfrm?.chOff;
  const childExt = grpSpPr.xfrm?.chExt;
  const childTransform: Transform = {
    offsetX: Number(childOff?.["@_x"] ?? 0),
    offsetY: Number(childOff?.["@_y"] ?? 0),
    extentWidth: Number(childExt?.["@_cx"] ?? transform.extentWidth),
    extentHeight: Number(childExt?.["@_cy"] ?? transform.extentHeight),
    rotation: 0,
    flipH: false,
    flipV: false,
  };

  const children = parseShapeTree(grp, rels, slidePath, archive, colorResolver);

  return { type: "group", transform, childTransform, children };
}

function parseGraphicFrame(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  gf: any,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
): ChartElement | null {
  const xfrm = gf.xfrm;
  const transform = parseTransform(xfrm);
  if (!transform) return null;

  const graphicData = gf.graphic?.graphicData;
  if (!graphicData) return null;

  const chartRef = graphicData.chart;
  if (!chartRef) return null;

  const rId = chartRef["@_r:id"] ?? chartRef["@_id"];
  if (!rId) return null;

  const rel = rels.get(rId);
  if (!rel) return null;

  const chartPath = resolveRelationshipTarget(slidePath, rel.target);
  const chartXml = archive.files.get(chartPath);
  if (!chartXml) return null;

  const chartData = parseChart(chartXml, colorResolver);
  if (!chartData) return null;

  return { type: "chart", transform, chart: chartData };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseTransform(xfrm: any): Transform | null {
  if (!xfrm) return null;

  const off = xfrm.off;
  const ext = xfrm.ext;
  if (!off || !ext) return null;

  return {
    offsetX: Number(off["@_x"] ?? 0),
    offsetY: Number(off["@_y"] ?? 0),
    extentWidth: Number(ext["@_cx"] ?? 0),
    extentHeight: Number(ext["@_cy"] ?? 0),
    rotation: Number(xfrm["@_rot"] ?? 0) / 60000,
    flipH: xfrm["@_flipH"] === "1" || xfrm["@_flipH"] === "true",
    flipV: xfrm["@_flipV"] === "1" || xfrm["@_flipV"] === "true",
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseGeometry(spPr: any): Geometry {
  if (spPr.prstGeom) {
    const preset = spPr.prstGeom["@_prst"] ?? "rect";
    const avLst = spPr.prstGeom.avLst;
    const adjustValues: Record<string, number> = {};

    if (avLst?.gd) {
      const guides = Array.isArray(avLst.gd) ? avLst.gd : [avLst.gd];
      for (const gd of guides) {
        const name = gd["@_name"] as string;
        const fmla = gd["@_fmla"] as string;
        const match = fmla?.match(/val\s+(\d+)/);
        if (name && match) {
          adjustValues[name] = Number(match[1]);
        }
      }
    }

    return { type: "preset", preset, adjustValues };
  }

  if (spPr.custGeom) {
    const pathData = parseCustomGeometryPaths(spPr.custGeom);
    if (pathData) {
      return { type: "custom", pathData };
    }
  }

  return { type: "preset", preset: "rect", adjustValues: {} };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseCustomGeometryPaths(custGeom: any): string | null {
  const pathLst = custGeom.pathLst;
  if (!pathLst?.path) return null;

  const paths = Array.isArray(pathLst.path) ? pathLst.path : [pathLst.path];
  const svgParts: string[] = [];

  for (const path of paths) {
    const w = Number(path["@_w"] ?? 0);
    const h = Number(path["@_h"] ?? 0);
    if (w === 0 && h === 0) continue;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const processCommands = (commands: any[] | undefined, prefix: string) => {
      if (!commands) return;
      const list = Array.isArray(commands) ? commands : [commands];
      for (const cmd of list) {
        if (prefix === "M" || prefix === "L") {
          const pt = cmd.pt;
          if (pt) {
            const pts = Array.isArray(pt) ? pt : [pt];
            svgParts.push(
              `${prefix} ${pts.map((p: Record<string, string>) => `${p["@_x"]} ${p["@_y"]}`).join(" ")}`,
            );
          }
        }
      }
    };

    if (path.moveTo) processCommands(Array.isArray(path.moveTo) ? path.moveTo : [path.moveTo], "M");
    if (path.lnTo) processCommands(Array.isArray(path.lnTo) ? path.lnTo : [path.lnTo], "L");

    if (path.cubicBezTo) {
      const bezList = Array.isArray(path.cubicBezTo) ? path.cubicBezTo : [path.cubicBezTo];
      for (const bez of bezList) {
        const pts = Array.isArray(bez.pt) ? bez.pt : [bez.pt];
        if (pts.length >= 3) {
          svgParts.push(
            `C ${pts.map((p: Record<string, string>) => `${p["@_x"]} ${p["@_y"]}`).join(", ")}`,
          );
        }
      }
    }

    if (path.close !== undefined) {
      svgParts.push("Z");
    }
  }

  return svgParts.length > 0 ? svgParts.join(" ") : null;
}

function parseTextBody(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  txBody: any,
  colorResolver: ColorResolver,
): TextBody | null {
  if (!txBody) return null;

  const bodyPr = txBody.bodyPr;
  const bodyProperties: BodyProperties = {
    anchor: (bodyPr?.["@_anchor"] as "t" | "ctr" | "b") ?? "t",
    marginLeft: Number(bodyPr?.["@_lIns"] ?? 91440),
    marginRight: Number(bodyPr?.["@_rIns"] ?? 91440),
    marginTop: Number(bodyPr?.["@_tIns"] ?? 45720),
    marginBottom: Number(bodyPr?.["@_bIns"] ?? 45720),
    wrap: (bodyPr?.["@_wrap"] as "square" | "none") ?? "square",
  };

  const paragraphs: Paragraph[] = [];
  const pList = txBody.p ?? [];
  for (const p of pList) {
    paragraphs.push(parseParagraph(p, colorResolver));
  }

  if (paragraphs.length === 0) return null;

  return { paragraphs, bodyProperties };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseParagraph(p: any, colorResolver: ColorResolver): Paragraph {
  const pPr = p.pPr;
  const properties = {
    alignment: (pPr?.["@_algn"] as "l" | "ctr" | "r" | "just") ?? "l",
    lineSpacing: pPr?.lnSpc?.spcPct ? Number(pPr.lnSpc.spcPct["@_val"]) : null,
    spaceBefore: pPr?.spcBef?.spcPts ? Number(pPr.spcBef.spcPts["@_val"]) : 0,
    spaceAfter: pPr?.spcAft?.spcPts ? Number(pPr.spcAft.spcPts["@_val"]) : 0,
    level: Number(pPr?.["@_lvl"] ?? 0),
  };

  const runs: TextRun[] = [];
  const rList = p.r ?? [];
  for (const r of rList) {
    const text = r.t ?? "";
    const textContent = typeof text === "object" ? (text["#text"] ?? "") : String(text);
    const rPr = r.rPr;
    const runProps = parseRunProperties(rPr, colorResolver);
    runs.push({ text: textContent, properties: runProps });
  }

  return { runs, properties };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseRunProperties(rPr: any, colorResolver: ColorResolver): RunProperties {
  if (!rPr) {
    return {
      fontSize: null,
      fontFamily: null,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      color: null,
    };
  }

  const fontSize = rPr["@_sz"] ? hundredthPointToPoint(Number(rPr["@_sz"])) : null;
  const fontFamily = rPr.latin?.["@_typeface"] ?? rPr.ea?.["@_typeface"] ?? null;
  const bold = rPr["@_b"] === "1" || rPr["@_b"] === "true";
  const italic = rPr["@_i"] === "1" || rPr["@_i"] === "true";
  const underline = rPr["@_u"] !== undefined && rPr["@_u"] !== "none";
  const strikethrough = rPr["@_strike"] !== undefined && rPr["@_strike"] !== "noStrike";

  let color = colorResolver.resolve(rPr.solidFill ?? rPr);
  if (!rPr.solidFill && !rPr.srgbClr && !rPr.schemeClr && !rPr.sysClr) {
    color = null;
  }

  return { fontSize, fontFamily, bold, italic, underline, strikethrough, color };
}
