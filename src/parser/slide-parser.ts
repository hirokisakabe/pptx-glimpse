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
import type { TableElement } from "../model/table.js";
import type {
  TextBody,
  BodyProperties,
  Paragraph,
  TextRun,
  RunProperties,
  BulletType,
  AutoNumScheme,
} from "../model/text.js";
import type { PptxArchive } from "./pptx-reader.js";
import type { Relationship } from "./relationship-parser.js";
import type { ColorResolver } from "../color/color-resolver.js";
import { parseXml } from "./xml-parser.js";
import { parseFillFromNode, parseOutline } from "./fill-parser.js";
import type { FillParseContext } from "./fill-parser.js";
import { parseChart } from "./chart-parser.js";
import { parseTable } from "./table-parser.js";
import { parseRelationships, resolveRelationshipTarget } from "./relationship-parser.js";
import { hundredthPointToPoint } from "../utils/emu.js";

const WARN_PREFIX = "[pptx-glimpse]";

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
  if (!sld) {
    console.warn(`${WARN_PREFIX} Slide ${slideNumber}: missing root element "sld" in XML`);
  }

  const relsPath = slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
  const relsXml = archive.files.get(relsPath);
  const rels = relsXml ? parseRelationships(relsXml) : new Map<string, Relationship>();

  const fillContext: FillParseContext = { rels, archive, basePath: slidePath };
  const background = parseBackground(sld?.cSld?.bg, colorResolver, fillContext);
  const elements = parseShapeTree(
    sld?.cSld?.spTree,
    rels,
    slidePath,
    archive,
    colorResolver,
    `Slide ${slideNumber}`,
  );

  return { slideNumber, background, elements };
}

function parseBackground(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  bgNode: any,
  colorResolver: ColorResolver,
  context?: FillParseContext,
): Background | null {
  if (!bgNode) return null;

  const bgPr = bgNode.bgPr;
  if (!bgPr) return null;

  const fill = parseFillFromNode(bgPr, colorResolver, context);
  return { fill };
}

export function parseShapeTree(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  spTree: any,
  rels: Map<string, Relationship>,
  slidePath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  context?: string,
): SlideElement[] {
  if (!spTree) return [];

  const ctx = context ?? slidePath;
  const elements: SlideElement[] = [];

  const shapes = spTree.sp ?? [];
  for (const sp of shapes) {
    const shape = parseShape(sp, colorResolver);
    if (shape) {
      elements.push(shape);
    } else {
      console.warn(`${WARN_PREFIX} ${ctx}: shape skipped (parse returned null)`);
    }
  }

  const pics = spTree.pic ?? [];
  for (const pic of pics) {
    const img = parseImage(pic, rels, slidePath, archive);
    if (img) {
      elements.push(img);
    } else {
      console.warn(`${WARN_PREFIX} ${ctx}: image skipped (parse returned null)`);
    }
  }

  const cxnSps = spTree.cxnSp ?? [];
  for (const cxn of cxnSps) {
    const connector = parseConnector(cxn, colorResolver);
    if (connector) {
      elements.push(connector);
    } else {
      console.warn(`${WARN_PREFIX} ${ctx}: connector skipped (parse returned null)`);
    }
  }

  const grpSps = spTree.grpSp ?? [];
  for (const grp of grpSps) {
    const group = parseGroup(grp, rels, slidePath, archive, colorResolver);
    if (group) {
      elements.push(group);
    } else {
      console.warn(`${WARN_PREFIX} ${ctx}: group skipped (parse returned null)`);
    }
  }

  const graphicFrames = spTree.graphicFrame ?? [];
  for (const gf of graphicFrames) {
    const chart = parseGraphicFrame(gf, rels, slidePath, archive, colorResolver);
    if (chart) {
      elements.push(chart);
    } else {
      console.warn(`${WARN_PREFIX} ${ctx}: graphicFrame skipped (parse returned null)`);
    }
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
): ChartElement | TableElement | null {
  const xfrm = gf.xfrm;
  const transform = parseTransform(xfrm);
  if (!transform) return null;

  const graphicData = gf.graphic?.graphicData;
  if (!graphicData) return null;

  // Chart
  const chartRef = graphicData.chart;
  if (chartRef) {
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

  // Table
  const tblNode = graphicData.tbl;
  if (tblNode) {
    const tableData = parseTable(tblNode, colorResolver);
    if (!tableData) return null;

    return { type: "table", transform, table: tableData };
  }

  return null;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseTransform(xfrm: any): Transform | null {
  if (!xfrm) return null;

  const off = xfrm.off;
  const ext = xfrm.ext;
  if (!off || !ext) return null;

  let offsetX = Number(off["@_x"] ?? 0);
  let offsetY = Number(off["@_y"] ?? 0);
  let extentWidth = Number(ext["@_cx"] ?? 0);
  let extentHeight = Number(ext["@_cy"] ?? 0);
  let rotation = Number(xfrm["@_rot"] ?? 0);

  if (Number.isNaN(offsetX)) {
    console.warn(`${WARN_PREFIX} NaN detected in transform offsetX, defaulting to 0`);
    offsetX = 0;
  }
  if (Number.isNaN(offsetY)) {
    console.warn(`${WARN_PREFIX} NaN detected in transform offsetY, defaulting to 0`);
    offsetY = 0;
  }
  if (Number.isNaN(extentWidth)) {
    console.warn(`${WARN_PREFIX} NaN detected in transform extentWidth, defaulting to 0`);
    extentWidth = 0;
  }
  if (Number.isNaN(extentHeight)) {
    console.warn(`${WARN_PREFIX} NaN detected in transform extentHeight, defaulting to 0`);
    extentHeight = 0;
  }
  if (Number.isNaN(rotation)) {
    console.warn(`${WARN_PREFIX} NaN detected in transform rotation, defaulting to 0`);
    rotation = 0;
  }

  return {
    offsetX,
    offsetY,
    extentWidth,
    extentHeight,
    rotation: rotation / 60000,
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

export function parseTextBody(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  txBody: any,
  colorResolver: ColorResolver,
): TextBody | null {
  if (!txBody) return null;

  const bodyPr = txBody.bodyPr;

  let autoFit: BodyProperties["autoFit"] = "noAutofit";
  let fontScale = 1;
  let lnSpcReduction = 0;
  if (bodyPr?.normAutofit !== undefined) {
    autoFit = "normAutofit";
    const normAutofit = bodyPr.normAutofit;
    if (typeof normAutofit === "object" && normAutofit !== null) {
      fontScale = normAutofit["@_fontScale"] ? Number(normAutofit["@_fontScale"]) / 100000 : 1;
      lnSpcReduction = normAutofit["@_lnSpcReduction"]
        ? Number(normAutofit["@_lnSpcReduction"]) / 100000
        : 0;
    }
  } else if (bodyPr?.spAutoFit !== undefined) {
    autoFit = "spAutofit";
  }

  const bodyProperties: BodyProperties = {
    anchor: (bodyPr?.["@_anchor"] as "t" | "ctr" | "b") ?? "t",
    marginLeft: Number(bodyPr?.["@_lIns"] ?? 91440),
    marginRight: Number(bodyPr?.["@_rIns"] ?? 91440),
    marginTop: Number(bodyPr?.["@_tIns"] ?? 45720),
    marginBottom: Number(bodyPr?.["@_bIns"] ?? 45720),
    wrap: (bodyPr?.["@_wrap"] as "square" | "none") ?? "square",
    autoFit,
    fontScale,
    lnSpcReduction,
  };

  const paragraphs: Paragraph[] = [];
  const pList = txBody.p ?? [];
  for (const p of pList) {
    paragraphs.push(parseParagraph(p, colorResolver));
  }

  if (paragraphs.length === 0) return null;

  return { paragraphs, bodyProperties };
}

const VALID_AUTO_NUM_SCHEMES = new Set([
  "arabicPeriod",
  "arabicParenR",
  "romanUcPeriod",
  "romanLcPeriod",
  "alphaUcPeriod",
  "alphaLcPeriod",
  "alphaLcParenR",
  "alphaUcParenR",
  "arabicPlain",
]);

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseBullet(pPr: any, colorResolver: ColorResolver) {
  let bullet: BulletType | null = null;
  let bulletFont: string | null = null;
  let bulletColor = colorResolver.resolve(pPr?.buClr);
  if (!pPr?.buClr) bulletColor = null;
  const bulletSizePct: number | null = pPr?.buSzPct ? Number(pPr.buSzPct["@_val"]) : null;

  if (pPr?.buNone !== undefined) {
    bullet = { type: "none" };
  } else if (pPr?.buChar) {
    bullet = { type: "char", char: pPr.buChar["@_char"] ?? "\u2022" };
  } else if (pPr?.buAutoNum) {
    const scheme = pPr.buAutoNum["@_type"] ?? "arabicPeriod";
    bullet = {
      type: "autoNum",
      scheme: VALID_AUTO_NUM_SCHEMES.has(scheme) ? (scheme as AutoNumScheme) : "arabicPeriod",
      startAt: Number(pPr.buAutoNum["@_startAt"] ?? 1),
    };
  }

  if (pPr?.buFont) {
    bulletFont = pPr.buFont["@_typeface"] ?? null;
  }

  return { bullet, bulletFont, bulletColor, bulletSizePct };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseParagraph(p: any, colorResolver: ColorResolver): Paragraph {
  const pPr = p.pPr;
  const { bullet, bulletFont, bulletColor, bulletSizePct } = parseBullet(pPr, colorResolver);
  const properties = {
    alignment: (pPr?.["@_algn"] as "l" | "ctr" | "r" | "just") ?? "l",
    lineSpacing: pPr?.lnSpc?.spcPct ? Number(pPr.lnSpc.spcPct["@_val"]) : null,
    spaceBefore: pPr?.spcBef?.spcPts ? Number(pPr.spcBef.spcPts["@_val"]) : 0,
    spaceAfter: pPr?.spcAft?.spcPts ? Number(pPr.spcAft.spcPts["@_val"]) : 0,
    level: Number(pPr?.["@_lvl"] ?? 0),
    bullet,
    bulletFont,
    bulletColor,
    bulletSizePct,
    marginLeft: Number(pPr?.["@_marL"] ?? 0),
    indent: Number(pPr?.["@_indent"] ?? 0),
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
      baseline: 0,
    };
  }

  const fontSize = rPr["@_sz"] ? hundredthPointToPoint(Number(rPr["@_sz"])) : null;
  const fontFamily = rPr.latin?.["@_typeface"] ?? rPr.ea?.["@_typeface"] ?? null;
  const bold = rPr["@_b"] === "1" || rPr["@_b"] === "true";
  const italic = rPr["@_i"] === "1" || rPr["@_i"] === "true";
  const underline = rPr["@_u"] !== undefined && rPr["@_u"] !== "none";
  const strikethrough = rPr["@_strike"] !== undefined && rPr["@_strike"] !== "noStrike";
  const baseline = rPr["@_baseline"] ? Number(rPr["@_baseline"]) / 1000 : 0;

  let color = colorResolver.resolve(rPr.solidFill ?? rPr);
  if (!rPr.solidFill && !rPr.srgbClr && !rPr.schemeClr && !rPr.sysClr) {
    color = null;
  }

  return { fontSize, fontFamily, bold, italic, underline, strikethrough, color, baseline };
}
