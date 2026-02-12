import type { Fill, GradientStop, ImageFill, PatternFill, SolidFill } from "../model/fill.js";
import type { Outline, DashStyle, ArrowEndpoint, ArrowType, ArrowSize } from "../model/line.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { PptxArchive } from "./pptx-reader.js";
import type { Relationship } from "./relationship-parser.js";
import { resolveRelationshipTarget } from "./relationship-parser.js";
import type { XmlNode } from "./xml-parser.js";
import { warn, debug } from "../warning-logger.js";

export interface FillParseContext {
  rels: Map<string, Relationship>;
  archive: PptxArchive;
  basePath: string;
  groupFill?: Fill;
}

export function parseFillFromNode(
  node: XmlNode,
  colorResolver: ColorResolver,
  context?: FillParseContext,
): Fill | null {
  if (!node) return null;

  if (node.noFill !== undefined) {
    return { type: "none" };
  }

  if (node.solidFill) {
    const color = colorResolver.resolve(node.solidFill as XmlNode);
    if (color) {
      return { type: "solid", color };
    }
  }

  if (node.gradFill) {
    return parseGradientFill(node.gradFill as XmlNode, colorResolver);
  }

  if (node.blipFill && context) {
    return parseBlipFill(node.blipFill as XmlNode, context);
  }

  if (node.pattFill) {
    return parsePatternFill(node.pattFill as XmlNode, colorResolver);
  }

  if (node.grpFill !== undefined && context?.groupFill) {
    return context.groupFill;
  }

  return null;
}

function parseBlipFill(blipFillNode: XmlNode, context: FillParseContext): ImageFill | null {
  const blip = blipFillNode?.blip as XmlNode | undefined;
  const rId =
    (blip?.["@_r:embed"] as string | undefined) ?? (blip?.["@_embed"] as string | undefined);
  if (!rId) return null;

  const rel = context.rels.get(rId);
  if (!rel) return null;

  const mediaPath = resolveRelationshipTarget(context.basePath, rel.target);
  const mediaData = context.archive.media.get(mediaPath);
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

  return { type: "image", imageData, mimeType };
}

function parseGradientFill(gradNode: XmlNode, colorResolver: ColorResolver): Fill | null {
  const gsLst = gradNode.gsLst as XmlNode | undefined;
  const gsArr = gsLst?.gs as XmlNode[] | undefined;
  if (!gsArr) {
    debug("gradientFill.gsLst", "GradientFill: gsLst not found, skipping gradient");
    return null;
  }

  const stops: GradientStop[] = [];
  for (const gs of gsArr) {
    const position = Number(gs["@_pos"] ?? 0) / 100000;
    const color = colorResolver.resolve(gs);
    if (color) {
      stops.push({ position, color });
    }
  }

  const pathNode = gradNode.path as XmlNode | undefined;
  if (pathNode) {
    const fillToRect = pathNode.fillToRect as XmlNode | undefined;
    let centerX = 0.5;
    let centerY = 0.5;
    if (fillToRect) {
      const l = Number(fillToRect["@_l"] ?? 0);
      const t = Number(fillToRect["@_t"] ?? 0);
      const r = Number(fillToRect["@_r"] ?? 0);
      const b = Number(fillToRect["@_b"] ?? 0);
      centerX = (l + (100000 - r)) / 2 / 100000;
      centerY = (t + (100000 - b)) / 2 / 100000;
    }
    return { type: "gradient", stops, angle: 0, gradientType: "radial", centerX, centerY };
  }

  let angle = 0;
  const lin = gradNode.lin as XmlNode | undefined;
  if (lin) {
    angle = Number(lin["@_ang"] ?? 0) / 60000;
  }

  return { type: "gradient", stops, angle, gradientType: "linear" };
}

function parsePatternFill(pattNode: XmlNode, colorResolver: ColorResolver): PatternFill | null {
  const preset = (pattNode["@_prst"] as string | undefined) ?? "ltDnDiag";
  const fgColor = pattNode.fgClr ? colorResolver.resolve(pattNode.fgClr as XmlNode) : null;
  const bgColor = pattNode.bgClr ? colorResolver.resolve(pattNode.bgClr as XmlNode) : null;

  if (!fgColor || !bgColor) return null;

  return { type: "pattern", preset, foregroundColor: fgColor, backgroundColor: bgColor };
}

export function parseOutline(lnNode: XmlNode, colorResolver: ColorResolver): Outline | null {
  if (!lnNode) return null;

  // Unsupported feature detection
  if (lnNode.headEnd) {
    const headType = ((lnNode.headEnd as XmlNode)["@_type"] as string | undefined) ?? "unknown";
    warn("ln.headEnd", `line head arrow (type="${headType}") not implemented`);
  }
  if (lnNode.tailEnd) {
    const tailType = ((lnNode.tailEnd as XmlNode)["@_type"] as string | undefined) ?? "unknown";
    warn("ln.tailEnd", `line tail arrow (type="${tailType}") not implemented`);
  }
  if (lnNode.gradFill) {
    warn("ln.gradFill", "gradient line fill not implemented");
  }
  if (lnNode.pattFill) {
    warn("ln.pattFill", "pattern line fill not implemented");
  }

  const width = Number(lnNode["@_w"] ?? 12700);

  let fill: SolidFill | null = null;
  if (lnNode.solidFill) {
    const color = colorResolver.resolve(lnNode.solidFill as XmlNode);
    if (color) {
      fill = { type: "solid", color };
    }
  }

  if (lnNode.noFill !== undefined) {
    return null;
  }

  const prstDash = lnNode.prstDash as XmlNode | undefined;
  const dashStyle = ((prstDash?.["@_val"] as string | undefined) ?? "solid") as DashStyle;

  const headEnd = parseArrowEndpoint(lnNode.headEnd as XmlNode);
  const tailEnd = parseArrowEndpoint(lnNode.tailEnd as XmlNode);

  return { width, fill, dashStyle, headEnd, tailEnd };
}

function parseArrowEndpoint(node: XmlNode): ArrowEndpoint | null {
  if (!node) return null;
  const type = ((node["@_type"] as string | undefined) ?? "none") as ArrowType;
  if (type === "none") return null;
  return {
    type,
    width: ((node["@_w"] as string | undefined) ?? "med") as ArrowSize,
    length: ((node["@_len"] as string | undefined) ?? "med") as ArrowSize,
  };
}
