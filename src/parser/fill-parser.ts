import type { Fill, GradientStop, ImageFill, PatternFill, SolidFill } from "../model/fill.js";
import type { Outline, DashStyle } from "../model/line.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { PptxArchive } from "./pptx-reader.js";
import type { Relationship } from "./relationship-parser.js";
import { resolveRelationshipTarget } from "./relationship-parser.js";
import { warn, debug } from "../warning-logger.js";

export interface FillParseContext {
  rels: Map<string, Relationship>;
  archive: PptxArchive;
  basePath: string;
  groupFill?: Fill;
}

export function parseFillFromNode(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  node: any,
  colorResolver: ColorResolver,
  context?: FillParseContext,
): Fill | null {
  if (!node) return null;

  if (node.noFill !== undefined) {
    return { type: "none" };
  }

  if (node.solidFill) {
    const color = colorResolver.resolve(node.solidFill);
    if (color) {
      return { type: "solid", color };
    }
  }

  if (node.gradFill) {
    return parseGradientFill(node.gradFill, colorResolver);
  }

  if (node.blipFill && context) {
    return parseBlipFill(node.blipFill, context);
  }

  if (node.pattFill) {
    return parsePatternFill(node.pattFill, colorResolver);
  }

  if (node.grpFill !== undefined && context?.groupFill) {
    return context.groupFill;
  }

  return null;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseBlipFill(blipFillNode: any, context: FillParseContext): ImageFill | null {
  const rId = blipFillNode?.blip?.["@_r:embed"] ?? blipFillNode?.blip?.["@_embed"];
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

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseGradientFill(gradNode: any, colorResolver: ColorResolver): Fill | null {
  const gsLst = gradNode.gsLst?.gs;
  if (!gsLst) {
    debug("gradientFill.gsLst", "GradientFill: gsLst not found, skipping gradient");
    return null;
  }

  const stops: GradientStop[] = [];
  for (const gs of gsLst) {
    const position = Number(gs["@_pos"] ?? 0) / 100000;
    const color = colorResolver.resolve(gs);
    if (color) {
      stops.push({ position, color });
    }
  }

  if (gradNode.path) {
    const fillToRect = gradNode.path.fillToRect;
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
  if (gradNode.lin) {
    angle = Number(gradNode.lin["@_ang"] ?? 0) / 60000;
  }

  return { type: "gradient", stops, angle, gradientType: "linear" };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parsePatternFill(pattNode: any, colorResolver: ColorResolver): PatternFill | null {
  const preset = pattNode["@_prst"] ?? "ltDnDiag";
  const fgColor = pattNode.fgClr ? colorResolver.resolve(pattNode.fgClr) : null;
  const bgColor = pattNode.bgClr ? colorResolver.resolve(pattNode.bgClr) : null;

  if (!fgColor || !bgColor) return null;

  return { type: "pattern", preset, foregroundColor: fgColor, backgroundColor: bgColor };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function parseOutline(lnNode: any, colorResolver: ColorResolver): Outline | null {
  if (!lnNode) return null;

  // Unsupported feature detection
  if (lnNode.headEnd) {
    const headType = lnNode.headEnd["@_type"] ?? "unknown";
    warn("ln.headEnd", `line head arrow (type="${headType}") not implemented`);
  }
  if (lnNode.tailEnd) {
    const tailType = lnNode.tailEnd["@_type"] ?? "unknown";
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
    const color = colorResolver.resolve(lnNode.solidFill);
    if (color) {
      fill = { type: "solid", color };
    }
  }

  if (lnNode.noFill !== undefined) {
    return null;
  }

  const dashStyle = (lnNode.prstDash?.["@_val"] ?? "solid") as DashStyle;

  return { width, fill, dashStyle };
}
