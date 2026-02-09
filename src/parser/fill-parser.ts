import type { Fill, GradientStop, SolidFill } from "../model/fill.js";
import type { Outline, DashStyle } from "../model/line.js";
import type { ColorResolver } from "../color/color-resolver.js";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function parseFillFromNode(node: any, colorResolver: ColorResolver): Fill | null {
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

  return null;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseGradientFill(gradNode: any, colorResolver: ColorResolver): Fill | null {
  const gsLst = gradNode.gsLst?.gs;
  if (!gsLst) return null;

  const stops: GradientStop[] = [];
  for (const gs of gsLst) {
    const position = Number(gs["@_pos"] ?? 0) / 100000;
    const color = colorResolver.resolve(gs);
    if (color) {
      stops.push({ position, color });
    }
  }

  let angle = 0;
  if (gradNode.lin) {
    angle = Number(gradNode.lin["@_ang"] ?? 0) / 60000;
  }

  return { type: "gradient", stops, angle };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function parseOutline(lnNode: any, colorResolver: ColorResolver): Outline | null {
  if (!lnNode) return null;

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
