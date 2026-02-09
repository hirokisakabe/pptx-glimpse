import type { ColorScheme, ColorMap, ColorSchemeKey } from "../model/theme.js";
import type { ResolvedColor } from "../model/fill.js";
import { applyColorTransforms } from "./color-transforms.js";

export class ColorResolver {
  constructor(
    private colorScheme: ColorScheme,
    private colorMap: ColorMap,
  ) {}

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  resolve(colorNode: any): ResolvedColor | null {
    if (!colorNode) return null;

    if (colorNode.srgbClr) {
      return this.resolveSrgbClr(colorNode.srgbClr);
    }
    if (colorNode.schemeClr) {
      return this.resolveSchemeClr(colorNode.schemeClr);
    }
    if (colorNode.sysClr) {
      return this.resolveSysClr(colorNode.sysClr);
    }

    return null;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private resolveSrgbClr(node: any): ResolvedColor {
    const hex = `#${node["@_val"]}`;
    const alpha = extractAlpha(node);
    return applyColorTransforms({ hex, alpha }, node);
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private resolveSchemeClr(node: any): ResolvedColor {
    const schemeName = node["@_val"] as string;
    const hex = this.resolveSchemeColorName(schemeName);
    const alpha = extractAlpha(node);
    return applyColorTransforms({ hex, alpha }, node);
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private resolveSysClr(node: any): ResolvedColor {
    const hex = `#${node["@_lastClr"] ?? "000000"}`;
    const alpha = extractAlpha(node);
    return applyColorTransforms({ hex, alpha }, node);
  }

  private resolveSchemeColorName(name: string): string {
    const mapped = this.mapColorName(name);
    return this.colorScheme[mapped] ?? "#000000";
  }

  private mapColorName(name: string): ColorSchemeKey {
    if (name in this.colorMap) {
      return this.colorMap[name as keyof ColorMap];
    }
    if (name in this.colorScheme) {
      return name as ColorSchemeKey;
    }
    return "dk1";
  }
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractAlpha(node: any): number {
  const alphaNode = node.alpha;
  if (alphaNode) {
    return Number(alphaNode["@_val"]) / 100000;
  }
  return 1;
}
