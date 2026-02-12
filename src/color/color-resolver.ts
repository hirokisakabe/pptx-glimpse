// OOXML テーマカラー解決 (ECMA-376 §20.1.6)
// 色の種類:
//   srgbClr — sRGB 直接指定 (e.g. "FF0000")
//   schemeClr — テーマカラー参照 (dk1, lt1, accent1-6, hlink, folHlink)
//     → colorMap でキーを変換 → colorScheme から実際の色を取得
//   sysClr — システムカラー (@_lastClr で保存された直近値を使用)

import type { ColorScheme, ColorMap, ColorSchemeKey } from "../model/theme.js";
import type { ResolvedColor } from "../model/fill.js";
import type { XmlNode } from "../parser/xml-parser.js";
import { applyColorTransforms } from "./color-transforms.js";

const WARN_PREFIX = "[pptx-glimpse]";

export class ColorResolver {
  constructor(
    private colorScheme: ColorScheme,
    private colorMap: ColorMap,
  ) {}

  resolve(colorNode: XmlNode): ResolvedColor | null {
    if (!colorNode) return null;

    if (colorNode.srgbClr) {
      return this.resolveSrgbClr(colorNode.srgbClr as XmlNode);
    }
    if (colorNode.schemeClr) {
      return this.resolveSchemeClr(colorNode.schemeClr as XmlNode);
    }
    if (colorNode.sysClr) {
      return this.resolveSysClr(colorNode.sysClr as XmlNode);
    }

    const keys = Object.keys(colorNode).filter((k) => !k.startsWith("@_"));
    if (keys.length > 0) {
      console.warn(
        `${WARN_PREFIX} ColorResolver: unknown color node structure [${keys.join(", ")}]`,
      );
    }

    return null;
  }

  private resolveSrgbClr(node: XmlNode): ResolvedColor {
    const hex = `#${node["@_val"] as string}`;
    const alpha = extractAlpha(node);
    return applyColorTransforms({ hex, alpha }, node);
  }

  private resolveSchemeClr(node: XmlNode): ResolvedColor {
    const schemeName = node["@_val"] as string;
    const hex = this.resolveSchemeColorName(schemeName);
    const alpha = extractAlpha(node);
    return applyColorTransforms({ hex, alpha }, node);
  }

  private resolveSysClr(node: XmlNode): ResolvedColor {
    const hex = `#${(node["@_lastClr"] as string | undefined) ?? "000000"}`;
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

function extractAlpha(node: XmlNode): number {
  const alphaNode = node.alpha as XmlNode | undefined;
  if (alphaNode) {
    return Number(alphaNode["@_val"]) / 100000;
  }
  return 1;
}
