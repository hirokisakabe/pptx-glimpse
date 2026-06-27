// OOXML テーマカラー解決 (ECMA-376 §20.1.6)
// 色の種類:
//   srgbClr — sRGB 直接指定 (e.g. "FF0000")
//   schemeClr — テーマカラー参照 (dk1, lt1, accent1-6, hlink, folHlink)
//     → colorMap でキーを変換 → colorScheme から実際の色を取得
//   sysClr — システムカラー (@_lastClr で保存された直近値を使用)

import type { ResolvedColor } from "@pptx-glimpse/renderer";
import type { ColorMap, ColorScheme, ColorSchemeKey } from "@pptx-glimpse/renderer";
import { debug } from "@pptx-glimpse/renderer";

import type { XmlNode } from "../parser/xml-parser.js";
import { unsafeTypeAssertion } from "../unsafe-type-assertion.js";
import { applyColorTransforms } from "./color-transforms.js";

export class ColorResolver {
  constructor(
    private colorScheme: ColorScheme,
    private colorMap: ColorMap,
  ) {}

  resolve(colorNode: XmlNode): ResolvedColor | null {
    if (!colorNode) return null;

    if (colorNode.srgbClr) {
      return this.resolveSrgbClr(unsafeTypeAssertion<XmlNode>(colorNode.srgbClr));
    }
    if (colorNode.schemeClr) {
      return this.resolveSchemeClr(unsafeTypeAssertion<XmlNode>(colorNode.schemeClr));
    }
    if (colorNode.sysClr) {
      return this.resolveSysClr(unsafeTypeAssertion<XmlNode>(colorNode.sysClr));
    }

    const keys = Object.keys(colorNode).filter((k) => !k.startsWith("@_"));
    if (keys.length > 0) {
      debug("colorResolver.unknown", `unknown color node structure [${keys.join(", ")}]`);
    }

    return null;
  }

  private resolveSrgbClr(node: XmlNode): ResolvedColor {
    const hex = `#${unsafeTypeAssertion<string>(node["@_val"])}`;
    const alpha = extractAlpha(node);
    return applyColorTransforms({ hex, alpha }, node);
  }

  private resolveSchemeClr(node: XmlNode): ResolvedColor {
    const schemeName = unsafeTypeAssertion<string>(node["@_val"]);
    const hex = this.resolveSchemeColorName(schemeName);
    const alpha = extractAlpha(node);
    return applyColorTransforms({ hex, alpha }, node);
  }

  private resolveSysClr(node: XmlNode): ResolvedColor {
    const hex = `#${unsafeTypeAssertion<string | undefined>(node["@_lastClr"]) ?? "000000"}`;
    const alpha = extractAlpha(node);
    return applyColorTransforms({ hex, alpha }, node);
  }

  private resolveSchemeColorName(name: string): string {
    const mapped = this.mapColorName(name);
    return this.colorScheme[mapped] ?? "#000000";
  }

  private mapColorName(name: string): ColorSchemeKey {
    if (name in this.colorMap) {
      return this.colorMap[unsafeTypeAssertion<keyof ColorMap>(name)];
    }
    if (name in this.colorScheme) {
      return unsafeTypeAssertion<ColorSchemeKey>(name);
    }
    return "dk1";
  }
}

function extractAlpha(node: XmlNode): number {
  const alphaNode = unsafeTypeAssertion<XmlNode | undefined>(node.alpha);
  if (alphaNode) {
    return Number(alphaNode["@_val"]) / 100000;
  }
  return 1;
}
