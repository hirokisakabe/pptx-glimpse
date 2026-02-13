import type { FormatScheme } from "../model/theme.js";
import type { Fill } from "../model/fill.js";
import type { ResolvedColor } from "../model/fill.js";
import type { Outline } from "../model/line.js";
import type { EffectList } from "../model/effect.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { XmlNode } from "./xml-parser.js";

export interface ResolvedStyleReference {
  fill: Fill | null;
  outline: Outline | null;
  effects: EffectList | null;
  fontRef?: { idx: string; color: ResolvedColor | null };
}

export function resolveShapeStyle(
  styleNode: XmlNode | undefined,
  fmtScheme: FormatScheme | undefined,
  colorResolver: ColorResolver,
): ResolvedStyleReference | null {
  if (!styleNode || !fmtScheme) return null;

  const fillRef = styleNode.fillRef as XmlNode | undefined;
  const lnRef = styleNode.lnRef as XmlNode | undefined;
  const effectRef = styleNode.effectRef as XmlNode | undefined;
  const fontRef = styleNode.fontRef as XmlNode | undefined;

  const fill = resolveFillRef(fillRef, fmtScheme, colorResolver);
  const outline = resolveLineRef(lnRef, fmtScheme, colorResolver);
  const effects = resolveEffectRef(effectRef, fmtScheme);

  let fontRefResult: { idx: string; color: ResolvedColor | null } | undefined;
  if (fontRef) {
    const idx = (fontRef["@_idx"] as string | undefined) ?? "minor";
    const color = colorResolver.resolve(fontRef);
    fontRefResult = { idx, color };
  }

  return { fill, outline, effects, fontRef: fontRefResult };
}

function resolveFillRef(
  ref: XmlNode | undefined,
  fmtScheme: FormatScheme,
  colorResolver: ColorResolver,
): Fill | null {
  if (!ref) return null;
  const idx = Number(ref["@_idx"] ?? 0);
  if (idx === 0) return null;

  // idx >= 1000 references bgFillStyleLst
  const list = idx >= 1000 ? fmtScheme.bgFillStyles : fmtScheme.fillStyles;
  const arrayIdx = idx >= 1000 ? idx - 1001 : idx - 1;
  const templateFill = list[arrayIdx];
  if (!templateFill) return null;

  // Override the template fill color with the ref's child color
  const overrideColor = colorResolver.resolve(ref);
  if (overrideColor && templateFill.type === "solid") {
    return { type: "solid", color: overrideColor };
  }
  if (overrideColor && templateFill.type === "gradient") {
    // For gradients, override all stop colors
    return {
      ...templateFill,
      stops: templateFill.stops.map((s) => ({ ...s, color: overrideColor })),
    };
  }

  return templateFill;
}

function resolveLineRef(
  ref: XmlNode | undefined,
  fmtScheme: FormatScheme,
  colorResolver: ColorResolver,
): Outline | null {
  if (!ref) return null;
  const idx = Number(ref["@_idx"] ?? 0);
  if (idx === 0) return null;

  const templateOutline = fmtScheme.lnStyles[idx - 1];
  if (!templateOutline) return null;

  const overrideColor = colorResolver.resolve(ref);
  if (overrideColor) {
    return {
      ...templateOutline,
      fill: { type: "solid", color: overrideColor },
    };
  }

  return templateOutline;
}

function resolveEffectRef(ref: XmlNode | undefined, fmtScheme: FormatScheme): EffectList | null {
  if (!ref) return null;
  const idx = Number(ref["@_idx"] ?? 0);
  // effectStyleLst uses 0-based indexing in the array
  return fmtScheme.effectStyles[idx] ?? null;
}
