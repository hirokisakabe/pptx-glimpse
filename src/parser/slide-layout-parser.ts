import type { Background } from "../model/slide.js";
import { parseXml } from "./xml-parser.js";
import { parseFillFromNode } from "./fill-parser.js";
import type { ColorResolver } from "../color/color-resolver.js";

export function parseSlideLayoutBackground(
  xml: string,
  colorResolver: ColorResolver,
): Background | null {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;
  const bg = parsed.sldLayout?.cSld?.bg;
  if (!bg) return null;

  const bgPr = bg.bgPr;
  if (!bgPr) return null;

  const fill = parseFillFromNode(bgPr, colorResolver);
  return { fill };
}
