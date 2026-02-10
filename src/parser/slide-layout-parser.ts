import type { Background } from "../model/slide.js";
import type { SlideElement } from "../model/shape.js";
import type { PptxArchive } from "./pptx-reader.js";
import { parseXml } from "./xml-parser.js";
import { parseFillFromNode } from "./fill-parser.js";
import { parseShapeTree } from "./slide-parser.js";
import { buildRelsPath, parseRelationships } from "./relationship-parser.js";
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

export function parseSlideLayoutElements(
  xml: string,
  layoutPath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
): SlideElement[] {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;
  const spTree = parsed.sldLayout?.cSld?.spTree;
  if (!spTree) return [];

  const relsPath = buildRelsPath(layoutPath);
  const relsXml = archive.files.get(relsPath);
  const rels = relsXml ? parseRelationships(relsXml) : new Map();

  return parseShapeTree(spTree, rels, layoutPath, archive, colorResolver);
}
