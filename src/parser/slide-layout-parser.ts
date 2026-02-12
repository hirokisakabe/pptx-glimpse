import type { Background } from "../model/slide.js";
import type { SlideElement } from "../model/shape.js";
import type { PptxArchive } from "./pptx-reader.js";
import { parseXml } from "./xml-parser.js";
import { parseFillFromNode } from "./fill-parser.js";
import type { FillParseContext } from "./fill-parser.js";
import { parseShapeTree } from "./slide-parser.js";
import { buildRelsPath, parseRelationships } from "./relationship-parser.js";
import type { ColorResolver } from "../color/color-resolver.js";
import type { FontScheme } from "../model/theme.js";

const WARN_PREFIX = "[pptx-glimpse]";

export function parseSlideLayoutBackground(
  xml: string,
  colorResolver: ColorResolver,
  context?: FillParseContext,
): Background | null {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;

  if (!parsed.sldLayout) {
    console.warn(`${WARN_PREFIX} SlideLayout: missing root element "sldLayout" in XML`);
    return null;
  }

  const bg = parsed.sldLayout.cSld?.bg;
  if (!bg) return null;

  const bgPr = bg.bgPr;
  if (!bgPr) return null;

  const fill = parseFillFromNode(bgPr, colorResolver, context);
  return { fill };
}

export function parseSlideLayoutElements(
  xml: string,
  layoutPath: string,
  archive: PptxArchive,
  colorResolver: ColorResolver,
  fontScheme?: FontScheme | null,
): SlideElement[] {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;

  if (!parsed.sldLayout) {
    console.warn(`${WARN_PREFIX} SlideLayout: missing root element "sldLayout" in XML`);
    return [];
  }

  const spTree = parsed.sldLayout.cSld?.spTree;
  if (!spTree) return [];

  const relsPath = buildRelsPath(layoutPath);
  const relsXml = archive.files.get(relsPath);
  const rels = relsXml ? parseRelationships(relsXml) : new Map();

  return parseShapeTree(
    spTree,
    rels,
    layoutPath,
    archive,
    colorResolver,
    undefined,
    undefined,
    fontScheme,
  );
}
