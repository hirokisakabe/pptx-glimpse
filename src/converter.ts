import { readPptx } from "./parser/pptx-reader.js";
import { parsePresentation } from "./parser/presentation-parser.js";
import { parseTheme } from "./parser/theme-parser.js";
import { parseSlideMasterColorMap } from "./parser/slide-master-parser.js";
import { parseSlide } from "./parser/slide-parser.js";
import { parseRelationships, resolveRelationshipTarget } from "./parser/relationship-parser.js";
import { ColorResolver } from "./color/color-resolver.js";
import { renderSlideToSvg } from "./renderer/svg-renderer.js";
import { svgToPng } from "./png/png-converter.js";
import { DEFAULT_OUTPUT_WIDTH } from "./utils/constants.js";
import type { ColorMap } from "./model/theme.js";

export interface ConvertOptions {
  /** 変換対象のスライド番号 (1始まり)。未指定で全スライド */
  slides?: number[];
  /** 出力画像の幅 (ピクセル)。デフォルト: 960 */
  width?: number;
  /** 出力画像の高さ (ピクセル)。widthと同時指定時はwidthが優先 */
  height?: number;
}

export interface SlideSvg {
  slideNumber: number;
  svg: string;
}

export interface SlideImage {
  slideNumber: number;
  png: Buffer;
  width: number;
  height: number;
}

export async function convertPptxToSvg(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideSvg[]> {
  const archive = await readPptx(input);

  // Parse presentation.xml
  const presentationXml = archive.files.get("ppt/presentation.xml");
  if (!presentationXml) throw new Error("Invalid PPTX: missing ppt/presentation.xml");
  const presInfo = parsePresentation(presentationXml);

  // Parse presentation relationships
  const presRelsXml = archive.files.get("ppt/_rels/presentation.xml.rels");
  const presRels = presRelsXml ? parseRelationships(presRelsXml) : new Map();

  // Parse theme
  let theme = {
    colorScheme: defaultColorScheme(),
    fontScheme: { majorFont: "Calibri", minorFont: "Calibri" },
  };
  for (const [, rel] of presRels) {
    if (rel.type.includes("theme")) {
      const themePath = resolveRelationshipTarget("ppt/presentation.xml", rel.target);
      const themeXml = archive.files.get(themePath);
      if (themeXml) {
        theme = parseTheme(themeXml);
      }
      break;
    }
  }

  // Parse slide master for color map
  let colorMap: ColorMap = defaultColorMap();
  for (const [, rel] of presRels) {
    if (rel.type.includes("slideMaster")) {
      const masterPath = resolveRelationshipTarget("ppt/presentation.xml", rel.target);
      const masterXml = archive.files.get(masterPath);
      if (masterXml) {
        colorMap = parseSlideMasterColorMap(masterXml);
      }
      break;
    }
  }

  const colorResolver = new ColorResolver(theme.colorScheme, colorMap);

  // Resolve slide paths from relationships
  const slidePaths: { slideNumber: number; path: string }[] = [];
  for (let i = 0; i < presInfo.slideRIds.length; i++) {
    const rId = presInfo.slideRIds[i];
    const rel = presRels.get(rId);
    if (rel) {
      const path = resolveRelationshipTarget("ppt/presentation.xml", rel.target);
      slidePaths.push({ slideNumber: i + 1, path });
    }
  }

  // Filter slides if specified
  const targetSlides = options?.slides
    ? slidePaths.filter((s) => options.slides!.includes(s.slideNumber))
    : slidePaths;

  // Parse and render each slide
  const results: SlideSvg[] = [];
  for (const { slideNumber, path } of targetSlides) {
    const slideXml = archive.files.get(path);
    if (!slideXml) continue;

    const slide = parseSlide(slideXml, path, slideNumber, archive, colorResolver);
    const svg = renderSlideToSvg(slide, presInfo.slideSize);
    results.push({ slideNumber, svg });
  }

  return results;
}

export async function convertPptxToPng(
  input: Buffer | Uint8Array,
  options?: ConvertOptions,
): Promise<SlideImage[]> {
  const svgResults = await convertPptxToSvg(input, options);

  const width = options?.width ?? DEFAULT_OUTPUT_WIDTH;
  const height = options?.height;

  const results: SlideImage[] = [];
  for (const { slideNumber, svg } of svgResults) {
    const pngResult = await svgToPng(svg, { width, height });
    results.push({
      slideNumber,
      png: pngResult.png,
      width: pngResult.width,
      height: pngResult.height,
    });
  }

  return results;
}

function defaultColorScheme() {
  return {
    dk1: "#000000",
    lt1: "#FFFFFF",
    dk2: "#44546A",
    lt2: "#E7E6E6",
    accent1: "#4472C4",
    accent2: "#ED7D31",
    accent3: "#A5A5A5",
    accent4: "#FFC000",
    accent5: "#5B9BD5",
    accent6: "#70AD47",
    hlink: "#0563C1",
    folHlink: "#954F72",
  };
}

function defaultColorMap() {
  return {
    bg1: "lt1" as const,
    tx1: "dk1" as const,
    bg2: "lt2" as const,
    tx2: "dk2" as const,
    accent1: "accent1" as const,
    accent2: "accent2" as const,
    accent3: "accent3" as const,
    accent4: "accent4" as const,
    accent5: "accent5" as const,
    accent6: "accent6" as const,
    hlink: "hlink" as const,
    folHlink: "folHlink" as const,
  };
}
