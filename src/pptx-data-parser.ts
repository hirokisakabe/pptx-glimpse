/**
 * PPTX の共通パース処理。
 * converter.ts と font-collector.ts で共有する。
 */
import { readPptx, type PptxArchive } from "./parser/pptx-reader.js";
import { parsePresentation, type PresentationInfo } from "./parser/presentation-parser.js";
import { parseTheme } from "./parser/theme-parser.js";
import type { Theme, ColorMap } from "./model/theme.js";
import {
  parseSlideMasterColorMap,
  parseSlideMasterBackground,
  parseSlideMasterElements,
  parseSlideMasterTxStyles,
  parseSlideMasterPlaceholderStyles,
} from "./parser/slide-master-parser.js";
import {
  parseSlideLayoutBackground,
  parseSlideLayoutElements,
  parseSlideLayoutPlaceholderStyles,
  parseSlideLayoutShowMasterSp,
} from "./parser/slide-layout-parser.js";
import { parseSlide } from "./parser/slide-parser.js";
import type { Slide, Background } from "./model/slide.js";
import {
  parseRelationships,
  resolveRelationshipTarget,
  buildRelsPath,
  type Relationship,
} from "./parser/relationship-parser.js";
import type { FillParseContext } from "./parser/fill-parser.js";
import { ColorResolver } from "./color/color-resolver.js";
import type { PlaceholderStyleInfo, TxStyles } from "./model/text.js";
import type { SlideElement } from "./model/shape.js";
import { applyTextStyleInheritance } from "./text-style-resolver.js";

export interface ParsedPptxData {
  presInfo: PresentationInfo;
  theme: Theme;
  colorResolver: ColorResolver;
  masterBackground: Background | null;
  masterElements: SlideElement[];
  masterTxStyles: TxStyles | undefined;
  masterPlaceholderStyles: PlaceholderStyleInfo[];
  slidePaths: { slideNumber: number; path: string }[];
  archive: PptxArchive;
}

export function parsePptxData(input: Buffer | Uint8Array): ParsedPptxData {
  const archive = readPptx(input);

  // Parse presentation.xml
  const presentationXml = archive.files.get("ppt/presentation.xml");
  if (!presentationXml) throw new Error("Invalid PPTX: missing ppt/presentation.xml");
  const presInfo = parsePresentation(presentationXml);

  // Parse presentation relationships
  const presRelsXml = archive.files.get("ppt/_rels/presentation.xml.rels");
  const presRels = presRelsXml ? parseRelationships(presRelsXml) : new Map<string, Relationship>();

  // Parse theme
  let theme: Theme = {
    colorScheme: defaultColorScheme(),
    fontScheme: {
      majorFont: "Calibri",
      minorFont: "Calibri",
      majorFontEa: null,
      minorFontEa: null,
      majorFontCs: null,
      minorFontCs: null,
    },
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

  // Parse slide master for color map and background
  let colorMap: ColorMap = defaultColorMap();
  let masterPath: string | null = null;
  for (const [, rel] of presRels) {
    if (rel.type.includes("slideMaster")) {
      masterPath = resolveRelationshipTarget("ppt/presentation.xml", rel.target);
      const masterXml = archive.files.get(masterPath);
      if (masterXml) {
        colorMap = parseSlideMasterColorMap(masterXml);
      }
      break;
    }
  }

  const colorResolver = new ColorResolver(theme.colorScheme, colorMap);

  // Parse master background and shapes (used as fallback)
  const masterXml = masterPath ? archive.files.get(masterPath) : undefined;
  let masterFillContext: FillParseContext | undefined;
  if (masterPath) {
    const masterRelsPath = buildRelsPath(masterPath);
    const masterRelsXml = archive.files.get(masterRelsPath);
    const masterRels = masterRelsXml
      ? parseRelationships(masterRelsXml)
      : new Map<string, Relationship>();
    masterFillContext = { rels: masterRels, archive, basePath: masterPath };
  }
  const masterBackground = masterXml
    ? parseSlideMasterBackground(masterXml, colorResolver, masterFillContext)
    : null;
  const masterElements =
    masterPath && masterXml
      ? parseSlideMasterElements(
          masterXml,
          masterPath,
          archive,
          colorResolver,
          theme.fontScheme,
          theme.fmtScheme,
        )
      : [];
  const masterTxStyles = masterXml ? parseSlideMasterTxStyles(masterXml, colorResolver) : undefined;
  const masterPlaceholderStyles = masterXml
    ? parseSlideMasterPlaceholderStyles(masterXml, colorResolver)
    : [];

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

  return {
    presInfo,
    theme,
    colorResolver,
    masterBackground,
    masterElements,
    masterTxStyles,
    masterPlaceholderStyles,
    slidePaths,
    archive,
  };
}

export interface ParsedSlide {
  slide: Slide;
  layoutElements: SlideElement[];
  layoutShowMasterSp: boolean;
}

export function parseSlideWithLayout(
  slideNumber: number,
  path: string,
  data: ParsedPptxData,
): ParsedSlide | null {
  const slideXml = data.archive.files.get(path);
  if (!slideXml) return null;

  const slide = parseSlide(
    slideXml,
    path,
    slideNumber,
    data.archive,
    data.colorResolver,
    data.theme.fontScheme,
    data.theme.fmtScheme,
  );

  // Resolve slide layout
  let layoutElements: SlideElement[] = [];
  let layoutPlaceholderStyles: PlaceholderStyleInfo[] = [];
  let layoutShowMasterSp = true;
  const slideRelsPath = buildRelsPath(path);
  const slideRelsXml = data.archive.files.get(slideRelsPath);
  if (slideRelsXml) {
    const slideRels = parseRelationships(slideRelsXml);
    for (const [, rel] of slideRels) {
      if (rel.type.includes("slideLayout")) {
        const layoutPath = resolveRelationshipTarget(path, rel.target);
        const layoutXml = data.archive.files.get(layoutPath);
        if (layoutXml) {
          // Fallback background: slide → layout → master
          if (!slide.background) {
            const layoutRelsPath = buildRelsPath(layoutPath);
            const layoutRelsXml = data.archive.files.get(layoutRelsPath);
            const layoutRels = layoutRelsXml
              ? parseRelationships(layoutRelsXml)
              : new Map<string, Relationship>();
            const layoutFillContext: FillParseContext = {
              rels: layoutRels,
              archive: data.archive,
              basePath: layoutPath,
            };
            slide.background = parseSlideLayoutBackground(
              layoutXml,
              data.colorResolver,
              layoutFillContext,
            );
          }
          // Parse layout shapes
          layoutElements = parseSlideLayoutElements(
            layoutXml,
            layoutPath,
            data.archive,
            data.colorResolver,
            data.theme.fontScheme,
            data.theme.fmtScheme,
          );
          // Extract placeholder styles for text style inheritance
          layoutPlaceholderStyles = parseSlideLayoutPlaceholderStyles(
            layoutXml,
            data.colorResolver,
          );
          layoutShowMasterSp = parseSlideLayoutShowMasterSp(layoutXml);
        }
        break;
      }
    }
  }
  if (!slide.background) {
    slide.background = data.masterBackground;
  }

  // Apply text style inheritance chain before merging
  applyTextStyleInheritance(slide.elements, {
    layoutPlaceholderStyles,
    masterPlaceholderStyles: data.masterPlaceholderStyles,
    txStyles: data.masterTxStyles,
    defaultTextStyle: data.presInfo.defaultTextStyle,
    fontScheme: data.theme.fontScheme,
  });

  return { slide, layoutElements, layoutShowMasterSp };
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
