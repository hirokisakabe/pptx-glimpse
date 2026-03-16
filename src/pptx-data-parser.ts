/**
 * PPTX の共通パース処理。
 * converter.ts と font-collector.ts で共有する。
 */
import { ColorResolver } from "./color/color-resolver.js";
import type { SlideElement } from "./model/shape.js";
import type { Background, Slide } from "./model/slide.js";
import type { PlaceholderStyleInfo, TxStyles } from "./model/text.js";
import type { ColorMap, Theme } from "./model/theme.js";
import type { FillParseContext } from "./parser/fill-parser.js";
import { type PptxArchive, readPptx } from "./parser/pptx-reader.js";
import { parsePresentation, type PresentationInfo } from "./parser/presentation-parser.js";
import {
  buildRelsPath,
  parseRelationships,
  type Relationship,
  resolveRelationshipTarget,
} from "./parser/relationship-parser.js";
import {
  parseSlideLayoutBackground,
  parseSlideLayoutElements,
  parseSlideLayoutPlaceholderStyles,
  parseSlideLayoutShowMasterSp,
} from "./parser/slide-layout-parser.js";
import {
  parseSlideMasterBackground,
  parseSlideMasterColorMap,
  parseSlideMasterElements,
  parseSlideMasterPlaceholderStyles,
  parseSlideMasterTxStyles,
} from "./parser/slide-master-parser.js";
import { parseSlide } from "./parser/slide-parser.js";
import { parseTheme } from "./parser/theme-parser.js";
import { parseXml, type XmlNode } from "./parser/xml-parser.js";
import { applyTextStyleInheritance } from "./text-style-resolver.js";

interface MasterData {
  colorMap: ColorMap;
  colorResolver: ColorResolver;
  background: Background | null;
  elements: SlideElement[];
  txStyles: TxStyles | undefined;
  placeholderStyles: PlaceholderStyleInfo[];
}

interface ParsedPptxData {
  presInfo: PresentationInfo;
  theme: Theme;
  /** @deprecated Use per-slide master data instead. Kept for compatibility. */
  colorResolver: ColorResolver;
  masterBackground: Background | null;
  masterElements: SlideElement[];
  masterTxStyles: TxStyles | undefined;
  masterPlaceholderStyles: PlaceholderStyleInfo[];
  slidePaths: { slideNumber: number; path: string }[];
  archive: PptxArchive;
  /** Cache of parsed master data keyed by master file path */
  masterCache: Map<string, MasterData>;
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
      majorFontJpan: null,
      minorFontJpan: null,
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

  // Parse all slide masters and cache their data
  const masterCache = new Map<string, MasterData>();
  let firstMasterPath: string | null = null;
  for (const [, rel] of presRels) {
    if (rel.type.includes("slideMaster")) {
      const mPath = resolveRelationshipTarget("ppt/presentation.xml", rel.target);
      if (!firstMasterPath) firstMasterPath = mPath;
      parseMasterDataCached(mPath, archive, theme, masterCache);
    }
  }

  // Use first master as default (for backward compatibility)
  const defaultMaster = firstMasterPath ? masterCache.get(firstMasterPath) : undefined;
  const colorResolver =
    defaultMaster?.colorResolver ?? new ColorResolver(theme.colorScheme, defaultColorMap());
  const masterBackground = defaultMaster?.background ?? null;
  const masterElements = defaultMaster?.elements ?? [];
  const masterTxStyles = defaultMaster?.txStyles;
  const masterPlaceholderStyles = defaultMaster?.placeholderStyles ?? [];

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
    masterCache,
  };
}

interface ParsedSlide {
  slide: Slide;
  layoutElements: SlideElement[];
  layoutShowMasterSp: boolean;
  masterElements: SlideElement[];
}

export function parseSlideWithLayout(
  slideNumber: number,
  path: string,
  data: ParsedPptxData,
): ParsedSlide | null {
  const slideXml = data.archive.files.get(path);
  if (!slideXml) return null;

  // Resolve slide → layout → master chain to get correct ColorResolver
  const slideRelsPath = buildRelsPath(path);
  const slideRelsXml = data.archive.files.get(slideRelsPath);
  const slideRels = slideRelsXml
    ? parseRelationships(slideRelsXml)
    : new Map<string, Relationship>();

  let layoutPath: string | null = null;
  let layoutXml: string | undefined;
  let layoutRels = new Map<string, Relationship>();

  for (const [, rel] of slideRels) {
    if (rel.type.includes("slideLayout")) {
      layoutPath = resolveRelationshipTarget(path, rel.target);
      layoutXml = data.archive.files.get(layoutPath);
      if (layoutXml) {
        const layoutRelsPath = buildRelsPath(layoutPath);
        const layoutRelsXml = data.archive.files.get(layoutRelsPath);
        layoutRels = layoutRelsXml
          ? parseRelationships(layoutRelsXml)
          : new Map<string, Relationship>();
      }
      break;
    }
  }

  // Resolve the correct slide master from layout → master chain
  let slideMasterData: MasterData | undefined;
  for (const [, rel] of layoutRels) {
    if (rel.type.includes("slideMaster") && layoutPath) {
      const masterPath = resolveRelationshipTarget(layoutPath, rel.target);
      slideMasterData = parseMasterDataCached(
        masterPath,
        data.archive,
        data.theme,
        data.masterCache,
      );
      break;
    }
  }

  // Apply clrMapOvr (color map override) from slide/layout if present
  const slideColorResolver = resolveSlideColorResolver(
    slideXml,
    layoutXml,
    slideMasterData,
    data.theme,
  );

  const slide = parseSlide(
    slideXml,
    path,
    slideNumber,
    data.archive,
    slideColorResolver,
    data.theme.fontScheme,
    data.theme.fmtScheme,
  );

  // Resolve slide layout
  let layoutElements: SlideElement[] = [];
  let layoutPlaceholderStyles: PlaceholderStyleInfo[] = [];
  let layoutShowMasterSp = true;
  if (layoutXml && layoutPath) {
    // Fallback background: slide → layout → master
    if (!slide.background) {
      const layoutFillContext: FillParseContext = {
        rels: layoutRels,
        archive: data.archive,
        basePath: layoutPath,
      };
      slide.background = parseSlideLayoutBackground(
        layoutXml,
        slideColorResolver,
        layoutFillContext,
      );
    }
    // Parse layout shapes
    layoutElements = parseSlideLayoutElements(
      layoutXml,
      layoutPath,
      data.archive,
      slideColorResolver,
      data.theme.fontScheme,
      data.theme.fmtScheme,
    );
    // Extract placeholder styles for text style inheritance
    layoutPlaceholderStyles = parseSlideLayoutPlaceholderStyles(layoutXml, slideColorResolver);
    layoutShowMasterSp = parseSlideLayoutShowMasterSp(layoutXml);
  }
  if (!slide.background) {
    slide.background = slideMasterData?.background ?? data.masterBackground;
  }

  const masterPlaceholderStyles =
    slideMasterData?.placeholderStyles ?? data.masterPlaceholderStyles;
  const masterTxStyles = slideMasterData?.txStyles ?? data.masterTxStyles;

  // Apply text style inheritance chain before merging
  applyTextStyleInheritance(slide.elements, {
    layoutPlaceholderStyles,
    masterPlaceholderStyles,
    txStyles: masterTxStyles,
    defaultTextStyle: data.presInfo.defaultTextStyle,
    fontScheme: data.theme.fontScheme,
  });

  const masterElements = slideMasterData?.elements ?? data.masterElements;

  return { slide, layoutElements, layoutShowMasterSp, masterElements };
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

function defaultColorMap(): ColorMap {
  return {
    bg1: "lt1",
    tx1: "dk1",
    bg2: "lt2",
    tx2: "dk2",
    accent1: "accent1",
    accent2: "accent2",
    accent3: "accent3",
    accent4: "accent4",
    accent5: "accent5",
    accent6: "accent6",
    hlink: "hlink",
    folHlink: "folHlink",
  };
}

function parseMasterDataCached(
  masterPath: string,
  archive: PptxArchive,
  theme: Theme,
  cache: Map<string, MasterData>,
): MasterData | undefined {
  const cached = cache.get(masterPath);
  if (cached) return cached;

  const masterXml = archive.files.get(masterPath);
  if (!masterXml) return undefined;

  const colorMap = parseSlideMasterColorMap(masterXml);
  const colorResolver = new ColorResolver(theme.colorScheme, colorMap);

  const masterRelsPath = buildRelsPath(masterPath);
  const masterRelsXml = archive.files.get(masterRelsPath);
  const masterRels = masterRelsXml
    ? parseRelationships(masterRelsXml)
    : new Map<string, Relationship>();
  const masterFillContext: FillParseContext = { rels: masterRels, archive, basePath: masterPath };

  const background = parseSlideMasterBackground(masterXml, colorResolver, masterFillContext);
  const elements = parseSlideMasterElements(
    masterXml,
    masterPath,
    archive,
    colorResolver,
    theme.fontScheme,
    theme.fmtScheme,
  );
  const txStyles = parseSlideMasterTxStyles(masterXml, colorResolver);
  const placeholderStyles = parseSlideMasterPlaceholderStyles(masterXml, colorResolver);

  const data: MasterData = {
    colorMap,
    colorResolver,
    background,
    elements,
    txStyles,
    placeholderStyles,
  };
  cache.set(masterPath, data);
  return data;
}

function resolveSlideColorResolver(
  slideXml: string,
  layoutXml: string | undefined,
  masterData: MasterData | undefined,
  theme: Theme,
): ColorResolver {
  const baseColorMap = masterData?.colorMap ?? defaultColorMap();

  // Check layout clrMapOvr first, then slide clrMapOvr
  const layoutOverride = layoutXml ? parseClrMapOverride(layoutXml) : null;
  const slideOverride = parseClrMapOverride(slideXml);

  // Apply overrides: layout override replaces master, slide override replaces layout
  let effectiveColorMap = baseColorMap;
  if (layoutOverride) {
    effectiveColorMap = { ...effectiveColorMap, ...layoutOverride };
  }
  if (slideOverride) {
    effectiveColorMap = { ...effectiveColorMap, ...slideOverride };
  }

  // If no overrides, reuse master's colorResolver
  if (!layoutOverride && !slideOverride && masterData) {
    return masterData.colorResolver;
  }

  return new ColorResolver(theme.colorScheme, effectiveColorMap);
}

function parseClrMapOverride(xml: string): Partial<ColorMap> | null {
  // Quick check to avoid parsing XML unnecessarily
  if (!xml.includes("clrMapOvr")) return null;

  const parsed = parseXml(xml);

  // Navigate to root element (sld or sldLayout)
  const root = (parsed.sld ?? parsed.sldLayout) as XmlNode | undefined;
  if (!root) return null;

  const clrMapOvr = root.clrMapOvr as XmlNode | undefined;
  if (!clrMapOvr) return null;

  // masterClrMapping = use master as-is (no override)
  if (clrMapOvr.masterClrMapping !== undefined) return null;

  // overrideClrMapping = individual attribute overrides
  const override = clrMapOvr.overrideClrMapping as XmlNode | undefined;
  if (!override) return null;

  const result: Partial<ColorMap> = {};
  const keys: (keyof ColorMap)[] = [
    "bg1",
    "tx1",
    "bg2",
    "tx2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
  ];
  for (const key of keys) {
    const val = override[`@_${key}`] as string | undefined;
    if (val) {
      (result as Record<string, string>)[key] = val;
    }
  }
  return Object.keys(result).length > 0 ? result : null;
}
