/**
 * From-scratch PptxSourceModel factory.
 *
 * MVP scope: create a valid one-slide presentation with the minimum OOXML skeleton
 * needed by the existing writer path: presentation, slide, blank layout, slide master,
 * theme, content types, and relationships. Text authoring is intentionally delegated to
 * the existing `addTextBox` editing API, so the first supported flow is
 * `createPptx()` -> `addTextBox(...)` -> `writePptx(...)`.
 *
 * API choice: this module exposes a factory-style builder entry point (`createPptx`)
 * instead of asking callers to hand-build `PptxSourceModel`. The source model requires
 * package parts, relationship ids, part paths, content types, raw writer material, and
 * typed slide/layout/master/theme references to stay consistent. Keeping those
 * invariants inside a factory prevents easy creation of broken packages. The alternative
 * considered for this issue was exporting minimal helpers for direct `PptxSourceModel`
 * construction and adding a higher-level builder later; that was not chosen because it
 * would expose the same cross-part invariants to early consumers while the MVP only
 * needs a stable empty-presentation seed for text-box generation.
 */

import type {
  ContentTypeOverride,
  PackagePartRef,
  PartPath,
  PartRelationships,
  PptxSourceModel,
  RawPackagePart,
  Relationship,
  SlideSize,
  SourceColorMap,
  SourceTheme,
} from "../source/index.js";
import { asEmu, asPartPath, asRelationshipId } from "../source/index.js";

export interface CreatePptxOptions {
  readonly slideSize?: SlideSize;
}

const textEncoder = new TextEncoder();

const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const XML_CONTENT_TYPE = "application/xml";
const PRESENTATION_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const SLIDE_LAYOUT_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const SLIDE_MASTER_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
const THEME_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml";
const APP_PROPS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.extended-properties+xml";
const CORE_PROPS_CONTENT_TYPE = "application/vnd.openxmlformats-package.core-properties+xml";

const ROOT_PART = asPartPath("");
const PRESENTATION_PART = asPartPath("ppt/presentation.xml");
const SLIDE_PART = asPartPath("ppt/slides/slide1.xml");
const SLIDE_LAYOUT_PART = asPartPath("ppt/slideLayouts/slideLayout1.xml");
const SLIDE_MASTER_PART = asPartPath("ppt/slideMasters/slideMaster1.xml");
const THEME_PART = asPartPath("ppt/theme/theme1.xml");
const APP_PROPS_PART = asPartPath("docProps/app.xml");
const CORE_PROPS_PART = asPartPath("docProps/core.xml");

const DEFAULT_SLIDE_SIZE: SlideSize = {
  width: asEmu(9144000),
  height: asEmu(5143500),
};

const DEFAULT_COLOR_MAP: SourceColorMap = {
  mapping: {
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
  },
};

const THEME_SOURCE: SourceTheme = {
  partPath: THEME_PART,
  name: "Office Theme",
  colorScheme: {
    colors: {
      dk1: { kind: "system", value: "windowText", lastColor: "000000" },
      lt1: { kind: "system", value: "window", lastColor: "FFFFFF" },
      dk2: { kind: "srgb", hex: "44546A" },
      lt2: { kind: "srgb", hex: "E7E6E6" },
      accent1: { kind: "srgb", hex: "4472C4" },
      accent2: { kind: "srgb", hex: "ED7D31" },
      accent3: { kind: "srgb", hex: "A5A5A5" },
      accent4: { kind: "srgb", hex: "FFC000" },
      accent5: { kind: "srgb", hex: "5B9BD5" },
      accent6: { kind: "srgb", hex: "70AD47" },
      hlink: { kind: "srgb", hex: "0563C1" },
      folHlink: { kind: "srgb", hex: "954F72" },
    },
  },
  fontScheme: {
    majorLatin: "Aptos Display",
    minorLatin: "Aptos",
    majorEastAsian: "",
    minorEastAsian: "",
    majorComplexScript: "",
    minorComplexScript: "",
  },
  handle: { partPath: THEME_PART },
};

export function createPptx(options: CreatePptxOptions = {}): PptxSourceModel {
  const slideSize = options.slideSize ?? DEFAULT_SLIDE_SIZE;
  const overrides = createContentTypeOverrides();
  const rawParts = createRawParts(slideSize);
  const parts = createPackageParts(overrides);

  return {
    packageGraph: {
      contentTypes: {
        defaults: [
          { extension: "rels", contentType: RELS_CONTENT_TYPE },
          { extension: "xml", contentType: XML_CONTENT_TYPE },
        ],
        overrides,
      },
      parts,
      relationships: createRelationships(),
      media: [],
      rawParts,
    },
    presentation: {
      partPath: PRESENTATION_PART,
      slideSize,
      slidePartPaths: [SLIDE_PART],
      handle: { partPath: PRESENTATION_PART },
    },
    slides: [
      {
        partPath: SLIDE_PART,
        layoutPartPath: SLIDE_LAYOUT_PART,
        shapes: [],
        handle: { partPath: SLIDE_PART },
      },
    ],
    slideLayouts: [
      {
        partPath: SLIDE_LAYOUT_PART,
        masterPartPath: SLIDE_MASTER_PART,
        type: "blank",
        show: true,
        shapes: [],
        handle: { partPath: SLIDE_LAYOUT_PART },
      },
    ],
    slideMasters: [
      {
        partPath: SLIDE_MASTER_PART,
        themePartPath: THEME_PART,
        layoutPartPaths: [SLIDE_LAYOUT_PART],
        colorMap: DEFAULT_COLOR_MAP,
        shapes: [],
        handle: { partPath: SLIDE_MASTER_PART },
      },
    ],
    themes: [THEME_SOURCE],
    diagnostics: [],
  };
}

function createContentTypeOverrides(): ContentTypeOverride[] {
  return [
    { partName: PRESENTATION_PART, contentType: PRESENTATION_CONTENT_TYPE },
    { partName: SLIDE_PART, contentType: SLIDE_CONTENT_TYPE },
    { partName: SLIDE_LAYOUT_PART, contentType: SLIDE_LAYOUT_CONTENT_TYPE },
    { partName: SLIDE_MASTER_PART, contentType: SLIDE_MASTER_CONTENT_TYPE },
    { partName: THEME_PART, contentType: THEME_CONTENT_TYPE },
    { partName: APP_PROPS_PART, contentType: APP_PROPS_CONTENT_TYPE },
    { partName: CORE_PROPS_PART, contentType: CORE_PROPS_CONTENT_TYPE },
  ];
}

function createPackageParts(overrides: readonly ContentTypeOverride[]): PackagePartRef[] {
  const partRefs = overrides.map((override) => ({
    partPath: override.partName,
    contentType: override.contentType,
  }));
  return [
    ...partRefs,
    { partPath: asPartPath("_rels/.rels"), contentType: RELS_CONTENT_TYPE },
    { partPath: asPartPath("ppt/_rels/presentation.xml.rels"), contentType: RELS_CONTENT_TYPE },
    { partPath: asPartPath("ppt/slides/_rels/slide1.xml.rels"), contentType: RELS_CONTENT_TYPE },
    {
      partPath: asPartPath("ppt/slideLayouts/_rels/slideLayout1.xml.rels"),
      contentType: RELS_CONTENT_TYPE,
    },
    {
      partPath: asPartPath("ppt/slideMasters/_rels/slideMaster1.xml.rels"),
      contentType: RELS_CONTENT_TYPE,
    },
  ];
}

function createRelationships(): PartRelationships[] {
  return [
    {
      sourcePartPath: ROOT_PART,
      relationships: [
        rel(
          "rId1",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
          "ppt/presentation.xml",
        ),
        rel(
          "rId2",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
          "docProps/app.xml",
        ),
        rel(
          "rId3",
          "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
          "docProps/core.xml",
        ),
      ],
    },
    {
      sourcePartPath: PRESENTATION_PART,
      relationships: [
        rel(
          "rId1",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
          "slideMasters/slideMaster1.xml",
        ),
        rel(
          "rId2",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
          "slides/slide1.xml",
        ),
      ],
    },
    {
      sourcePartPath: SLIDE_PART,
      relationships: [
        rel(
          "rId1",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
          "../slideLayouts/slideLayout1.xml",
        ),
      ],
    },
    {
      sourcePartPath: SLIDE_LAYOUT_PART,
      relationships: [
        rel(
          "rId1",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
          "../slideMasters/slideMaster1.xml",
        ),
      ],
    },
    {
      sourcePartPath: SLIDE_MASTER_PART,
      relationships: [
        rel(
          "rId1",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
          "../slideLayouts/slideLayout1.xml",
        ),
        rel(
          "rId2",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
          "../theme/theme1.xml",
        ),
      ],
    },
  ];
}

function rel(id: string, type: string, target: string): Relationship {
  return { id: asRelationshipId(id), type, target };
}

function createRawParts(slideSize: SlideSize): RawPackagePart[] {
  return [
    rawXml(PRESENTATION_PART, PRESENTATION_CONTENT_TYPE, presentationXml(slideSize)),
    rawXml(SLIDE_PART, SLIDE_CONTENT_TYPE, slideXml()),
    rawXml(SLIDE_LAYOUT_PART, SLIDE_LAYOUT_CONTENT_TYPE, slideLayoutXml()),
    rawXml(SLIDE_MASTER_PART, SLIDE_MASTER_CONTENT_TYPE, slideMasterXml()),
    rawXml(THEME_PART, THEME_CONTENT_TYPE, themeXml()),
    rawXml(APP_PROPS_PART, APP_PROPS_CONTENT_TYPE, appPropertiesXml()),
    rawXml(CORE_PROPS_PART, CORE_PROPS_CONTENT_TYPE, corePropertiesXml()),
  ];
}

function rawXml(partPath: PartPath, contentType: string, xml: string): RawPackagePart {
  return { kind: "binary", partPath, contentType, bytes: textEncoder.encode(xml) };
}

function xmlPart(content: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${content}`;
}

function presentationXml(slideSize: SlideSize): string {
  return xmlPart(
    `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
      `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
      `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
      `<p:sldSz cx="${slideSize.width}" cy="${slideSize.height}" type="screen16x9"/>` +
      `<p:notesSz cx="6858000" cy="9144000"/>` +
      `<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle>` +
      `</p:presentation>`,
  );
}

function slideXml(): string {
  return xmlPart(
    `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
      `<p:cSld><p:spTree>${emptyGroupShapeProperties()}</p:spTree></p:cSld>` +
      `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` +
      `</p:sld>`,
  );
}

function slideLayoutXml(): string {
  return xmlPart(
    `<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
      `type="blank" preserve="1">` +
      `<p:cSld name="Blank"><p:spTree>${emptyGroupShapeProperties()}</p:spTree></p:cSld>` +
      `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` +
      `</p:sldLayout>`,
  );
}

function slideMasterXml(): string {
  return xmlPart(
    `<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
      `<p:cSld><p:spTree>${emptyGroupShapeProperties()}</p:spTree></p:cSld>` +
      `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" ` +
      `accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" ` +
      `accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
      `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
      `<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>` +
      `</p:sldMaster>`,
  );
}

function emptyGroupShapeProperties(): string {
  return (
    `<p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`
  );
}

function themeXml(): string {
  return xmlPart(
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `name="Office Theme">` +
      `<a:themeElements>${colorSchemeXml()}${fontSchemeXml()}${formatSchemeXml()}</a:themeElements>` +
      `<a:objectDefaults/><a:extraClrSchemeLst/>` +
      `</a:theme>`,
  );
}

function colorSchemeXml(): string {
  return (
    `<a:clrScheme name="Office">` +
    `<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>` +
    `<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>` +
    `<a:dk2><a:srgbClr val="44546A"/></a:dk2>` +
    `<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>` +
    `<a:accent1><a:srgbClr val="4472C4"/></a:accent1>` +
    `<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>` +
    `<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>` +
    `<a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
    `<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>` +
    `<a:accent6><a:srgbClr val="70AD47"/></a:accent6>` +
    `<a:hlink><a:srgbClr val="0563C1"/></a:hlink>` +
    `<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>` +
    `</a:clrScheme>`
  );
}

function fontSchemeXml(): string {
  return (
    `<a:fontScheme name="Office">` +
    `<a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
    `<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>` +
    `</a:fontScheme>`
  );
}

function formatSchemeXml(): string {
  return (
    `<a:fmtScheme name="Office">` +
    `<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>` +
    `<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">` +
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>` +
    `<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">` +
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>` +
    `<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">` +
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>` +
    `</a:lnStyleLst>` +
    `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>` +
    `<a:effectStyle><a:effectLst/></a:effectStyle>` +
    `<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
    `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>` +
    `</a:fmtScheme>`
  );
}

function appPropertiesXml(): string {
  return xmlPart(
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
      `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
      `<Application>pptx-glimpse</Application>` +
      `<PresentationFormat>On-screen Show (16:9)</PresentationFormat>` +
      `<Slides>1</Slides>` +
      `</Properties>`,
  );
}

function corePropertiesXml(): string {
  return xmlPart(
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      `xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
      `xmlns:dcterms="http://purl.org/dc/terms/" ` +
      `xmlns:dcmitype="http://purl.org/dc/dcmitype/" ` +
      `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dc:creator>pptx-glimpse</dc:creator>` +
      `<cp:lastModifiedBy>pptx-glimpse</cp:lastModifiedBy>` +
      `</cp:coreProperties>`,
  );
}
