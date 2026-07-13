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
  Emu,
  MediaPart,
  PackagePartRef,
  PartPath,
  PartRelationships,
  PptxSourceModel,
  RawPackagePart,
  Relationship,
  SlideSize,
  SourceBackground,
  SourceColorMap,
  SourceTheme,
  SourceThemeFormatScheme,
} from "../source/index.js";
import { asEmu, asPartPath, asRelationshipId } from "../source/index.js";

export interface CreatePptxOptions {
  readonly slideSize?: SlideSize;
  readonly slideMaster?: CreatePptxSlideMasterOptions;
  readonly slideLayout?: CreatePptxSlideLayoutOptions;
}

export type CreatePptxBackground =
  | { readonly kind: "solid"; readonly color: { readonly kind: "srgb"; readonly hex: string } }
  | { readonly kind: "image"; readonly bytes: Uint8Array };

export interface CreatePptxSlideMasterOptions {
  readonly name?: string;
  readonly background?: CreatePptxBackground;
}

export interface SlideLayoutMargin {
  readonly left: Emu;
  readonly right: Emu;
  readonly top: Emu;
  readonly bottom: Emu;
}

export interface CreatePptxSlideLayoutOptions {
  readonly name?: string;
  readonly margin?: SlideLayoutMargin;
}

interface NormalizedCreatePptxOptions {
  readonly slideSize: SlideSize;
  readonly masterName: string;
  readonly layoutName: string;
  readonly background?: CreatePptxBackground;
  readonly margin?: SlideLayoutMargin;
  readonly backgroundImage?: DetectedImageType;
}

interface DetectedImageType {
  readonly contentType: "image/png" | "image/jpeg";
  readonly extension: "png" | "jpeg";
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
const BACKGROUND_IMAGE_PART_PNG = asPartPath("ppt/media/image1.png");
const BACKGROUND_IMAGE_PART_JPEG = asPartPath("ppt/media/image1.jpeg");

const DEFAULT_SLIDE_WIDTH = 9144000;
const DEFAULT_SLIDE_HEIGHT = 5143500;

export function createPptx(options: CreatePptxOptions = {}): PptxSourceModel {
  const normalized = normalizeOptions(options);
  const { slideSize } = normalized;
  const overrides = createContentTypeOverrides();
  const rawParts = createRawParts(normalized);
  const media = createMedia(normalized);
  const parts = createPackageParts(overrides, media);
  const masterBackground = createSourceBackground(normalized);

  return {
    packageGraph: {
      contentTypes: {
        defaults: [
          { extension: "rels", contentType: RELS_CONTENT_TYPE },
          { extension: "xml", contentType: XML_CONTENT_TYPE },
          ...(normalized.backgroundImage !== undefined
            ? [
                {
                  extension: normalized.backgroundImage.extension,
                  contentType: normalized.backgroundImage.contentType,
                },
              ]
            : []),
        ],
        overrides,
      },
      parts,
      relationships: createRelationships(normalized),
      media,
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
        name: normalized.layoutName,
        type: "blank",
        show: true,
        shapes: [],
        ...(normalized.margin !== undefined
          ? {
              defaultTextBodyProperties: {
                marginLeft: normalized.margin.left,
                marginRight: normalized.margin.right,
                marginTop: normalized.margin.top,
                marginBottom: normalized.margin.bottom,
              },
            }
          : {}),
        handle: { partPath: SLIDE_LAYOUT_PART },
      },
    ],
    slideMasters: [
      {
        partPath: SLIDE_MASTER_PART,
        name: normalized.masterName,
        themePartPath: THEME_PART,
        layoutPartPaths: [SLIDE_LAYOUT_PART],
        colorMap: createDefaultColorMap(),
        ...(masterBackground !== undefined ? { background: masterBackground } : {}),
        shapes: [],
        handle: { partPath: SLIDE_MASTER_PART },
      },
    ],
    themes: [createThemeSource()],
    diagnostics: [],
  };
}

function normalizeOptions(options: CreatePptxOptions): NormalizedCreatePptxOptions {
  const slideSize = normalizeSlideSize(options.slideSize);
  const masterName = normalizeName(options.slideMaster?.name, "slideMaster.name", "Blank Master");
  const layoutName = normalizeName(options.slideLayout?.name, "slideLayout.name", "Blank");
  const background = options.slideMaster?.background;
  let backgroundImage: DetectedImageType | undefined;
  if (background !== undefined) {
    const backgroundKind: unknown = Reflect.get(background, "kind");
    if (backgroundKind !== "solid" && backgroundKind !== "image") {
      throw new Error("createPptx: slideMaster.background.kind must be solid or image");
    }
    if (background.kind === "solid") {
      assertHexColor(background.color.hex, "slideMaster.background");
    } else {
      if (!(background.bytes instanceof Uint8Array)) {
        throw new Error("createPptx: slideMaster.background.bytes must be a Uint8Array");
      }
      backgroundImage = detectSupportedImageType(background.bytes);
      if (backgroundImage === undefined) {
        throw new Error("createPptx: slideMaster.background uses an unsupported image format");
      }
    }
  }
  const margin = options.slideLayout?.margin;
  if (margin !== undefined) {
    assertFiniteEmu(margin.left, "slideLayout.margin.left");
    assertFiniteEmu(margin.right, "slideLayout.margin.right");
    assertFiniteEmu(margin.top, "slideLayout.margin.top");
    assertFiniteEmu(margin.bottom, "slideLayout.margin.bottom");
  }
  return {
    slideSize,
    masterName,
    layoutName,
    ...(background !== undefined ? { background } : {}),
    ...(margin !== undefined ? { margin } : {}),
    ...(backgroundImage !== undefined ? { backgroundImage } : {}),
  };
}

function normalizeName(value: string | undefined, field: string, fallback: string): string {
  if (value === undefined) return fallback;
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`createPptx: ${field} must be a non-empty string`);
  }
  assertValidXmlName(value, field);
  return value.trim();
}

function assertValidXmlName(value: string, field: string): void {
  for (let index = 0; index < value.length; index += 1) {
    const codePoint = value.codePointAt(index);
    if (codePoint === undefined) continue;
    const valid =
      codePoint >= 0x20 &&
      codePoint <= 0x10ffff &&
      !(codePoint >= 0xd800 && codePoint <= 0xdfff) &&
      codePoint !== 0xfffe &&
      codePoint !== 0xffff;
    if (!valid) {
      throw new Error(`createPptx: ${field} contains a character forbidden in an XML attribute`);
    }
    if (codePoint > 0xffff) index += 1;
  }
}

function assertHexColor(value: string, field: string): void {
  if (!/^[0-9A-Fa-f]{6}$/.test(value)) {
    throw new Error(`createPptx: ${field} must be a 6-digit RGB color`);
  }
}

function assertFiniteEmu(value: unknown, field: string): void {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    throw new Error(`createPptx: ${field} must be a finite non-negative EMU value`);
  }
}

function normalizeSlideSize(slideSize: SlideSize | undefined): SlideSize {
  const width = slideSize?.width ?? DEFAULT_SLIDE_WIDTH;
  const height = slideSize?.height ?? DEFAULT_SLIDE_HEIGHT;
  assertPositiveFiniteNumber(width, "slideSize.width");
  assertPositiveFiniteNumber(height, "slideSize.height");
  return { width: asEmu(width), height: asEmu(height) };
}

function assertPositiveFiniteNumber(value: unknown, fieldName: string): asserts value is number {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw new Error(`createPptx: ${fieldName} must be a finite positive EMU value`);
  }
}

function createDefaultColorMap(): SourceColorMap {
  return {
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
}

function createThemeSource(): SourceTheme {
  return {
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
    formatScheme: createThemeFormatScheme(),
    handle: { partPath: THEME_PART },
  };
}

function createThemeFormatScheme(): SourceThemeFormatScheme {
  const placeholderFill = {
    kind: "solid" as const,
    color: { kind: "scheme" as const, scheme: "phClr" },
  };
  return {
    fillStyles: [placeholderFill, placeholderFill, placeholderFill],
    lineStyles: [
      { width: asEmu(6350), fill: placeholderFill, dashStyle: "solid" },
      { width: asEmu(12700), fill: placeholderFill, dashStyle: "solid" },
      { width: asEmu(19050), fill: placeholderFill, dashStyle: "solid" },
    ],
    effectStyles: [undefined, undefined, undefined],
    backgroundFillStyles: [placeholderFill, placeholderFill, placeholderFill],
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

function createPackageParts(
  overrides: readonly ContentTypeOverride[],
  media: readonly MediaPart[],
): PackagePartRef[] {
  const partRefs = overrides.map((override) => ({
    partPath: override.partName,
    contentType: override.contentType,
  }));
  return [
    ...partRefs,
    ...media.map((entry) => ({ partPath: entry.partPath, contentType: entry.contentType })),
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

function createRelationships(options: NormalizedCreatePptxOptions): PartRelationships[] {
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
        ...(options.backgroundImage !== undefined
          ? [
              rel(
                "rId3",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                `../media/image1.${options.backgroundImage.extension}`,
              ),
            ]
          : []),
      ],
    },
  ];
}

function rel(id: string, type: string, target: string): Relationship {
  return { id: asRelationshipId(id), type, target };
}

function createRawParts(options: NormalizedCreatePptxOptions): RawPackagePart[] {
  return [
    rawXml(PRESENTATION_PART, PRESENTATION_CONTENT_TYPE, presentationXml(options.slideSize)),
    rawXml(SLIDE_PART, SLIDE_CONTENT_TYPE, slideXml()),
    rawXml(SLIDE_LAYOUT_PART, SLIDE_LAYOUT_CONTENT_TYPE, slideLayoutXml(options.layoutName)),
    rawXml(SLIDE_MASTER_PART, SLIDE_MASTER_CONTENT_TYPE, slideMasterXml(options)),
    rawXml(THEME_PART, THEME_CONTENT_TYPE, themeXml()),
    rawXml(APP_PROPS_PART, APP_PROPS_CONTENT_TYPE, appPropertiesXml(options.slideSize)),
    rawXml(CORE_PROPS_PART, CORE_PROPS_CONTENT_TYPE, corePropertiesXml()),
  ];
}

function createMedia(options: NormalizedCreatePptxOptions): MediaPart[] {
  if (options.background?.kind !== "image" || options.backgroundImage === undefined) return [];
  return [
    {
      partPath:
        options.backgroundImage.extension === "png"
          ? BACKGROUND_IMAGE_PART_PNG
          : BACKGROUND_IMAGE_PART_JPEG,
      contentType: options.backgroundImage.contentType,
      bytes: new Uint8Array(options.background.bytes),
    },
  ];
}

function createSourceBackground(
  options: NormalizedCreatePptxOptions,
): SourceBackground | undefined {
  switch (options.background?.kind) {
    case undefined:
      return undefined;
    case "solid":
      return {
        kind: "fill",
        fill: {
          kind: "solid",
          color: { kind: "srgb", hex: options.background.color.hex.toUpperCase() },
        },
      };
    case "image":
      return {
        kind: "fill",
        fill: { kind: "image", blipRelationshipId: asRelationshipId("rId3") },
      };
  }
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
      `<p:sldSz cx="${slideSize.width}" cy="${slideSize.height}"${slideSizeTypeAttribute(
        slideSize,
      )}/>` +
      `<p:notesSz cx="6858000" cy="9144000"/>` +
      `<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle>` +
      `</p:presentation>`,
  );
}

function slideSizeTypeAttribute(slideSize: SlideSize): string {
  return isDefaultSlideSize(slideSize) ? ` type="screen16x9"` : "";
}

function isDefaultSlideSize(slideSize: SlideSize): boolean {
  return slideSize.width === DEFAULT_SLIDE_WIDTH && slideSize.height === DEFAULT_SLIDE_HEIGHT;
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

function slideLayoutXml(name: string): string {
  return xmlPart(
    `<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
      `type="blank" preserve="1">` +
      `<p:cSld name="${escapeXmlAttribute(name)}"><p:spTree>${emptyGroupShapeProperties()}</p:spTree></p:cSld>` +
      `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` +
      `</p:sldLayout>`,
  );
}

function slideMasterXml(options: NormalizedCreatePptxOptions): string {
  return xmlPart(
    `<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
      `<p:cSld name="${escapeXmlAttribute(options.masterName)}">${backgroundXml(options.background)}` +
      `<p:spTree>${emptyGroupShapeProperties()}</p:spTree></p:cSld>` +
      `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" ` +
      `accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" ` +
      `accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
      `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
      `<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>` +
      `</p:sldMaster>`,
  );
}

function backgroundXml(background: CreatePptxBackground | undefined): string {
  switch (background?.kind) {
    case undefined:
      return "";
    case "solid":
      return (
        `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${background.color.hex.toUpperCase()}"/>` +
        `</a:solidFill><a:effectLst/></p:bgPr></p:bg>`
      );
    case "image":
      return (
        `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId3"/>` +
        `<a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`
      );
  }
}

function escapeXmlAttribute(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
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

function appPropertiesXml(slideSize: SlideSize): string {
  return xmlPart(
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
      `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
      `<Application>pptx-glimpse</Application>` +
      (isDefaultSlideSize(slideSize)
        ? `<PresentationFormat>On-screen Show (16:9)</PresentationFormat>`
        : "") +
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

function detectSupportedImageType(bytes: Uint8Array): DetectedImageType | undefined {
  if (startsWithBytes(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
    return { contentType: "image/png", extension: "png" };
  }
  if (startsWithBytes(bytes, [0xff, 0xd8, 0xff])) {
    return { contentType: "image/jpeg", extension: "jpeg" };
  }
  return undefined;
}

function startsWithBytes(bytes: Uint8Array, prefix: readonly number[]): boolean {
  if (bytes.length < prefix.length) return false;
  return prefix.every((value, index) => bytes[index] === value);
}
