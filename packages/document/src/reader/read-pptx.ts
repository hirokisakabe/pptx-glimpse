/**
 * `readPptx(input)` — PptxSourceModel source reader  initial slice.
 *
 * Reads the PPTX ZIP package、package graph (content types / relationships /
 * Internal note.
 * Internal note.
 * Internal note.
 * Internal note.
 * Internal note.
 *
 * Internal note.
 * Internal note.
 * Internal note.
 *
 * Computed view generation and writer output are responsibilities of modules separate from the reader.
 */

import { unzipSync } from "fflate";

import type {
  ContentTypeDefault,
  ContentTypeOverride,
  Diagnostic,
  MediaPart,
  PackagePartRef,
  PartPath,
  PartRelationships,
  PptxSourceModel,
  RawPackagePart,
  Relationship,
  SlideSize,
  SourcePresentation,
  SourceSlide,
  SourceSlideLayout,
  SourceSlideMaster,
  SourceTheme,
} from "../source/index.js";
import { asEmu, asPartPath, asRelationshipId } from "../source/index.js";
import {
  isRelationshipPart,
  parseRelationshipTargetMode,
  relationshipsSourcePartPath,
  resolveInternalRelationshipTarget,
  resolveRelationshipTarget,
} from "../source/package-paths.js";
import { createSidecarIdFactory } from "./raw-node.js";
import { parseSlide, parseSlideLayout, parseSlideMaster, parseTheme } from "./slide-parts.js";
import { parseTextStyle } from "./text.js";
import {
  getAttr,
  getChild,
  getChildArray,
  getNamespacedAttr,
  navigateOrdered,
  parseXml,
  parseXmlOrdered,
  type XmlNode,
  type XmlOrderedNode,
} from "./xml.js";

/** Internal note. */
export type ReadPptxInput = Uint8Array;

const CONTENT_TYPES_PART = "[Content_Types].xml";
const PACKAGE_ROOT_PART = "";

const OFFICE_DOCUMENT_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_LAYOUT_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const SLIDE_MASTER_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const THEME_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const PRESENTATION_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

const textDecoder = new TextDecoder();

/**
 * Internal note.
 *
 * Internal note.
 */
export function readPptx(input: ReadPptxInput): PptxSourceModel {
  const entries = unzipPackage(input);
  const diagnostics: Diagnostic[] = [];

  const contentTypes = readContentTypes(entries);
  const relationships = readRelationships(entries);

  const parts: PackagePartRef[] = [];
  const media: MediaPart[] = [];
  const rawParts: RawPackagePart[] = [];

  for (const [path, bytes] of entries) {
    if (path === CONTENT_TYPES_PART) continue;
    const contentType = resolveContentType(path, contentTypes.defaults, contentTypes.overrides);
    parts.push({ partPath: asPartPath(path), contentType });

    if (isMediaContentType(contentType)) {
      media.push({ partPath: asPartPath(path), contentType, bytes });
      continue;
    }

    // Internal note.
    // Internal note.
    if (isRelationshipPart(path)) continue;

    // Internal note.
    // Byte equality is not a goal, but writing untouched parts back as original bytes
    // is the most faithful structural round trip.
    rawParts.push({ kind: "binary", partPath: asPartPath(path), contentType, bytes });
  }

  const presentation = readPresentation(
    entries,
    relationships,
    contentTypes.overrides,
    diagnostics,
  );

  const hierarchy = readSlideHierarchy(entries, relationships, presentation, diagnostics);

  return {
    packageGraph: {
      contentTypes,
      parts,
      relationships,
      media,
      rawParts,
    },
    presentation,
    slides: hierarchy.slides,
    slideLayouts: hierarchy.slideLayouts,
    slideMasters: hierarchy.slideMasters,
    themes: hierarchy.themes,
    diagnostics,
  };
}

interface SlideHierarchy {
  readonly slides: readonly SourceSlide[];
  readonly slideLayouts: readonly SourceSlideLayout[];
  readonly slideMasters: readonly SourceSlideMaster[];
  readonly themes: readonly SourceTheme[];
}

/**
 * Internal note.
 * Internal note.
 * and collected in discovery order.
 */
function readSlideHierarchy(
  entries: Map<string, Uint8Array>,
  relationships: readonly PartRelationships[],
  presentation: SourcePresentation,
  diagnostics: Diagnostic[],
): SlideHierarchy {
  const slides: SourceSlide[] = [];
  const layoutPaths = new OrderedPathSet();

  for (const slidePath of presentation.slidePartPaths) {
    const part = parsePartRoot(entries, slidePath, "sld", diagnostics, true);
    if (part === undefined) continue;
    const layoutPath = resolveSingleRel(relationships, slidePath, SLIDE_LAYOUT_REL_TYPE);
    if (layoutPath === undefined) {
      diagnostics.push({
        severity: "warning",
        code: "slide-layout-unresolved",
        message: `slide '${slidePath}' has no resolvable slideLayout relationship`,
        handle: { partPath: slidePath },
      });
    } else {
      layoutPaths.add(layoutPath);
    }
    slides.push(
      parseSlide(
        part.root,
        slidePath,
        layoutPath ?? asPartPath(""),
        createSidecarIdFactory(slidePath),
        navigateOrdered(part.orderedRoot, ["cSld", "spTree"]),
      ),
    );
  }

  const slideLayouts: SourceSlideLayout[] = [];
  const masterPaths = new OrderedPathSet();
  for (const layoutPath of layoutPaths.values()) {
    const part = parsePartRoot(entries, layoutPath, "sldLayout", diagnostics, true);
    if (part === undefined) continue;
    const masterPath = resolveSingleRel(relationships, layoutPath, SLIDE_MASTER_REL_TYPE);
    if (masterPath !== undefined) masterPaths.add(masterPath);
    slideLayouts.push(
      parseSlideLayout(
        part.root,
        layoutPath,
        masterPath ?? asPartPath(""),
        createSidecarIdFactory(layoutPath),
        navigateOrdered(part.orderedRoot, ["cSld", "spTree"]),
      ),
    );
  }

  const slideMasters: SourceSlideMaster[] = [];
  const themePaths = new OrderedPathSet();
  for (const masterPath of masterPaths.values()) {
    const part = parsePartRoot(entries, masterPath, "sldMaster", diagnostics, true);
    if (part === undefined) continue;
    const themePath = resolveSingleRel(relationships, masterPath, THEME_REL_TYPE);
    if (themePath !== undefined) themePaths.add(themePath);
    const masterLayoutPaths = resolveAllRels(relationships, masterPath, SLIDE_LAYOUT_REL_TYPE);
    slideMasters.push(
      parseSlideMaster(
        part.root,
        masterPath,
        themePath,
        masterLayoutPaths,
        createSidecarIdFactory(masterPath),
        navigateOrdered(part.orderedRoot, ["cSld", "spTree"]),
      ),
    );
  }

  const themes: SourceTheme[] = [];
  for (const themePath of themePaths.values()) {
    const part = parsePartRoot(entries, themePath, "theme", diagnostics, true);
    if (part === undefined) continue;
    themes.push(
      parseTheme(part.root, themePath, createSidecarIdFactory(themePath), part.orderedRoot),
    );
  }

  return { slides, slideLayouts, slideMasters, themes };
}

interface ParsedPartRoot {
  readonly root: XmlNode;
  readonly orderedRoot: readonly XmlOrderedNode[];
}

/** Internal note. */
function parsePartRoot(
  entries: Map<string, Uint8Array>,
  partPath: PartPath,
  rootLocalName: string,
  diagnostics: Diagnostic[],
  includeOrderedRoot: boolean,
): ParsedPartRoot | undefined {
  const bytes = entries.get(partPath);
  if (!bytes) {
    diagnostics.push({
      severity: "warning",
      code: "part-missing",
      message: `part '${partPath}' referenced by the package graph is missing`,
      handle: { partPath },
    });
    return undefined;
  }
  const xml = textDecoder.decode(bytes);
  const root = getChild(parseXml(xml), rootLocalName);
  if (root === undefined) {
    diagnostics.push({
      severity: "warning",
      code: "part-root-unexpected",
      message: `part '${partPath}' does not have the expected <${rootLocalName}> root`,
      handle: { partPath },
    });
    return undefined;
  }
  const orderedRoot = includeOrderedRoot
    ? (navigateOrdered(parseXmlOrdered(xml), [rootLocalName]) ?? [])
    : [];
  return { root, orderedRoot };
}

/** Internal note. */
function resolveSingleRel(
  relationships: readonly PartRelationships[],
  sourcePart: PartPath,
  relType: string,
): PartPath | undefined {
  const rels = relationships.find((rel) => rel.sourcePartPath === sourcePart)?.relationships;
  const match = rels?.find((rel) => rel.type === relType && rel.targetMode !== "External");
  if (match === undefined) return undefined;
  return resolveInternalRelationshipTarget(sourcePart, match);
}

/** Internal note. */
function resolveAllRels(
  relationships: readonly PartRelationships[],
  sourcePart: PartPath,
  relType: string,
): PartPath[] {
  const rels = relationships.find((rel) => rel.sourcePartPath === sourcePart)?.relationships ?? [];
  return rels
    .filter((rel) => rel.type === relType && rel.targetMode !== "External")
    .flatMap((rel) => {
      const target = resolveInternalRelationshipTarget(sourcePart, rel);
      return target === undefined ? [] : [target];
    });
}

/** Small set that deduplicates part paths while preserving insertion order. */
class OrderedPathSet {
  private readonly seen = new Set<string>();
  private readonly order: PartPath[] = [];

  add(path: PartPath): void {
    if (this.seen.has(path)) return;
    this.seen.add(path);
    this.order.push(path);
  }

  values(): readonly PartPath[] {
    return this.order;
  }
}

/** Internal note. */
function unzipPackage(input: ReadPptxInput): Map<string, Uint8Array> {
  const unzipped = unzipSync(input);
  const entries = new Map<string, Uint8Array>();
  for (const [path, bytes] of Object.entries(unzipped)) {
    if (path.endsWith("/")) continue; // Internal note.
    entries.set(path, bytes);
  }
  return entries;
}

interface ContentTypes {
  readonly defaults: readonly ContentTypeDefault[];
  readonly overrides: readonly ContentTypeOverride[];
}

function readContentTypes(entries: Map<string, Uint8Array>): ContentTypes {
  const bytes = entries.get(CONTENT_TYPES_PART);
  if (!bytes) return { defaults: [], overrides: [] };

  const root = getChild(parseXml(textDecoder.decode(bytes)), "Types");

  const defaults: ContentTypeDefault[] = [];
  for (const node of getChildArray(root, "Default")) {
    const extension = getAttr(node, "Extension");
    const contentType = getAttr(node, "ContentType");
    if (extension === undefined || contentType === undefined) continue;
    defaults.push({ extension, contentType });
  }

  const overrides: ContentTypeOverride[] = [];
  for (const node of getChildArray(root, "Override")) {
    const partName = getAttr(node, "PartName");
    const contentType = getAttr(node, "ContentType");
    if (partName === undefined || contentType === undefined) continue;
    overrides.push({ partName: asPartPath(stripLeadingSlash(partName)), contentType });
  }

  return { defaults, overrides };
}

function readRelationships(entries: Map<string, Uint8Array>): PartRelationships[] {
  const result: PartRelationships[] = [];

  for (const [path, bytes] of entries) {
    if (!isRelationshipPart(path)) continue;

    const root = getChild(parseXml(textDecoder.decode(bytes)), "Relationships");
    const relationships: Relationship[] = [];
    for (const node of getChildArray(root, "Relationship")) {
      const id = getAttr(node, "Id");
      const type = getAttr(node, "Type");
      const target = getAttr(node, "Target");
      if (id === undefined || type === undefined || target === undefined) continue;
      const targetMode = parseRelationshipTargetMode(getAttr(node, "TargetMode"));
      relationships.push({
        id: asRelationshipId(id),
        type,
        target,
        ...(targetMode !== undefined ? { targetMode } : {}),
      });
    }

    result.push({ sourcePartPath: relationshipsSourcePartPath(path), relationships });
  }

  return result;
}

function readPresentation(
  entries: Map<string, Uint8Array>,
  relationships: readonly PartRelationships[],
  overrides: readonly ContentTypeOverride[],
  diagnostics: Diagnostic[],
): SourcePresentation {
  const presentationPath = locatePresentationPart(relationships, overrides);
  if (presentationPath === undefined) {
    throw new Error("readPptx: presentation part not found; input is not a valid PPTX package");
  }

  const bytes = entries.get(presentationPath);
  if (!bytes) {
    throw new Error(
      `readPptx: presentation part '${presentationPath}' is missing from the package`,
    );
  }

  const presentationPartPath = asPartPath(presentationPath);
  const root = getChild(parseXml(textDecoder.decode(bytes)), "presentation");
  if (root === undefined) {
    // If the resolved part is `<p:presentation>` not (relationship / content type
    // Internal note.
    throw new Error(
      `readPptx: part '${presentationPath}' is not a presentation part (missing p:presentation root)`,
    );
  }

  const slideSize = readSlideSize(root);
  const defaultTextStyle = parseTextStyle(getChild(root, "defaultTextStyle"));
  const presentationRels = relationships.find(
    (rel) => rel.sourcePartPath === presentationPath,
  )?.relationships;

  const slidePartPaths: PartPath[] = [];
  const sldIdLst = getChild(root, "sldIdLst");
  for (const sldId of getChildArray(sldIdLst, "sldId")) {
    const relId = getNamespacedAttr(sldId, "id");
    if (relId === undefined) continue;
    const handle = { partPath: presentationPartPath, relationshipId: asRelationshipId(relId) };
    const relationship = presentationRels?.find((rel) => rel.id === relId);
    if (relationship === undefined) {
      diagnostics.push({
        severity: "warning",
        code: "slide-relationship-unresolved",
        message: `slide relationship '${relId}' referenced by presentation could not be resolved`,
        handle,
      });
      continue;
    }
    // A relationship referenced by sldId should be an internal slide reference. If type / targetMode
    // does not match, exclude it to avoid breaking the slidePartPaths contract.
    if (relationship.type !== SLIDE_REL_TYPE || relationship.targetMode === "External") {
      diagnostics.push({
        severity: "warning",
        code: "slide-relationship-invalid",
        message: `relationship '${relId}' referenced by p:sldId is not an internal slide relationship`,
        handle,
      });
      continue;
    }
    slidePartPaths.push(
      asPartPath(resolveRelationshipTarget(presentationPath, relationship.target)),
    );
  }

  return {
    partPath: presentationPartPath,
    ...(slideSize !== undefined ? { slideSize } : {}),
    ...(defaultTextStyle !== undefined ? { defaultTextStyle } : {}),
    slidePartPaths,
    handle: { partPath: presentationPartPath },
  };
}

function readSlideSize(presentationRoot: XmlNode | undefined): SlideSize | undefined {
  const sldSz = getChild(presentationRoot, "sldSz");
  const cx = getAttr(sldSz, "cx");
  const cy = getAttr(sldSz, "cy");
  if (cx === undefined || cy === undefined) return undefined;
  const width = Number(cx);
  const height = Number(cy);
  if (!Number.isFinite(width) || !Number.isFinite(height)) return undefined;
  return { width: asEmu(width), height: asEmu(height) };
}

/**
 * Internal note.
 * and falls back to searching content type overrides.
 */
function locatePresentationPart(
  relationships: readonly PartRelationships[],
  overrides: readonly ContentTypeOverride[],
): string | undefined {
  const rootRels = relationships.find(
    (rel) => rel.sourcePartPath === PACKAGE_ROOT_PART,
  )?.relationships;
  const officeDocumentRel = rootRels?.find(
    (rel) => rel.type === OFFICE_DOCUMENT_REL_TYPE && rel.targetMode !== "External",
  );
  if (officeDocumentRel !== undefined) {
    return resolveRelationshipTarget(PACKAGE_ROOT_PART, officeDocumentRel.target);
  }

  const override = overrides.find((entry) => entry.contentType === PRESENTATION_CONTENT_TYPE);
  return override?.partName;
}

const MEDIA_CONTENT_TYPE_PREFIXES = ["image/", "audio/", "video/"];

function isMediaContentType(contentType: string): boolean {
  return MEDIA_CONTENT_TYPE_PREFIXES.some((prefix) => contentType.startsWith(prefix));
}

function resolveContentType(
  path: string,
  defaults: readonly ContentTypeDefault[],
  overrides: readonly ContentTypeOverride[],
): string {
  const override = overrides.find((entry) => entry.partName === path);
  if (override !== undefined) return override.contentType;

  const extension = extensionOf(path);
  if (extension !== undefined) {
    const fallback = defaults.find(
      (entry) => entry.extension.toLowerCase() === extension.toLowerCase(),
    );
    if (fallback !== undefined) return fallback.contentType;
  }

  return "application/octet-stream";
}

function extensionOf(path: string): string | undefined {
  const dot = path.lastIndexOf(".");
  const slash = path.lastIndexOf("/");
  if (dot === -1 || dot < slash) return undefined;
  return path.slice(dot + 1);
}

function stripLeadingSlash(path: string): string {
  return path.startsWith("/") ? path.slice(1) : path;
}
