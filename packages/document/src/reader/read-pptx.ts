/**
 * `readPptx(input)` - initial slice of the PptxSourceModel source reader.
 *
 * Reads the PPTX ZIP package, package graph (content types, relationships, and media),
 * and presentation metadata (slide size, slide order, and master order), then returns them as a
 * PptxSourceModel source. It follows both the presentation master / master layout id lists and
 * each slide's slide -> layout -> master -> theme chain, merging reachable parts as typed source
 * nodes. Unedited or unsupported parts and child elements are retained as
 * raw package material or raw sidecars and used as the basis for structural round-trip
 * preservation.
 *
 * Slide, layout, master, and theme parts that contain typed nodes also keep their
 * original bytes through `packageGraph.rawParts` for writeback when the part is
 * untouched. The typed model is for editing and computed views; the raw part is for
 * round-tripping.
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
  hasChild,
  navigateOrdered,
  parseXml,
  parseXmlOrdered,
  parseXmlOrderedQualified,
  type XmlNode,
  type XmlOrderedNode,
} from "./xml.js";

/** Input bytes for `readPptx`. */
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
 * Reads a PPTX byte string and returns a PptxSourceModel source.
 *
 * @throws If presentation part is not found (= not a valid PPTX).
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

    // content types / relationships as structural data
    // It can be reconstructed from `contentTypes` / `relationships`, so it is not included in raw.
    if (isRelationshipPart(path)) continue;

    // In this slice, all parts that are not interpreted as typed are retained in the original byte sequence.
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
 * Reads the presentation's ordered masters and their ordered layouts, plus the
 * layout -> master -> theme chain from each slide in presentation order. Parts are
 * deduplicated while retaining the applicable authoring/discovery order.
 */
function readSlideHierarchy(
  entries: Map<string, Uint8Array>,
  relationships: readonly PartRelationships[],
  presentation: SourcePresentation,
  diagnostics: Diagnostic[],
): SlideHierarchy {
  const slides: SourceSlide[] = [];
  const layoutPaths = new OrderedPathSet();
  const masterPaths = new OrderedPathSet();

  for (const masterPath of presentation.slideMasterPartPaths) {
    masterPaths.add(masterPath);
  }

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
    const masterLayoutPaths = resolveOrderedRelationshipPaths({
      root: part.root,
      sourcePartPath: masterPath,
      relationships,
      listName: "sldLayoutIdLst",
      itemName: "sldLayoutId",
      relationshipType: SLIDE_LAYOUT_REL_TYPE,
      relationshipLabel: "slide layout",
      diagnosticCodePrefix: "slide-layout",
      diagnostics,
    });
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

  const readLayoutPaths = new Set(slideLayouts.map((layout) => layout.partPath));
  for (const layoutPath of slideMasters.flatMap((master) => master.layoutPartPaths)) {
    if (readLayoutPaths.has(layoutPath)) continue;
    const part = parsePartRoot(entries, layoutPath, "sldLayout", diagnostics, true);
    if (part === undefined) continue;
    const masterPath = resolveSingleRel(relationships, layoutPath, SLIDE_MASTER_REL_TYPE);
    slideLayouts.push(
      parseSlideLayout(
        part.root,
        layoutPath,
        masterPath ?? asPartPath(""),
        createSidecarIdFactory(layoutPath),
        navigateOrdered(part.orderedRoot, ["cSld", "spTree"]),
      ),
    );
    readLayoutPaths.add(layoutPath);
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

/** Parses the byte string of part and returns the root element of the specified local name. */
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
    ? (navigateOrdered(
        rootLocalName === "theme" ? parseXmlOrdered(xml) : parseXmlOrderedQualified(xml),
        [rootLocalName],
      ) ?? [])
    : [];
  return { root, orderedRoot };
}

/** Resolves the first internal (non-External) match in the relationship for the specified source part. */
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

/** Resolve all applicable internal types among the relationships of the specified source part. */
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

/** Unzips the ZIP and returns a Map of part path -> bytes excluding directory entries. */
function unzipPackage(input: ReadPptxInput): Map<string, Uint8Array> {
  const unzipped = unzipSync(input);
  const entries = new Map<string, Uint8Array>();
  for (const [path, bytes] of Object.entries(unzipped)) {
    if (path.endsWith("/")) continue; // Directory entries are ignored.
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
    // (e.g. pointing to a different part), it does not make a broken package appear ``readable''.
    throw new Error(
      `readPptx: part '${presentationPath}' is not a presentation part (missing p:presentation root)`,
    );
  }

  const slideSize = readSlideSize(root);
  const defaultTextStyle = parseTextStyle(getChild(root, "defaultTextStyle"));
  const presentationRels = relationships.find(
    (rel) => rel.sourcePartPath === presentationPath,
  )?.relationships;

  const slideMasterPartPaths = resolveOrderedRelationshipPaths({
    root,
    sourcePartPath: presentationPartPath,
    relationships,
    listName: "sldMasterIdLst",
    itemName: "sldMasterId",
    relationshipType: SLIDE_MASTER_REL_TYPE,
    relationshipLabel: "slide master",
    diagnosticCodePrefix: "slide-master",
    diagnostics,
  });

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
    slideMasterPartPaths,
    slidePartPaths,
    handle: { partPath: presentationPartPath },
  };
}

interface OrderedRelationshipPathsInput {
  readonly root: XmlNode;
  readonly sourcePartPath: PartPath;
  readonly relationships: readonly PartRelationships[];
  readonly listName: string;
  readonly itemName: string;
  readonly relationshipType: string;
  readonly relationshipLabel: string;
  readonly diagnosticCodePrefix: string;
  readonly diagnostics: Diagnostic[];
}

/**
 * Resolves an OOXML relationship-id list in authoring order. Relationship order is used only
 * when the id-list element itself is absent, matching Office's legacy-package fallback.
 */
function resolveOrderedRelationshipPaths(input: OrderedRelationshipPathsInput): PartPath[] {
  if (!hasChild(input.root, input.listName)) {
    return resolveAllRels(input.relationships, input.sourcePartPath, input.relationshipType);
  }

  const sourceRelationships = input.relationships.find(
    (entry) => entry.sourcePartPath === input.sourcePartPath,
  )?.relationships;
  const paths: PartPath[] = [];
  const list = getChild(input.root, input.listName);

  for (const item of getChildArray(list, input.itemName)) {
    const relId = getNamespacedAttr(item, "id");
    const handle = {
      partPath: input.sourcePartPath,
      ...(relId !== undefined ? { relationshipId: asRelationshipId(relId) } : {}),
    };
    if (relId === undefined) {
      input.diagnostics.push({
        severity: "warning",
        code: `${input.diagnosticCodePrefix}-relationship-unresolved`,
        message: `${input.itemName} in ${input.sourcePartPath} has no relationship id`,
        handle,
      });
      continue;
    }

    const relationship = sourceRelationships?.find((candidate) => candidate.id === relId);
    if (relationship === undefined) {
      input.diagnostics.push({
        severity: "warning",
        code: `${input.diagnosticCodePrefix}-relationship-unresolved`,
        message: `${input.relationshipLabel} relationship '${relId}' referenced by '${input.sourcePartPath}' could not be resolved`,
        handle,
      });
      continue;
    }
    if (relationship.type !== input.relationshipType || relationship.targetMode === "External") {
      input.diagnostics.push({
        severity: "warning",
        code: `${input.diagnosticCodePrefix}-relationship-invalid`,
        message: `relationship '${relId}' referenced by ${input.itemName} is not an internal ${input.relationshipLabel} relationship`,
        handle,
      });
      continue;
    }

    const target = resolveInternalRelationshipTarget(input.sourcePartPath, relationship);
    if (target === undefined) {
      input.diagnostics.push({
        severity: "warning",
        code: `${input.diagnosticCodePrefix}-relationship-unresolved`,
        message: `${input.relationshipLabel} relationship '${relId}' referenced by '${input.sourcePartPath}' could not be resolved`,
        handle,
      });
      continue;
    }
    paths.push(target);
  }

  return paths;
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
 * Resolve presentation part path. root relationship (officeDocument type)
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
