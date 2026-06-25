/**
 * `readPptx(input)` — CleanDoc source reader の最初の slice。
 *
 * PPTX ZIP package を読み、package graph (content types / relationships /
 * media) と presentation metadata (slide size / slide order) を CleanDoc source
 * として返す。さらに presentation order の各 slide から layout → master → theme
 * の chain を辿り、simple shape / text / image を typed source node として読む。
 * 未編集・未対応の part / 子要素は raw package material / raw sidecar として
 * 保持し、structural round-trip の土台にする (`docs/raw-ooxml-round-trip.md`)。
 *
 * typed node を持つ slide/layout/master/theme part も、未編集 part の忠実な
 * 書き戻しのために `packageGraph.rawParts` 経由で元バイト列を併せて保持する
 * (typed model は編集・computed view 用、raw part は round-trip 用の二重表現)。
 *
 * 本 slice の scope 外 (後続 issue): computed view 生成 / writer 出力。
 */

import { unzipSync } from "fflate";

import type {
  CleanDocSource,
  ContentTypeDefault,
  ContentTypeOverride,
  Diagnostic,
  MediaPart,
  PackagePartRef,
  PartPath,
  PartRelationships,
  RawPackagePart,
  Relationship,
  RelationshipTargetMode,
  SlideSize,
  SourcePresentation,
  SourceSlide,
  SourceSlideLayout,
  SourceSlideMaster,
  SourceTheme,
} from "../source/index.js";
import { asEmu, asPartPath, asRelationshipId } from "../source/index.js";
import { createSidecarIdFactory } from "./raw-node.js";
import { parseSlide, parseSlideLayout, parseSlideMaster, parseTheme } from "./slide-parts.js";
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

/** `readPptx` の入力。`Buffer` は `Uint8Array` のサブクラスとして受け付ける。 */
export type ReadPptxInput = Uint8Array;

const CONTENT_TYPES_PART = "[Content_Types].xml";
const RELS_SUFFIX = ".rels";
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
 * PPTX バイト列を読み、CleanDoc source を返す。
 *
 * @throws presentation part が見つからない (= 有効な PPTX ではない) 場合。
 */
export function readPptx(input: ReadPptxInput): CleanDocSource {
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

    // content types / relationships は structural data として
    // `contentTypes` / `relationships` から再構成できるため raw には含めない。
    if (isRelationshipPart(path)) continue;

    // 本 slice では typed に解釈しない part をすべて元バイト列で保持する。
    // byte 一致は目標にしないが、untouched part は原バイトのまま書き戻すのが
    // 最も忠実な structural round-trip になる (`docs/raw-ooxml-round-trip.md`)。
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
 * presentation order の各 slide から layout → master → theme を辿り、到達可能な
 * part を typed source node として読む。layout / master / theme は重複排除しつつ
 * 発見順に収集する。
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

/** part のバイト列をパースして指定 local name の root 要素を返す。 */
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

/** 指定 source part の relationship のうち、内部 (非 External) の最初の一致を解決する。 */
function resolveSingleRel(
  relationships: readonly PartRelationships[],
  sourcePart: PartPath,
  relType: string,
): PartPath | undefined {
  const rels = relationships.find((rel) => rel.sourcePartPath === sourcePart)?.relationships;
  const match = rels?.find((rel) => rel.type === relType && rel.targetMode !== "External");
  if (match === undefined) return undefined;
  return asPartPath(resolveTarget(sourcePart, match.target));
}

/** 指定 source part の relationship のうち、内部の該当 type をすべて解決する。 */
function resolveAllRels(
  relationships: readonly PartRelationships[],
  sourcePart: PartPath,
  relType: string,
): PartPath[] {
  const rels = relationships.find((rel) => rel.sourcePartPath === sourcePart)?.relationships ?? [];
  return rels
    .filter((rel) => rel.type === relType && rel.targetMode !== "External")
    .map((rel) => asPartPath(resolveTarget(sourcePart, rel.target)));
}

/** 挿入順を保ちつつ part path を重複排除する小さな集合。 */
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

/** ZIP を展開し、ディレクトリエントリを除いた part path → bytes の Map を返す。 */
function unzipPackage(input: ReadPptxInput): Map<string, Uint8Array> {
  const unzipped = unzipSync(input);
  const entries = new Map<string, Uint8Array>();
  for (const [path, bytes] of Object.entries(unzipped)) {
    if (path.endsWith("/")) continue; // ディレクトリエントリは無視。
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
      const targetMode = getAttr(node, "TargetMode");
      relationships.push({
        id: asRelationshipId(id),
        type,
        target,
        ...(targetMode !== undefined ? { targetMode: targetMode as RelationshipTargetMode } : {}),
      });
    }

    result.push({ sourcePartPath: asPartPath(relsSourcePartPath(path)), relationships });
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
    // 解決した part が `<p:presentation>` でない (relationship / content type が
    // 別 part を指している等) 場合に、壊れた package を「読めた」ように見せない。
    throw new Error(
      `readPptx: part '${presentationPath}' is not a presentation part (missing p:presentation root)`,
    );
  }

  const slideSize = readSlideSize(root);
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
    // sldId が指す relationship は slide 内部参照のはず。type / targetMode が
    // 一致しないものを slidePartPaths に混ぜると契約が壊れるため除外する。
    if (relationship.type !== SLIDE_REL_TYPE || relationship.targetMode === "External") {
      diagnostics.push({
        severity: "warning",
        code: "slide-relationship-invalid",
        message: `relationship '${relId}' referenced by p:sldId is not an internal slide relationship`,
        handle,
      });
      continue;
    }
    slidePartPaths.push(asPartPath(resolveTarget(presentationPath, relationship.target)));
  }

  return {
    partPath: presentationPartPath,
    ...(slideSize !== undefined ? { slideSize } : {}),
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
 * presentation part path を解決する。root relationship (officeDocument 型) を
 * 優先し、無ければ content type override から探す。
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
    return resolveTarget(PACKAGE_ROOT_PART, officeDocumentRel.target);
  }

  const override = overrides.find((entry) => entry.contentType === PRESENTATION_CONTENT_TYPE);
  return override?.partName;
}

const MEDIA_CONTENT_TYPE_PREFIXES = ["image/", "audio/", "video/"];

function isMediaContentType(contentType: string): boolean {
  return MEDIA_CONTENT_TYPE_PREFIXES.some((prefix) => contentType.startsWith(prefix));
}

function isRelationshipPart(path: string): boolean {
  return path.endsWith(RELS_SUFFIX) && path.includes("_rels/");
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

/**
 * `_rels/*.rels` part path から、その rels が属する source part path を求める。
 * 例: `ppt/_rels/presentation.xml.rels` → `ppt/presentation.xml`、
 * `_rels/.rels` → `""` (package root)。
 */
function relsSourcePartPath(relsPath: string): string {
  const marker = "_rels/";
  const idx = relsPath.lastIndexOf(marker);
  const dir = relsPath.slice(0, idx); // "" / "ppt/" / "ppt/slides/"
  const file = relsPath.slice(idx + marker.length); // ".rels" / "presentation.xml.rels"
  const base = file.endsWith(RELS_SUFFIX) ? file.slice(0, -RELS_SUFFIX.length) : file;
  return dir + base;
}

/**
 * relationship target を source part path 基準で解決し、package 内の絶対 part
 * path (先頭スラッシュ無し) に正規化する。external (絶対 URI) はそのまま返す。
 */
function resolveTarget(sourcePartPath: string, target: string): string {
  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(target)) return target; // 絶対 URI (external)。

  let combined: string;
  if (target.startsWith("/")) {
    combined = target.slice(1); // package root 起点。
  } else {
    const slash = sourcePartPath.lastIndexOf("/");
    const baseDir = slash === -1 ? "" : sourcePartPath.slice(0, slash);
    combined = baseDir === "" ? target : `${baseDir}/${target}`;
  }

  const segments: string[] = [];
  for (const segment of combined.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") {
      segments.pop();
      continue;
    }
    segments.push(segment);
  }
  return segments.join("/");
}
