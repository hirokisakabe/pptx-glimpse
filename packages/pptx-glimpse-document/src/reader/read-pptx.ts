/**
 * `readPptx(input)` — CleanDoc source reader の最初の slice。
 *
 * PPTX ZIP package を読み、package graph (content types / relationships /
 * media) と presentation metadata (slide size / slide order) を CleanDoc source
 * として返す。未編集・未対応の part は raw package material として保持し、
 * structural round-trip の土台にする (`docs/raw-ooxml-round-trip.md`)。
 *
 * 本 slice の scope 外 (後続 issue):
 * - shape / text / image source node の typed parsing → slide/layout/master/
 *   theme part は raw として保持しつつ、typed 配列は空のままにする。
 * - computed view 生成 / writer 出力。
 *
 * そのため `slides` / `slideLayouts` / `slideMasters` / `themes` は空配列を返し、
 * これらの part は `packageGraph.rawParts` 経由で round-trip 可能な形で保持する。
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
} from "../source/index.js";
import { asEmu, asPartPath, asRelationshipId } from "../source/index.js";
import {
  getAttr,
  getChild,
  getChildArray,
  getNamespacedAttr,
  parseXml,
  type XmlNode,
} from "./xml.js";

/** `readPptx` の入力。`Buffer` は `Uint8Array` のサブクラスとして受け付ける。 */
export type ReadPptxInput = Uint8Array;

const CONTENT_TYPES_PART = "[Content_Types].xml";
const RELS_SUFFIX = ".rels";
const PACKAGE_ROOT_PART = "";

const OFFICE_DOCUMENT_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
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

  return {
    packageGraph: {
      contentTypes,
      parts,
      relationships,
      media,
      rawParts,
    },
    presentation,
    // typed slide/layout/master/theme reading は後続 slice。冒頭コメント参照。
    slides: [],
    slideLayouts: [],
    slideMasters: [],
    themes: [],
    diagnostics,
  };
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
