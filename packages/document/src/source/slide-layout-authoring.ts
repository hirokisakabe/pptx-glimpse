import { getAttr, getChild, getChildArray, parseXml } from "../reader/xml.js";
import { sourceHandlesEqual } from "./edit-descriptors.js";
import {
  copyBytes,
  IMAGE_REL_TYPE,
  relativeTarget,
  requireRawBinaryPart,
  SLIDE_LAYOUT_REL_TYPE,
} from "./editing-shared.js";
import { detectSupportedImageType } from "./image-type.js";
import type { MediaPart, PartRelationships, PptxSourceModel, Relationship } from "./index.js";
import type {
  PartPath,
  PptxSourceModelAddSlideLayoutEdit,
  RelationshipId,
  SourceBackground,
  SourceHandle,
  SourceSlideLayout,
} from "./index.js";
import { asRelationshipId } from "./index.js";
import {
  addMediaPartRelationship,
  addPackagePart,
  addPartRelationship,
  nextNumberedPartPath,
  nextRelationshipId,
} from "./package-graph-mutations.js";
import type { Emu } from "./units.js";

export type SlideLayoutType =
  | "title"
  | "tx"
  | "twoColTx"
  | "tbl"
  | "txAndChart"
  | "chartAndTx"
  | "dgm"
  | "chart"
  | "txAndClipArt"
  | "clipArtAndTx"
  | "titleOnly"
  | "blank"
  | "txAndObj"
  | "objAndTx"
  | "objOnly"
  | "obj"
  | "txAndMedia"
  | "mediaAndTx"
  | "objOverTx"
  | "txOverObj"
  | "txAndTwoObj"
  | "twoObjAndTx"
  | "twoObjOverTx"
  | "fourObj"
  | "vertTx"
  | "clipArtAndVertTx"
  | "vertTitleAndTx"
  | "vertTitleAndTxOverChart"
  | "twoObj"
  | "objAndTwoObj"
  | "twoObjAndObj"
  | "cust"
  | "secHead"
  | "twoTxTwoObj"
  | "objTx"
  | "picTx";

export type AddSlideLayoutBackgroundInput =
  | { readonly kind: "solid"; readonly color: { readonly kind: "srgb"; readonly hex: string } }
  | { readonly kind: "image"; readonly bytes: Uint8Array };

export interface AddSlideLayoutMarginInput {
  readonly left: Emu;
  readonly right: Emu;
  readonly top: Emu;
  readonly bottom: Emu;
}

export interface AddSlideLayoutInput {
  /** Non-empty author-visible `p:cSld@name`. */
  readonly name: string;
  /** OOXML `ST_SlideLayoutType`; defaults to `blank`. */
  readonly type?: SlideLayoutType;
  /** Layout catalog visibility; defaults to `true`. */
  readonly show?: boolean;
  /** Optional direct layout background. */
  readonly background?: AddSlideLayoutBackgroundInput;
  /**
   * Non-negative finite EMU defaults materialized into text-bearing shapes authored on slides
   * that use this layout. These authoring-only defaults are not serialized as layout metadata.
   */
  readonly margin?: AddSlideLayoutMarginInput;
}

const SLIDE_LAYOUT_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const SLIDE_MASTER_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const SLIDE_LAYOUT_PART_PREFIX = "ppt/slideLayouts/slideLayout";
const IMAGE_MEDIA_PREFIX = "ppt/media/image";
const FIRST_LAYOUT_NUMERIC_ID = 2147483649;
const MAX_UINT32 = 4294967295;
const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

const SLIDE_LAYOUT_TYPES: ReadonlySet<string> = new Set<SlideLayoutType>([
  "title",
  "tx",
  "twoColTx",
  "tbl",
  "txAndChart",
  "chartAndTx",
  "dgm",
  "chart",
  "txAndClipArt",
  "clipArtAndTx",
  "titleOnly",
  "blank",
  "txAndObj",
  "objAndTx",
  "objOnly",
  "obj",
  "txAndMedia",
  "mediaAndTx",
  "objOverTx",
  "txOverObj",
  "txAndTwoObj",
  "twoObjAndTx",
  "twoObjOverTx",
  "fourObj",
  "vertTx",
  "clipArtAndVertTx",
  "vertTitleAndTx",
  "vertTitleAndTxOverChart",
  "twoObj",
  "objAndTwoObj",
  "twoObjAndObj",
  "cust",
  "secHead",
  "twoTxTwoObj",
  "objTx",
  "picTx",
]);

/** Adds a new empty layout to an existing slide master and appends it to authoring order. */
export function addSlideLayout(
  source: PptxSourceModel,
  masterHandle: SourceHandle,
  input: AddSlideLayoutInput,
): PptxSourceModel {
  const normalized = normalizeInput(input);
  const masterIndex = source.slideMasters.findIndex((master) =>
    sourceHandlesEqual(master.handle, masterHandle),
  );
  const master = source.slideMasters[masterIndex];
  if (master === undefined) {
    throw new Error("addSlideLayout: slide master handle was not found in PptxSourceModel source");
  }
  const masterRelationships = requireMasterRelationships(source, master.partPath);
  const newLayoutPartPath = nextNumberedPartPath(
    source.packageGraph,
    source.edits?.flatMap((edit) =>
      edit.kind === "addSlideLayout" ? [edit.newLayoutPartPath] : [],
    ) ?? [],
    SLIDE_LAYOUT_PART_PREFIX,
    ".xml",
  );
  const newRelationshipId = nextRelationshipId(masterRelationships.relationships);
  const newLayoutNumericId = nextLayoutNumericId(source, master.partPath);
  const layoutMasterRelationship: Relationship = {
    id: asRelationshipId("rId1"),
    type: SLIDE_MASTER_REL_TYPE,
    target: relativeTarget(newLayoutPartPath, master.partPath),
  };

  let image:
    | {
        readonly relationshipId: RelationshipId;
        readonly media: MediaPart;
        readonly extension: string;
      }
    | undefined;
  if (normalized.background?.kind === "image") {
    const imageType = detectSupportedImageType(normalized.background.bytes);
    if (imageType === undefined) {
      throw new Error("addSlideLayout: background uses an unsupported image format");
    }
    const mediaPartPath = nextNumberedPartPath(
      source.packageGraph,
      [],
      IMAGE_MEDIA_PREFIX,
      `.${imageType.extension}`,
    );
    image = {
      relationshipId: asRelationshipId("rId2"),
      extension: imageType.extension,
      media: {
        partPath: mediaPartPath,
        contentType: imageType.contentType,
        bytes: copyBytes(normalized.background.bytes),
      },
    };
  }

  const layoutRelationships: PartRelationships = {
    sourcePartPath: newLayoutPartPath,
    relationships: [layoutMasterRelationship],
  };
  let packageGraph = addPackagePart(source.packageGraph, {
    partPath: newLayoutPartPath,
    contentType: SLIDE_LAYOUT_CONTENT_TYPE,
    bytes: textEncoder.encode(slideLayoutXml(normalized, image?.relationshipId)),
    relationships: layoutRelationships,
  });
  if (image !== undefined) {
    packageGraph = addMediaPartRelationship(packageGraph, {
      ownerPartPath: newLayoutPartPath,
      media: image.media,
      extension: image.extension,
      relationship: {
        id: image.relationshipId,
        type: IMAGE_REL_TYPE,
        target: relativeTarget(newLayoutPartPath, image.media.partPath),
      },
      contentTypeDefaultConflictError: (existingContentType) =>
        new Error(
          `addSlideLayout: content type default for extension '${image.extension}' already maps to '${existingContentType}'`,
        ),
    });
  }
  packageGraph = addPartRelationship(packageGraph, master.partPath, {
    id: newRelationshipId,
    type: SLIDE_LAYOUT_REL_TYPE,
    target: relativeTarget(master.partPath, newLayoutPartPath),
  });

  const background = sourceBackground(normalized.background, image?.relationshipId);
  const newLayout: SourceSlideLayout = {
    partPath: newLayoutPartPath,
    masterPartPath: master.partPath,
    name: normalized.name,
    type: normalized.type,
    show: normalized.show,
    ...(background === undefined ? {} : { background }),
    ...(normalized.margin === undefined
      ? {}
      : {
          defaultTextBodyProperties: {
            marginLeft: normalized.margin.left,
            marginRight: normalized.margin.right,
            marginTop: normalized.margin.top,
            marginBottom: normalized.margin.bottom,
          },
        }),
    shapes: [],
    handle: { partPath: newLayoutPartPath },
  };
  const edit = {
    kind: "addSlideLayout",
    masterPartPath: master.partPath,
    newLayoutPartPath,
    newRelationshipId,
    newLayoutNumericId,
  } satisfies PptxSourceModelAddSlideLayoutEdit;

  return {
    ...source,
    slideLayouts: [...source.slideLayouts, newLayout],
    slideMasters: source.slideMasters.map((candidate, index) =>
      index === masterIndex
        ? { ...candidate, layoutPartPaths: [...candidate.layoutPartPaths, newLayoutPartPath] }
        : candidate,
    ),
    packageGraph,
    edits: [...(source.edits ?? []), edit],
  };
}

interface NormalizedAddSlideLayoutInput {
  readonly name: string;
  readonly type: SlideLayoutType;
  readonly show: boolean;
  readonly background?: AddSlideLayoutBackgroundInput;
  readonly margin?: AddSlideLayoutMarginInput;
}

function normalizeInput(input: AddSlideLayoutInput): NormalizedAddSlideLayoutInput {
  if (typeof input !== "object" || input === null) {
    throw new Error("addSlideLayout: input must be a slide layout object");
  }
  if (typeof input.name !== "string" || input.name.trim() === "") {
    throw new Error("addSlideLayout: name must be a non-empty string");
  }
  assertValidXmlAttribute(input.name, "name");
  const type = input.type ?? "blank";
  if (!SLIDE_LAYOUT_TYPES.has(type)) {
    throw new Error("addSlideLayout: type must be a supported OOXML slide layout type");
  }
  const show = input.show ?? true;
  if (typeof show !== "boolean") {
    throw new Error("addSlideLayout: show must be a boolean");
  }
  if (input.background !== undefined) assertBackground(input.background);
  if (input.margin !== undefined) {
    assertMargin(input.margin.left, "margin.left");
    assertMargin(input.margin.right, "margin.right");
    assertMargin(input.margin.top, "margin.top");
    assertMargin(input.margin.bottom, "margin.bottom");
  }
  return {
    name: input.name.trim(),
    type,
    show,
    ...(input.background === undefined ? {} : { background: input.background }),
    ...(input.margin === undefined ? {} : { margin: input.margin }),
  };
}

function assertBackground(input: AddSlideLayoutBackgroundInput): void {
  if (typeof input !== "object" || input === null) {
    throw new Error("addSlideLayout: background must be a background object");
  }
  if (input.kind === "solid") {
    if (
      input.color?.kind !== "srgb" ||
      typeof input.color.hex !== "string" ||
      !/^[0-9A-Fa-f]{6}$/.test(input.color.hex)
    ) {
      throw new Error("addSlideLayout: background color must be a 6-digit srgb color");
    }
    return;
  }
  if (input.kind === "image") {
    if (!(input.bytes instanceof Uint8Array)) {
      throw new Error("addSlideLayout: background bytes must be a Uint8Array");
    }
    return;
  }
  throw new Error("addSlideLayout: background kind must be solid or image");
}

function assertMargin(value: unknown, field: string): void {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    throw new Error(`addSlideLayout: ${field} must be a finite non-negative EMU value`);
  }
}

function assertValidXmlAttribute(value: string, field: string): void {
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
      throw new Error(`addSlideLayout: ${field} contains a character forbidden in XML`);
    }
    if (codePoint > 0xffff) index += 1;
  }
}

function requireMasterRelationships(
  source: PptxSourceModel,
  masterPartPath: PartPath,
): PartRelationships {
  const relationships = source.packageGraph.relationships.find(
    (candidate) => candidate.sourcePartPath === masterPartPath,
  );
  if (relationships === undefined) {
    throw new Error("addSlideLayout: slide master relationships were not found");
  }
  return relationships;
}

function nextLayoutNumericId(source: PptxSourceModel, masterPartPath: PartPath): number {
  const rawPart = requireRawBinaryPart(source, masterPartPath, "addSlideLayout");
  const root = parseXml(textDecoder.decode(rawPart.bytes));
  const master = getChild(root, "sldMaster");
  if (master === undefined) {
    throw new Error("addSlideLayout: slide master part does not contain p:sldMaster root");
  }
  const used = new Set<number>();
  for (const item of getChildArray(getChild(master, "sldLayoutIdLst"), "sldLayoutId")) {
    const id = Number(getAttr(item, "id"));
    if (Number.isInteger(id) && id >= 0 && id <= MAX_UINT32) used.add(id);
  }
  for (const edit of source.edits ?? []) {
    if (edit.kind === "addSlideLayout" && edit.masterPartPath === masterPartPath) {
      used.add(edit.newLayoutNumericId);
    }
  }
  for (let candidate = FIRST_LAYOUT_NUMERIC_ID; candidate <= MAX_UINT32; candidate += 1) {
    if (!used.has(candidate)) return candidate;
  }
  throw new Error("addSlideLayout: no unused p:sldLayoutId ID remains");
}

function sourceBackground(
  input: AddSlideLayoutBackgroundInput | undefined,
  relationshipId: RelationshipId | undefined,
): SourceBackground | undefined {
  if (input === undefined) return undefined;
  if (input.kind === "solid") {
    return {
      kind: "fill",
      fill: {
        kind: "solid",
        color: { kind: "srgb", hex: input.color.hex.toUpperCase() },
      },
    };
  }
  if (relationshipId === undefined) {
    throw new Error("addSlideLayout: image relationship ID was not allocated");
  }
  return {
    kind: "fill",
    fill: { kind: "image", blipRelationshipId: relationshipId },
  };
}

function slideLayoutXml(
  input: NormalizedAddSlideLayoutInput,
  imageRelationshipId: RelationshipId | undefined,
): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
    `type="${input.type}" show="${input.show ? "1" : "0"}" preserve="1">` +
    `<p:cSld name="${escapeXmlAttribute(input.name)}">` +
    backgroundXml(input.background, imageRelationshipId) +
    `<p:spTree>${emptyGroupShapeProperties()}</p:spTree></p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` +
    `</p:sldLayout>`
  );
}

function backgroundXml(
  background: AddSlideLayoutBackgroundInput | undefined,
  imageRelationshipId: RelationshipId | undefined,
): string {
  if (background === undefined) return "";
  if (background.kind === "solid") {
    return (
      `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${background.color.hex.toUpperCase()}"/>` +
      `</a:solidFill><a:effectLst/></p:bgPr></p:bg>`
    );
  }
  if (imageRelationshipId === undefined) {
    throw new Error("addSlideLayout: image relationship ID was not allocated");
  }
  return (
    `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1">` +
    `<a:blip r:embed="${imageRelationshipId}"/>` +
    `<a:stretch><a:fillRect/></a:stretch></a:blipFill>` +
    `<a:effectLst/></p:bgPr></p:bg>`
  );
}

function emptyGroupShapeProperties(): string {
  return (
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`
  );
}

function escapeXmlAttribute(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}
