import { XMLBuilder } from "fast-xml-parser";

import { sourceHandlesEqual } from "./edit-descriptors.js";
import { copyBytes, IMAGE_REL_TYPE, relativeTarget } from "./editing-shared.js";
import { asPartPath, type PartPath } from "./handles.js";
import type {
  ContentTypeDefault,
  MediaPart,
  PackageGraph,
  PartRelationships,
  PptxSourceModel,
  PptxSourceModelSetSlideBackgroundEdit,
  Relationship,
  SourceBackground,
  SourceHandle,
} from "./index.js";
import { nextNumberedPartPath, nextRelationshipId } from "./package-graph-mutations.js";
import { relationshipsPartPath } from "./package-paths.js";
import type { SourceColor, SourceGradientStop } from "./shapes.js";
import type { OoxmlAngle, OoxmlPercent } from "./units.js";

export type SlideBackgroundColorInput = { readonly kind: "srgb"; readonly hex: string };

export interface SlideBackgroundGradientStopInput {
  readonly position: OoxmlPercent;
  readonly color: SlideBackgroundColorInput;
}

export type SetSlideBackgroundInput =
  | { readonly kind: "solid"; readonly color: SlideBackgroundColorInput }
  | {
      readonly kind: "gradient";
      readonly gradientType: "linear";
      readonly stops: readonly SlideBackgroundGradientStopInput[];
      readonly angle: OoxmlAngle;
    }
  | {
      readonly kind: "gradient";
      readonly gradientType: "radial";
      readonly stops: readonly SlideBackgroundGradientStopInput[];
      readonly centerX?: OoxmlPercent;
      readonly centerY?: OoxmlPercent;
    }
  | { readonly kind: "image"; readonly bytes: Uint8Array };

interface DetectedImageType {
  readonly contentType: "image/png" | "image/jpeg";
  readonly extension: "png" | "jpeg";
}

const IMAGE_MEDIA_PREFIX = "ppt/media/image";
const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: true,
});

/** Sets a direct `p:bgPr` background on one slide. */
export function setSlideBackground(
  source: PptxSourceModel,
  slideHandle: SourceHandle,
  input: SetSlideBackgroundInput,
): PptxSourceModel {
  assertBackgroundInput(input);
  const slideIndex = source.slides.findIndex((slide) =>
    sourceHandlesEqual(slide.handle, slideHandle),
  );
  const slide = source.slides[slideIndex];
  if (slide === undefined) {
    throw new Error("setSlideBackground: slide handle was not found in PptxSourceModel source");
  }
  if (
    (source.edits ?? []).some(
      (edit) => edit.kind === "setSlideBackground" && edit.slidePartPath === slide.partPath,
    )
  ) {
    throw new Error("setSlideBackground: the slide already has a pending background edit");
  }

  if (input.kind !== "image") {
    const background = sourceBackground(input);
    const edit = {
      kind: "setSlideBackground",
      slidePartPath: slide.partPath,
      xml: buildBackgroundXml(input),
    } satisfies PptxSourceModelSetSlideBackgroundEdit;
    return {
      ...source,
      slides: source.slides.map((candidate, index) =>
        index === slideIndex ? { ...candidate, background } : candidate,
      ),
      edits: [...(source.edits ?? []), edit],
    };
  }

  const imageType = detectSupportedImageType(input.bytes);
  if (imageType === undefined) {
    throw new Error("setSlideBackground: unsupported or unknown image format");
  }
  const relationshipGroup = relationshipGroupForPart(source.packageGraph, slide.partPath);
  const relationshipId = nextRelationshipId(relationshipGroup.relationships);
  const mediaPartPath = nextNumberedPartPath(
    source.packageGraph,
    backgroundMediaPartPaths(source.edits ?? []),
    IMAGE_MEDIA_PREFIX,
    `.${imageType.extension}`,
  );
  const media: MediaPart = {
    partPath: mediaPartPath,
    contentType: imageType.contentType,
    bytes: copyBytes(input.bytes),
  };
  const relationship: Relationship = {
    id: relationshipId,
    type: IMAGE_REL_TYPE,
    target: relativeTarget(slide.partPath, mediaPartPath),
  };
  const background: SourceBackground = {
    kind: "fill",
    fill: { kind: "image", blipRelationshipId: relationshipId },
  };
  const edit = {
    kind: "setSlideBackground",
    slidePartPath: slide.partPath,
    relationshipId,
    mediaPartPath,
    contentType: imageType.contentType,
    xml: buildBackgroundXml(input, relationshipId),
  } satisfies PptxSourceModelSetSlideBackgroundEdit;

  return {
    ...source,
    slides: source.slides.map((candidate, index) =>
      index === slideIndex ? { ...candidate, background } : candidate,
    ),
    packageGraph: addImagePackageEntries(
      source.packageGraph,
      slide.partPath,
      media,
      imageType.extension,
      relationship,
    ),
    edits: [...(source.edits ?? []), edit],
  };
}

function sourceBackground(
  input: Exclude<SetSlideBackgroundInput, { readonly kind: "image" }>,
): SourceBackground {
  if (input.kind === "solid") {
    return {
      kind: "fill",
      fill: { kind: "solid", color: sourceColor(input.color) },
    };
  }
  const stops: readonly SourceGradientStop[] = input.stops.map((stop) => ({
    position: stop.position / 100000,
    color: sourceColor(stop.color),
  }));
  return input.gradientType === "linear"
    ? {
        kind: "fill",
        fill: { kind: "gradient", gradientType: "linear", stops, angle: input.angle },
      }
    : {
        kind: "fill",
        fill: {
          kind: "gradient",
          gradientType: "radial",
          stops,
          centerX: (input.centerX ?? 50000) / 100000,
          centerY: (input.centerY ?? 50000) / 100000,
        },
      };
}

function sourceColor(color: SlideBackgroundColorInput): SourceColor {
  return { kind: "srgb", hex: color.hex.toUpperCase() };
}

function buildBackgroundXml(input: SetSlideBackgroundInput, relationshipId?: string): string {
  return xmlBuilder.build({
    "p:bg": {
      "p:bgPr": {
        ...backgroundFillXml(input, relationshipId),
        "a:effectLst": {},
      },
    },
  });
}

function backgroundFillXml(
  input: SetSlideBackgroundInput,
  relationshipId?: string,
): Record<string, unknown> {
  switch (input.kind) {
    case "solid":
      return { "a:solidFill": colorXml(input.color) };
    case "image":
      if (relationshipId === undefined) {
        throw new Error("setSlideBackground: image relationship id was not allocated");
      }
      return {
        "a:blipFill": {
          "@_dpi": "0",
          "@_rotWithShape": "1",
          "a:blip": { "@_r:embed": relationshipId },
          "a:stretch": { "a:fillRect": {} },
        },
      };
    case "gradient":
      return {
        "a:gradFill": {
          "a:gsLst": {
            "a:gs": input.stops.map((stop) => ({
              "@_pos": String(stop.position),
              ...colorXml(stop.color),
            })),
          },
          ...(input.gradientType === "linear"
            ? { "a:lin": { "@_ang": String(input.angle), "@_scaled": "1" } }
            : { "a:path": radialPathXml(input.centerX ?? 50000, input.centerY ?? 50000) }),
        },
      };
  }
}

function radialPathXml(centerX: number, centerY: number): Record<string, unknown> {
  return {
    "@_path": "circle",
    "a:fillToRect": {
      "@_l": String(centerX),
      "@_t": String(centerY),
      "@_r": String(100000 - centerX),
      "@_b": String(100000 - centerY),
    },
  };
}

function colorXml(color: SlideBackgroundColorInput): Record<string, unknown> {
  return { "a:srgbClr": { "@_val": color.hex.toUpperCase() } };
}

function addImagePackageEntries(
  graph: PackageGraph,
  slidePartPath: PartPath,
  media: MediaPart,
  extension: string,
  relationship: Relationship,
): PackageGraph {
  const relationshipGroup = relationshipGroupForPart(graph, slidePartPath);
  const relationshipPartPath = asPartPath(relationshipsPartPath(slidePartPath));
  const hasRelationshipGroup = graph.relationships.some(
    (candidate) => candidate.sourcePartPath === slidePartPath,
  );
  const hasRelationshipPart = graph.parts.some((part) => part.partPath === relationshipPartPath);
  const needsRelationshipOverride =
    !hasRelationshipPart &&
    !graph.contentTypes.defaults.some(
      (entry) => entry.extension === "rels" && entry.contentType === RELS_CONTENT_TYPE,
    ) &&
    !graph.contentTypes.overrides.some((entry) => entry.partName === relationshipPartPath);

  return {
    ...graph,
    contentTypes: {
      ...graph.contentTypes,
      defaults: withImageContentTypeDefault(
        graph.contentTypes.defaults,
        extension,
        media.contentType,
      ),
      overrides: [
        ...graph.contentTypes.overrides,
        ...(needsRelationshipOverride
          ? [{ partName: relationshipPartPath, contentType: RELS_CONTENT_TYPE }]
          : []),
      ],
    },
    parts: [
      ...graph.parts,
      { partPath: media.partPath, contentType: media.contentType },
      ...(hasRelationshipPart
        ? []
        : [{ partPath: relationshipPartPath, contentType: RELS_CONTENT_TYPE }]),
    ],
    relationships: hasRelationshipGroup
      ? graph.relationships.map((candidate) =>
          candidate.sourcePartPath === slidePartPath
            ? { ...candidate, relationships: [...candidate.relationships, relationship] }
            : candidate,
        )
      : [...graph.relationships, { ...relationshipGroup, relationships: [relationship] }],
    media: [...graph.media, media],
  };
}

function relationshipGroupForPart(graph: PackageGraph, partPath: PartPath): PartRelationships {
  return (
    graph.relationships.find((candidate) => candidate.sourcePartPath === partPath) ?? {
      sourcePartPath: partPath,
      relationships: [],
    }
  );
}

function withImageContentTypeDefault(
  defaults: readonly ContentTypeDefault[],
  extension: string,
  contentType: string,
): readonly ContentTypeDefault[] {
  const existing = defaults.find((entry) => entry.extension === extension);
  if (existing === undefined) return [...defaults, { extension, contentType }];
  if (existing.contentType === contentType) return defaults;
  throw new Error(
    `setSlideBackground: content type default for extension '${extension}' already maps to '${existing.contentType}'`,
  );
}

function backgroundMediaPartPaths(
  edits: readonly { readonly kind: string; readonly mediaPartPath?: PartPath }[],
): readonly PartPath[] {
  return edits.flatMap((edit) =>
    edit.kind === "setSlideBackground" && edit.mediaPartPath !== undefined
      ? [edit.mediaPartPath]
      : [],
  );
}

function assertBackgroundInput(input: unknown): asserts input is SetSlideBackgroundInput {
  if (!isRecord(input)) {
    throw new Error("setSlideBackground: input must be a background object");
  }
  const value = input;
  switch (value.kind) {
    case "solid":
      assertColor(value.color, "color");
      return;
    case "image":
      if (!(value.bytes instanceof Uint8Array)) {
        throw new Error("setSlideBackground: bytes must be a Uint8Array");
      }
      return;
    case "gradient":
      assertGradientStops(value.stops);
      if (value.gradientType === "linear") {
        if (!Number.isInteger(value.angle)) {
          throw new Error("setSlideBackground: angle must be an integer OOXML angle");
        }
        return;
      }
      if (value.gradientType !== "radial") {
        throw new Error("setSlideBackground: gradientType must be linear or radial");
      }
      assertOptionalPercent(value.centerX, "centerX");
      assertOptionalPercent(value.centerY, "centerY");
      return;
    default:
      throw new Error("setSlideBackground: background kind is not supported");
  }
}

function assertGradientStops(
  stops: unknown,
): asserts stops is readonly SlideBackgroundGradientStopInput[] {
  if (!isUnknownArray(stops) || stops.length < 2) {
    throw new Error("setSlideBackground: gradient stops must contain at least two entries");
  }
  let previous = -1;
  for (const [index, stop] of stops.entries()) {
    if (!isRecord(stop)) {
      throw new Error(`setSlideBackground: stops[${index}] must be a gradient stop object`);
    }
    const value = stop;
    assertRequiredPercent(value.position, `stops[${index}].position`);
    const position = value.position;
    if (position < previous) {
      throw new Error("setSlideBackground: gradient stop positions must be in ascending order");
    }
    assertColor(value.color, `stops[${index}].color`);
    previous = position;
  }
}

function assertColor(color: unknown, field: string): asserts color is SlideBackgroundColorInput {
  if (!isRecord(color)) {
    throw new Error(`setSlideBackground: ${field} must be a 6-digit srgb color`);
  }
  const value = color;
  if (
    value.kind !== "srgb" ||
    typeof value.hex !== "string" ||
    !/^[0-9A-Fa-f]{6}$/.test(value.hex)
  ) {
    throw new Error(`setSlideBackground: ${field} must be a 6-digit srgb color`);
  }
}

function isUnknownArray(value: unknown): value is unknown[] {
  return Array.isArray(value);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function assertRequiredPercent(value: unknown, field: string): asserts value is OoxmlPercent {
  if (!Number.isInteger(value) || Number(value) < 0 || Number(value) > 100000) {
    throw new Error(`setSlideBackground: ${field} must be an integer from 0 to 100000`);
  }
}

function assertOptionalPercent(
  value: unknown,
  field: string,
): asserts value is OoxmlPercent | undefined {
  if (value !== undefined) assertRequiredPercent(value, field);
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
