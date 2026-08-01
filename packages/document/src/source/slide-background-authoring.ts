import { XMLBuilder } from "fast-xml-parser";

import { sourceHandlesEqual } from "./edit-descriptors.js";
import { copyBytes, IMAGE_REL_TYPE, relativeTarget } from "./editing-shared.js";
import type { PartPath } from "./handles.js";
import { detectSupportedImageType } from "./image-type.js";
import type {
  MediaPart,
  PackageGraph,
  PartRelationships,
  PptxSourceModel,
  PptxSourceModelSetBackgroundEdit,
  PptxSourceModelSetSlideBackgroundEdit,
  Relationship,
  SourceBackground,
  SourceHandle,
} from "./index.js";
import {
  addMediaPartRelationship,
  nextNumberedPartPath,
  nextRelationshipId,
} from "./package-graph-mutations.js";
import type { SourceColor, SourceGradientStop } from "./shapes.js";
import type { OoxmlAngle, OoxmlPercent } from "./units.js";

export type BackgroundColorInput = { readonly kind: "srgb"; readonly hex: string };

export interface BackgroundGradientStopInput {
  readonly position: OoxmlPercent;
  readonly color: BackgroundColorInput;
}

export type SetBackgroundInput =
  | { readonly kind: "solid"; readonly color: BackgroundColorInput }
  | {
      readonly kind: "gradient";
      readonly gradientType: "linear";
      readonly stops: readonly BackgroundGradientStopInput[];
      readonly angle: OoxmlAngle;
    }
  | {
      readonly kind: "gradient";
      readonly gradientType: "radial";
      readonly stops: readonly BackgroundGradientStopInput[];
      readonly centerX?: OoxmlPercent;
      readonly centerY?: OoxmlPercent;
    }
  | { readonly kind: "image"; readonly bytes: Uint8Array };

/** @deprecated Use {@link BackgroundColorInput}. */
export type SlideBackgroundColorInput = BackgroundColorInput;
/** @deprecated Use {@link BackgroundGradientStopInput}. */
export type SlideBackgroundGradientStopInput = BackgroundGradientStopInput;
/** @deprecated Use {@link SetBackgroundInput}. */
export type SetSlideBackgroundInput = SetBackgroundInput;

const IMAGE_MEDIA_PREFIX = "ppt/media/image";

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: true,
});

/** Sets a direct `p:bgPr` background on a slide, layout, or master. */
export function setBackground(
  source: PptxSourceModel,
  targetHandle: SourceHandle,
  input: SetBackgroundInput,
): PptxSourceModel {
  return updateBackground(source, targetHandle, input, "setBackground");
}

/** Removes the direct background so the target resumes inheriting from its parent layer. */
export function clearBackground(
  source: PptxSourceModel,
  targetHandle: SourceHandle,
): PptxSourceModel {
  return updateBackground(source, targetHandle, undefined, "clearBackground");
}

/** Sets a direct background on one slide. Retained for source compatibility. */
export function setSlideBackground(
  source: PptxSourceModel,
  slideHandle: SourceHandle,
  input: SetSlideBackgroundInput,
): PptxSourceModel {
  if (!source.slides.some((slide) => sourceHandlesEqual(slide.handle, slideHandle))) {
    throw new Error("setSlideBackground: slide handle was not found in PptxSourceModel source");
  }
  return updateBackground(source, slideHandle, input, "setSlideBackground");
}

function updateBackground(
  source: PptxSourceModel,
  targetHandle: SourceHandle,
  input: SetBackgroundInput | undefined,
  operationName: "setBackground" | "clearBackground" | "setSlideBackground",
): PptxSourceModel {
  if (input !== undefined) assertBackgroundInput(input, operationName);
  const target = findBackgroundTarget(source, targetHandle);
  if (target === undefined) {
    throw new Error(
      `${operationName}: slide, layout, or master handle was not found in PptxSourceModel source`,
    );
  }
  if (
    (source.edits ?? []).some(
      (edit) =>
        (edit.kind === "setBackground" && edit.targetPartPath === target.partPath) ||
        (edit.kind === "setSlideBackground" && edit.slidePartPath === target.partPath),
    )
  ) {
    throw new Error(`${operationName}: the target already has a pending background edit`);
  }

  if (input === undefined) {
    if (target.background === undefined) return source;
    const edit = {
      kind: "setBackground",
      targetPartPath: target.partPath,
    } satisfies PptxSourceModelSetBackgroundEdit;
    return replaceTargetBackground(source, target, undefined, edit);
  }

  if (input.kind !== "image") {
    const background = sourceBackground(input);
    const xml = buildBackgroundXml(input);
    const edit =
      operationName === "setSlideBackground"
        ? ({
            kind: "setSlideBackground",
            slidePartPath: target.partPath,
            xml,
          } satisfies PptxSourceModelSetSlideBackgroundEdit)
        : ({
            kind: "setBackground",
            targetPartPath: target.partPath,
            xml,
          } satisfies PptxSourceModelSetBackgroundEdit);
    return replaceTargetBackground(source, target, background, edit);
  }

  const imageType = detectSupportedImageType(input.bytes);
  if (imageType === undefined) {
    throw new Error(`${operationName}: unsupported or unknown image format`);
  }
  const relationshipGroup = relationshipGroupForPart(source.packageGraph, target.partPath);
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
    target: relativeTarget(target.partPath, mediaPartPath),
  };
  const background: SourceBackground = {
    kind: "fill",
    fill: { kind: "image", blipRelationshipId: relationshipId },
  };
  const xml = buildBackgroundXml(input, relationshipId);
  const edit =
    operationName === "setSlideBackground"
      ? ({
          kind: "setSlideBackground",
          slidePartPath: target.partPath,
          relationshipId,
          mediaPartPath,
          contentType: imageType.contentType,
          xml,
        } satisfies PptxSourceModelSetSlideBackgroundEdit)
      : ({
          kind: "setBackground",
          targetPartPath: target.partPath,
          relationshipId,
          mediaPartPath,
          contentType: imageType.contentType,
          xml,
        } satisfies PptxSourceModelSetBackgroundEdit);

  const withBackground = replaceTargetBackground(source, target, background, edit);
  return {
    ...withBackground,
    packageGraph: addMediaPartRelationship(source.packageGraph, {
      ownerPartPath: target.partPath,
      media,
      extension: imageType.extension,
      relationship,
      contentTypeDefaultConflictError: (existingContentType) =>
        new Error(
          `${operationName}: content type default for extension '${imageType.extension}' already maps to '${existingContentType}'`,
        ),
    }),
  };
}

type BackgroundTarget =
  | {
      readonly layer: "slide";
      readonly index: number;
      readonly partPath: PartPath;
      readonly background?: SourceBackground;
    }
  | {
      readonly layer: "layout";
      readonly index: number;
      readonly partPath: PartPath;
      readonly background?: SourceBackground;
    }
  | {
      readonly layer: "master";
      readonly index: number;
      readonly partPath: PartPath;
      readonly background?: SourceBackground;
    };

function findBackgroundTarget(
  source: PptxSourceModel,
  handle: SourceHandle,
): BackgroundTarget | undefined {
  const slideIndex = source.slides.findIndex((item) => sourceHandlesEqual(item.handle, handle));
  const slide = source.slides[slideIndex];
  if (slide !== undefined)
    return {
      layer: "slide",
      index: slideIndex,
      partPath: slide.partPath,
      background: slide.background,
    };
  const layoutIndex = source.slideLayouts.findIndex((item) =>
    sourceHandlesEqual(item.handle, handle),
  );
  const layout = source.slideLayouts[layoutIndex];
  if (layout !== undefined)
    return {
      layer: "layout",
      index: layoutIndex,
      partPath: layout.partPath,
      background: layout.background,
    };
  const masterIndex = source.slideMasters.findIndex((item) =>
    sourceHandlesEqual(item.handle, handle),
  );
  const master = source.slideMasters[masterIndex];
  return master === undefined
    ? undefined
    : {
        layer: "master",
        index: masterIndex,
        partPath: master.partPath,
        background: master.background,
      };
}

function replaceTargetBackground(
  source: PptxSourceModel,
  target: BackgroundTarget,
  background: SourceBackground | undefined,
  edit: PptxSourceModelSetBackgroundEdit | PptxSourceModelSetSlideBackgroundEdit,
): PptxSourceModel {
  const slides =
    target.layer === "slide"
      ? source.slides.map((item, index) => {
          if (index !== target.index) return item;
          if (background !== undefined) return { ...item, background };
          const { background: _background, ...rest } = item;
          void _background;
          return rest;
        })
      : source.slides;
  const slideLayouts =
    target.layer === "layout"
      ? source.slideLayouts.map((item, index) => {
          if (index !== target.index) return item;
          if (background !== undefined) return { ...item, background };
          const { background: _background, ...rest } = item;
          void _background;
          return rest;
        })
      : source.slideLayouts;
  const slideMasters =
    target.layer === "master"
      ? source.slideMasters.map((item, index) => {
          if (index !== target.index) return item;
          if (background !== undefined) return { ...item, background };
          const { background: _background, ...rest } = item;
          void _background;
          return rest;
        })
      : source.slideMasters;
  return {
    ...source,
    slides,
    slideLayouts,
    slideMasters,
    edits: [...(source.edits ?? []), edit],
  };
}

function sourceBackground(
  input: Exclude<SetBackgroundInput, { readonly kind: "image" }>,
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

function sourceColor(color: BackgroundColorInput): SourceColor {
  return { kind: "srgb", hex: color.hex.toUpperCase() };
}

function buildBackgroundXml(input: SetBackgroundInput, relationshipId?: string): string {
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
  input: SetBackgroundInput,
  relationshipId?: string,
): Record<string, unknown> {
  switch (input.kind) {
    case "solid":
      return { "a:solidFill": colorXml(input.color) };
    case "image":
      if (relationshipId === undefined) {
        throw new Error("setBackground: image relationship id was not allocated");
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

function colorXml(color: BackgroundColorInput): Record<string, unknown> {
  return { "a:srgbClr": { "@_val": color.hex.toUpperCase() } };
}

function relationshipGroupForPart(graph: PackageGraph, partPath: PartPath): PartRelationships {
  return (
    graph.relationships.find((candidate) => candidate.sourcePartPath === partPath) ?? {
      sourcePartPath: partPath,
      relationships: [],
    }
  );
}

function backgroundMediaPartPaths(
  edits: readonly { readonly kind: string; readonly mediaPartPath?: PartPath }[],
): readonly PartPath[] {
  return edits.flatMap((edit) =>
    (edit.kind === "setBackground" || edit.kind === "setSlideBackground") &&
    edit.mediaPartPath !== undefined
      ? [edit.mediaPartPath]
      : [],
  );
}

function assertBackgroundInput(
  input: unknown,
  operationName: string,
): asserts input is SetBackgroundInput {
  if (!isRecord(input)) {
    throw new Error(`${operationName}: input must be a background object`);
  }
  const value = input;
  switch (value.kind) {
    case "solid":
      assertColor(value.color, "color", operationName);
      return;
    case "image":
      if (!(value.bytes instanceof Uint8Array)) {
        throw new Error(`${operationName}: bytes must be a Uint8Array`);
      }
      return;
    case "gradient":
      assertGradientStops(value.stops, operationName);
      if (value.gradientType === "linear") {
        if (!Number.isInteger(value.angle)) {
          throw new Error(`${operationName}: angle must be an integer OOXML angle`);
        }
        return;
      }
      if (value.gradientType !== "radial") {
        throw new Error(`${operationName}: gradientType must be linear or radial`);
      }
      assertOptionalPercent(value.centerX, "centerX", operationName);
      assertOptionalPercent(value.centerY, "centerY", operationName);
      return;
    default:
      throw new Error(`${operationName}: background kind is not supported`);
  }
}

function assertGradientStops(
  stops: unknown,
  operationName: string,
): asserts stops is readonly BackgroundGradientStopInput[] {
  if (!isUnknownArray(stops) || stops.length < 2) {
    throw new Error(`${operationName}: gradient stops must contain at least two entries`);
  }
  let previous = -1;
  for (const [index, stop] of stops.entries()) {
    if (!isRecord(stop)) {
      throw new Error(`${operationName}: stops[${index}] must be a gradient stop object`);
    }
    const value = stop;
    assertRequiredPercent(value.position, `stops[${index}].position`, operationName);
    const position = value.position;
    if (position < previous) {
      throw new Error(`${operationName}: gradient stop positions must be in ascending order`);
    }
    assertColor(value.color, `stops[${index}].color`, operationName);
    previous = position;
  }
}

function assertColor(
  color: unknown,
  field: string,
  operationName: string,
): asserts color is BackgroundColorInput {
  if (!isRecord(color)) {
    throw new Error(`${operationName}: ${field} must be a 6-digit srgb color`);
  }
  const value = color;
  if (
    value.kind !== "srgb" ||
    typeof value.hex !== "string" ||
    !/^[0-9A-Fa-f]{6}$/.test(value.hex)
  ) {
    throw new Error(`${operationName}: ${field} must be a 6-digit srgb color`);
  }
}

function isUnknownArray(value: unknown): value is unknown[] {
  return Array.isArray(value);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function assertRequiredPercent(
  value: unknown,
  field: string,
  operationName: string,
): asserts value is OoxmlPercent {
  if (!Number.isInteger(value) || Number(value) < 0 || Number(value) > 100000) {
    throw new Error(`${operationName}: ${field} must be an integer from 0 to 100000`);
  }
}

function assertOptionalPercent(
  value: unknown,
  field: string,
  operationName: string,
): asserts value is OoxmlPercent | undefined {
  if (value !== undefined) assertRequiredPercent(value, field, operationName);
}
