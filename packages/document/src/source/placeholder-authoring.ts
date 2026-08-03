/** Native master/layout placeholder authoring for the initial text-placeholder subset. */

import { nextDrawingOrderingSlot, nextDrawingShapeId } from "./drawing-authoring-allocation.js";
import { sourceHandlesEqual } from "./edit-descriptors.js";
import type { SourceHandle } from "./handles.js";
import type { PptxSourceModel, PptxSourceModelAddShapeEdit } from "./pptx-source-model.js";
import type {
  AddShapeGeometryInput,
  AddTextBoxBodyPropertiesInput,
  AddTextBoxRunPropertiesInput,
} from "./shape-authoring.js";
import { addShape } from "./shape-authoring.js";
import { buildPlaceholderXml, parseShapeNodeXml } from "./shape-xml.js";
import type { SourcePlaceholder, SourceShapeNode, SourceTransform } from "./shapes.js";
import { asEmu } from "./units.js";

export type PlaceholderType = "title" | "ctrTitle" | "body" | "subTitle";

export interface AddPlaceholderInput {
  readonly type: PlaceholderType;
  readonly index: number;
  readonly name?: string;
  readonly orientation?: "horizontal" | "vertical";
  readonly size?: "full" | "half" | "quarter";
  readonly transform?: {
    readonly offsetX: SourceTransform["offsetX"];
    readonly offsetY: SourceTransform["offsetY"];
    readonly width: SourceTransform["width"];
    readonly height: SourceTransform["height"];
  };
  readonly geometry?: AddShapeGeometryInput;
  readonly promptText?: string;
  readonly promptProperties?: AddTextBoxRunPropertiesInput;
  readonly body?: AddTextBoxBodyPropertiesInput;
}

interface PlaceholderTarget {
  readonly kind: "layout" | "master";
  readonly index: number;
  readonly partPath: SourceHandle["partPath"];
  readonly shapes: readonly SourceShapeNode[];
}

export function addPlaceholder(
  source: PptxSourceModel,
  targetHandle: SourceHandle,
  input: AddPlaceholderInput,
): PptxSourceModel {
  const target = findPlaceholderTarget(source, targetHandle);
  if (target === undefined) {
    throw new Error(
      "addPlaceholder: layout or master handle was not found in PptxSourceModel source",
    );
  }
  assertPlaceholderInput(source, targetHandle, target, input);

  if (
    target.shapes.some((shape) => {
      const placeholder = sourceNodePlaceholder(shape);
      return placeholder !== undefined && (placeholder.index ?? 0) === input.index;
    })
  ) {
    throw new Error(`addPlaceholder: effective index '${input.index}' is already in use`);
  }

  if (target.kind === "layout" && input.transform === undefined) {
    const masterMatches = compatibleMasterPlaceholders(source, target, input.type);
    if (masterMatches.length !== 1 || sourceNodeTransform(masterMatches[0]) === undefined) {
      throw new Error(
        "addPlaceholder: layout transform may be omitted only with one compatible master placeholder transform",
      );
    }
  }

  const shapeIdValue = String(nextDrawingShapeId(source, target.shapes, target.partPath));
  const orderingSlot = nextDrawingOrderingSlot(target.shapes);
  const xml = buildPlaceholderXml({
    shapeId: shapeIdValue,
    name: input.name?.trim() || `${input.type} Placeholder ${shapeIdValue}`,
    placeholder: {
      type: input.type,
      index: input.index,
      ...(input.orientation !== undefined
        ? { orientation: input.orientation === "horizontal" ? "horz" : "vert" }
        : {}),
      ...(input.size !== undefined ? { size: input.size } : {}),
      ...(input.promptText !== undefined ? { hasCustomPrompt: true } : {}),
    },
    ...(input.transform !== undefined ? { transform: input.transform } : {}),
    ...(input.geometry !== undefined ? { geometry: input.geometry } : {}),
    ...(input.promptText !== undefined ? { promptText: input.promptText } : {}),
    ...(input.promptProperties !== undefined ? { promptProperties: input.promptProperties } : {}),
    ...(input.body !== undefined ? { body: input.body } : {}),
  });
  const shape = parseShapeNodeXml(xml, target.partPath, orderingSlot);
  const edit = {
    kind: "addShape",
    slidePartPath: target.partPath,
    shapeId: shapeIdValue,
    xml,
  } satisfies PptxSourceModelAddShapeEdit;
  return {
    ...withTargetShapes(source, target, [...target.shapes, shape]),
    edits: [...(source.edits ?? []), edit],
  };
}

function assertPlaceholderInput(
  source: PptxSourceModel,
  targetHandle: SourceHandle,
  target: PlaceholderTarget,
  input: AddPlaceholderInput,
): void {
  if (input === null || typeof input !== "object" || Array.isArray(input)) {
    throw new Error("addPlaceholder: input must be a placeholder object");
  }
  const allowedTypes =
    target.kind === "master"
      ? new Set<unknown>(["title", "body"])
      : new Set<unknown>(["title", "ctrTitle", "body", "subTitle"]);
  if (!allowedTypes.has(input.type)) {
    throw new Error(
      `addPlaceholder: type '${String(input.type)}' is unsupported for ${target.kind}`,
    );
  }
  if (!Number.isInteger(input.index) || input.index < 0 || input.index > 0xffffffff) {
    throw new Error("addPlaceholder: index must be an unsigned 32-bit integer");
  }
  if (input.name !== undefined && (typeof input.name !== "string" || input.name.trim() === "")) {
    throw new Error("addPlaceholder: name must be a non-empty string when provided");
  }
  if (
    input.orientation !== undefined &&
    input.orientation !== "horizontal" &&
    input.orientation !== "vertical"
  ) {
    throw new Error("addPlaceholder: orientation is unsupported");
  }
  if (
    input.size !== undefined &&
    input.size !== "full" &&
    input.size !== "half" &&
    input.size !== "quarter"
  ) {
    throw new Error("addPlaceholder: size is unsupported");
  }
  if (target.kind === "master" && input.transform === undefined) {
    throw new Error("addPlaceholder: master placeholders require a transform");
  }
  if (input.transform !== undefined) {
    if (
      input.transform === null ||
      typeof input.transform !== "object" ||
      Array.isArray(input.transform)
    ) {
      throw new Error("addPlaceholder: transform must be an object");
    }
    for (const [field, value] of Object.entries(input.transform)) {
      const positive = field === "width" || field === "height";
      if (typeof value !== "number" || !Number.isFinite(value) || (positive && value <= 0)) {
        throw new Error(
          `addPlaceholder: transform.${field} must be a finite ${positive ? "positive " : ""}EMU value`,
        );
      }
    }
    for (const field of ["offsetX", "offsetY", "width", "height"]) {
      if (!(field in input.transform)) {
        throw new Error(`addPlaceholder: transform.${field} is required`);
      }
    }
  }
  if (input.promptText !== undefined && typeof input.promptText !== "string") {
    throw new Error("addPlaceholder: promptText must be a string when provided");
  }
  if (input.promptProperties !== undefined && input.promptText === undefined) {
    throw new Error("addPlaceholder: promptProperties require promptText");
  }
  for (const value of [input.name, input.promptText]) {
    if (value !== undefined && /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(value)) {
      throw new Error("addPlaceholder: text contains an invalid XML character");
    }
  }

  // Reuse the existing public shape validators without retaining their immutable result.
  try {
    addShape(source, targetHandle, {
      geometry: input.geometry ?? { kind: "preset", preset: "rect" },
      offsetX: input.transform?.offsetX ?? asEmu(0),
      offsetY: input.transform?.offsetY ?? asEmu(0),
      width: input.transform?.width ?? asEmu(1),
      height: input.transform?.height ?? asEmu(1),
      text: input.promptProperties === undefined ? "" : undefined,
      ...(input.promptProperties !== undefined
        ? {
            paragraphs: [
              { runs: [{ text: input.promptText ?? "", properties: input.promptProperties }] },
            ],
          }
        : {}),
      ...(input.body !== undefined ? { body: input.body } : {}),
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(message.replace(/^addShape:/u, "addPlaceholder:"));
  }
}

function compatibleMasterPlaceholders(
  source: PptxSourceModel,
  target: PlaceholderTarget,
  type: PlaceholderType,
): SourceShapeNode[] {
  if (target.kind !== "layout") return [];
  const layout = source.slideLayouts[target.index];
  const master = source.slideMasters.find(
    (candidate) => candidate.partPath === layout.masterPartPath,
  );
  const masterType = type === "title" || type === "ctrTitle" ? "title" : "body";
  return (master?.shapes ?? []).filter(
    (shape) => sourceNodePlaceholder(shape)?.type === masterType,
  );
}

function sourceNodePlaceholder(shape: SourceShapeNode): SourcePlaceholder | undefined {
  switch (shape.kind) {
    case "shape":
    case "image":
    case "table":
    case "chart":
    case "smartArt":
    case "raw":
      return shape.placeholder;
    case "connector":
    case "group":
      return undefined;
  }
}

function sourceNodeTransform(shape: SourceShapeNode | undefined): SourceTransform | undefined {
  if (shape === undefined) return undefined;
  switch (shape.kind) {
    case "shape":
    case "connector":
    case "group":
    case "image":
    case "table":
    case "chart":
    case "smartArt":
      return shape.transform;
    case "raw":
      return undefined;
  }
}

function findPlaceholderTarget(
  source: PptxSourceModel,
  handle: SourceHandle,
): PlaceholderTarget | undefined {
  const layoutIndex = source.slideLayouts.findIndex((candidate) =>
    sourceHandlesEqual(candidate.handle, handle),
  );
  if (layoutIndex >= 0) {
    const layout = source.slideLayouts[layoutIndex];
    return { kind: "layout", index: layoutIndex, partPath: layout.partPath, shapes: layout.shapes };
  }
  const masterIndex = source.slideMasters.findIndex((candidate) =>
    sourceHandlesEqual(candidate.handle, handle),
  );
  if (masterIndex >= 0) {
    const master = source.slideMasters[masterIndex];
    return { kind: "master", index: masterIndex, partPath: master.partPath, shapes: master.shapes };
  }
  return undefined;
}

function withTargetShapes(
  source: PptxSourceModel,
  target: PlaceholderTarget,
  shapes: readonly SourceShapeNode[],
): PptxSourceModel {
  return target.kind === "layout"
    ? {
        ...source,
        slideLayouts: source.slideLayouts.map((candidate, index) =>
          index === target.index ? { ...candidate, shapes } : candidate,
        ),
      }
    : {
        ...source,
        slideMasters: source.slideMasters.map((candidate, index) =>
          index === target.index ? { ...candidate, shapes } : candidate,
        ),
      };
}
