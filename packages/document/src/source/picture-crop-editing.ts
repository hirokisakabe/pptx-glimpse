import { sourceHandlesEqual } from "./edit-descriptors.js";
import type {
  OoxmlPercent,
  PptxSourceModel,
  PptxSourceModelEdit,
  SourceHandle,
  SourceImage,
  SourceImageCrop,
} from "./index.js";
import {
  assertNotAlternateContentTarget,
  replaceSlideShapeNode,
  requireUniqueSlideShapeTarget,
} from "./shape-editing.js";
import { asOoxmlPercent } from "./units.js";

/**
 * Crop insets in OOXML percentage units (`100000` = 100%). Omitted edges mean zero.
 * Values are rejected rather than clamped.
 */
export interface SetPictureCropInput {
  readonly left?: OoxmlPercent;
  readonly top?: OoxmlPercent;
  readonly right?: OoxmlPercent;
  readonly bottom?: OoxmlPercent;
}

export function setPictureCrop(
  source: PptxSourceModel,
  handle: SourceHandle,
  input: SetPictureCropInput,
): PptxSourceModel {
  const crop = normalizeCrop(input, "setPictureCrop");
  return updatePictureCrop(source, handle, crop, "setPictureCrop");
}

export function clearPictureCrop(source: PptxSourceModel, handle: SourceHandle): PptxSourceModel {
  return updatePictureCrop(source, handle, undefined, "clearPictureCrop");
}

function updatePictureCrop(
  source: PptxSourceModel,
  handle: SourceHandle,
  crop: SourceImageCrop | undefined,
  operationName: "setPictureCrop" | "clearPictureCrop",
): PptxSourceModel {
  if (handle.nodeId === undefined) {
    throw new Error(`${operationName}: picture crop edit requires a node id`);
  }
  const target = requireUniqueSlideShapeTarget(source, handle, operationName);
  assertNotAlternateContentTarget(target, operationName);
  if (target.node.kind !== "image") {
    throw new Error(`${operationName}: shape handle does not reference a pic image shape`);
  }
  if (target.node.blipFillMode !== "stretch") {
    throw new Error(
      `${operationName}: only picture blipFill with exactly one stretch is supported`,
    );
  }
  if (cropEqual(target.node.crop, crop)) return source;

  const slides = replaceSlideShapeNode(source, handle, (shape) =>
    shape.kind === "image"
      ? crop === undefined
        ? withoutCrop(shape)
        : ({ ...shape, crop } satisfies SourceImage)
      : shape,
  );

  return {
    ...source,
    slides,
    edits: appendPictureCropEdit(source.edits ?? [], handle, crop),
  };
}

function withoutCrop(image: SourceImage): Omit<SourceImage, "crop"> {
  const { crop: _crop, ...rest } = image;
  void _crop;
  return rest;
}

function normalizeCrop(
  input: SetPictureCropInput,
  operationName: string,
): SourceImageCrop | undefined {
  const left = normalizeInset(input.left, operationName, "left");
  const top = normalizeInset(input.top, operationName, "top");
  const right = normalizeInset(input.right, operationName, "right");
  const bottom = normalizeInset(input.bottom, operationName, "bottom");
  if (left + right >= 100000) {
    throw new Error(`${operationName}: left + right must be less than 100000`);
  }
  if (top + bottom >= 100000) {
    throw new Error(`${operationName}: top + bottom must be less than 100000`);
  }
  const crop: SourceImageCrop = {
    ...(left !== 0 ? { left } : {}),
    ...(top !== 0 ? { top } : {}),
    ...(right !== 0 ? { right } : {}),
    ...(bottom !== 0 ? { bottom } : {}),
  };
  return Object.keys(crop).length === 0 ? undefined : crop;
}

function normalizeInset(
  value: OoxmlPercent | undefined,
  operationName: string,
  fieldName: keyof SetPictureCropInput,
): OoxmlPercent {
  if (value === undefined) return asOoxmlPercent(0);
  if (!Number.isInteger(value) || value < 0 || value > 100000) {
    throw new Error(
      `${operationName}: ${fieldName} must be an integer OOXML percentage from 0 through 100000`,
    );
  }
  return value;
}

function cropEqual(left: SourceImageCrop | undefined, right: SourceImageCrop | undefined): boolean {
  if (right === undefined) return left === undefined;
  if (left === undefined) return false;
  return (
    (left.left ?? 0) === (right.left ?? 0) &&
    (left.top ?? 0) === (right.top ?? 0) &&
    (left.right ?? 0) === (right.right ?? 0) &&
    (left.bottom ?? 0) === (right.bottom ?? 0)
  );
}

function appendPictureCropEdit(
  edits: readonly PptxSourceModelEdit[],
  handle: SourceHandle,
  crop: SourceImageCrop | undefined,
): PptxSourceModelEdit[] {
  const retained = edits.filter(
    (edit) => edit.kind !== "updatePictureCrop" || !sourceHandlesEqual(edit.handle, handle),
  );
  return [
    ...retained,
    {
      kind: "updatePictureCrop",
      handle,
      ...(crop !== undefined ? { crop } : {}),
    },
  ];
}
