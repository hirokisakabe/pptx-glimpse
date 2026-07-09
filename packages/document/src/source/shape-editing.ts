/**
 * Shape editing and authoring operations for PptxSourceModel.
 *
 * Preset-geometry shape authoring uses the same edit-time XML finalization model as
 * text boxes and connectors. The API deliberately keeps the preset name as a
 * pass-through string because OOXML preset coverage is broader than the initial pom
 * swap needs, while styling inputs stay inside the typed subset that the source reader
 * can parse back into `SourceShape`.
 *
 * Text-box authoring intentionally stays on the existing `addTextBox` operation
 * instead of introducing a second rich-text builder. The adopted API keeps the simple
 * `text` shortcut for one unformatted run and adds structured `paragraphs`, `body`,
 * and `rotation` options for formatted text boxes. This mirrors the source model
 * hierarchy (`a:rPr` / `a:pPr` / `a:bodyPr` / `a:xfrm`) while preserving edit-time XML
 * finalization. The alternatives considered were a fluent text builder and a raw OOXML
 * escape hatch. A fluent builder would add another mutable authoring layer before the
 * source model has broader primitive coverage, and raw OOXML would shift validity and
 * escaping responsibility to consumers, so both were rejected for this slice.
 */

import {
  editInsertedShape,
  editReservedShapeId,
  editTargetsShape,
  sourceHandlesEqual,
} from "./edit-descriptors.js";
import { asSourceNodeId } from "./handles.js";
import type {
  ConnectorPresetGeometry,
  EditableShapeFill,
  EditableShapeOutline,
  Emu,
  PartPath,
  PptxSourceModel,
  PptxSourceModelAddConnectorEdit,
  PptxSourceModelAddShapeEdit,
  PptxSourceModelEdit,
  PptxSourceModelShapeOutlineEdit,
  SourceArrowEndpoint,
  SourceConnector,
  SourceFill,
  SourceHandle,
  SourceNodeId,
  SourceOutline,
  SourceShape,
  SourceShapeNode,
  SourceSlide,
} from "./index.js";
import { nextNumberedName } from "./package-graph-mutations.js";
import type {
  ShapeColorInput,
  ShapeEffectsInput,
  ShapeFillInput,
  ShapeGlowInput,
  ShapeGradientFillInput,
  ShapeOutlineInput,
  TextBoxBaselineInput,
  TextBoxBodyPropertiesInput,
  TextBoxColorInput,
  TextBoxGlowInput,
  TextBoxGradientFillInput,
  TextBoxOutlineInput,
  TextBoxParagraphInput,
  TextBoxParagraphPropertiesInput,
  TextBoxRunInput,
  TextBoxRunPropertiesInput,
  TextBoxUnderlineInput,
  TextBoxUnderlineStyle,
} from "./shape-xml.js";
import {
  buildConnectorXml,
  buildShapeXml,
  buildTextBoxXml,
  parseShapeNodeXml,
} from "./shape-xml.js";

export type AddTextBoxBaselineInput = TextBoxBaselineInput;
export type AddTextBoxBodyPropertiesInput = TextBoxBodyPropertiesInput;
export type AddTextBoxColorInput = TextBoxColorInput;
export type AddTextBoxGlowInput = TextBoxGlowInput;
export type AddTextBoxGradientFillInput = TextBoxGradientFillInput;
export type AddTextBoxOutlineInput = TextBoxOutlineInput;
export type AddTextBoxParagraphInput = TextBoxParagraphInput;
export type AddTextBoxParagraphPropertiesInput = TextBoxParagraphPropertiesInput;
export type AddTextBoxRunInput = TextBoxRunInput;
export type AddTextBoxRunPropertiesInput = TextBoxRunPropertiesInput;
export type AddTextBoxUnderlineInput = TextBoxUnderlineInput;
export type AddTextBoxUnderlineStyle = TextBoxUnderlineStyle;
export type AddShapeBodyPropertiesInput = TextBoxBodyPropertiesInput;
export type AddShapeColorInput = ShapeColorInput;
export type AddShapeEffectsInput = ShapeEffectsInput;
export type AddShapeFillInput = ShapeFillInput;
export type AddShapeGlowInput = ShapeGlowInput;
export type AddShapeGradientFillInput = ShapeGradientFillInput;
export type AddShapeOutlineInput = ShapeOutlineInput;
export type AddShapeParagraphInput = TextBoxParagraphInput;
export type AddShapeParagraphPropertiesInput = TextBoxParagraphPropertiesInput;
export type AddShapeRunInput = TextBoxRunInput;
export type AddShapeRunPropertiesInput = TextBoxRunPropertiesInput;

type TransformableShapeNode = Exclude<SourceShapeNode, { readonly kind: "raw" }>;

const CONNECTOR_PRESETS: ReadonlySet<ConnectorPresetGeometry> = new Set([
  "straightConnector1",
  "bentConnector3",
  "curvedConnector3",
]);
const ARROW_TYPES = new Set(["triangle", "stealth", "diamond", "oval", "arrow"]);
const ARROW_SIZES = new Set(["sm", "med", "lg"]);
const DASH_STYLES = new Set([
  "solid",
  "dash",
  "dot",
  "dashDot",
  "lgDash",
  "lgDashDot",
  "sysDash",
  "sysDot",
]);
const UNDERLINE_STYLES: ReadonlySet<string> = new Set<AddTextBoxUnderlineStyle>([
  "sng",
  "dbl",
  "heavy",
  "dotted",
  "dottedHeavy",
  "dash",
  "dashHeavy",
  "dashLong",
  "dashLongHeavy",
  "dotDash",
  "dotDashHeavy",
  "dotDotDash",
  "dotDotDashHeavy",
  "wavy",
  "wavyHeavy",
  "wavyDbl",
  "none",
]);

export interface UpdateShapeTransformInput {
  readonly offsetX: Emu;
  readonly offsetY: Emu;
  readonly width: Emu;
  readonly height: Emu;
}

export interface AddTextBoxInput extends UpdateShapeTransformInput {
  readonly text?: string;
  readonly paragraphs?: readonly AddTextBoxParagraphInput[];
  readonly body?: AddTextBoxBodyPropertiesInput;
  readonly rotation?: NonNullable<SourceShape["transform"]>["rotation"];
  readonly name?: string;
}

export interface AddShapeInput extends UpdateShapeTransformInput {
  readonly preset: string;
  readonly fill?: AddShapeFillInput;
  readonly outline?: AddShapeOutlineInput;
  readonly effects?: AddShapeEffectsInput;
  readonly text?: string;
  readonly paragraphs?: readonly AddShapeParagraphInput[];
  readonly body?: AddShapeBodyPropertiesInput;
  readonly rotation?: NonNullable<SourceShape["transform"]>["rotation"];
  readonly name?: string;
}

export interface AddConnectorConnectionEndpointInput {
  readonly shapeHandle: SourceHandle;
  readonly connectionSiteIndex: number;
}

export interface AddConnectorOutlineInput {
  readonly headEnd?: SourceArrowEndpoint;
  readonly tailEnd?: SourceArrowEndpoint;
}

export interface AddConnectorInput extends UpdateShapeTransformInput {
  readonly preset: ConnectorPresetGeometry;
  readonly start?: AddConnectorConnectionEndpointInput;
  readonly end?: AddConnectorConnectionEndpointInput;
  readonly name?: string;
  readonly outline?: AddConnectorOutlineInput;
}

type StyleEditableShapeNode = SourceShape | SourceConnector;

export function findShapeNodeBySourceHandle(
  source: PptxSourceModel,
  handle: SourceHandle,
): SourceShapeNode | undefined {
  for (const slide of source.slides) {
    const shape = findShapeNodeInTree(slide.shapes, handle);
    if (shape !== undefined) return shape;
  }
  return undefined;
}

export function updateShapeTransform(
  source: PptxSourceModel,
  handle: SourceHandle,
  transform: UpdateShapeTransformInput,
): PptxSourceModel {
  if (handle.nodeId === undefined) {
    throw new Error("updateShapeTransform: shape transform edit requires a node id");
  }

  let matched = false;
  let changed = false;

  const slides = source.slides.map((slide) => {
    let slideChanged = false;
    const shapes = slide.shapes.map((shape) => {
      if (!sourceHandlesEqual(shape.handle, handle)) return shape;
      matched = true;
      if (hasAlternateContentSidecar(shape)) {
        throw new Error("updateShapeTransform: shapes inside AlternateContent are not supported");
      }
      if (!hasEditableTransform(shape)) {
        throw new Error("updateShapeTransform: shape handle does not reference a shape with xfrm");
      }
      if (shapeTransformPositionAndSizeEqual(shape.transform, transform)) return shape;
      changed = true;
      slideChanged = true;
      return {
        ...shape,
        transform: {
          ...shape.transform,
          offsetX: transform.offsetX,
          offsetY: transform.offsetY,
          width: transform.width,
          height: transform.height,
        },
      };
    });
    return slideChanged ? { ...slide, shapes } : slide;
  });

  if (!matched) {
    if (source.slides.some((slide) => hasNestedShapeNodeWithHandle(slide.shapes, handle))) {
      throw new Error("updateShapeTransform: nested group shape editing is not supported");
    }
    throw new Error("updateShapeTransform: shape handle was not found in PptxSourceModel source");
  }
  if (!changed) return source;

  return {
    ...source,
    slides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "updateShapeTransform",
        handle,
        offsetX: transform.offsetX,
        offsetY: transform.offsetY,
        width: transform.width,
        height: transform.height,
      },
    ],
  };
}

export function setShapeFill(
  source: PptxSourceModel,
  handle: SourceHandle,
  fill: EditableShapeFill,
): PptxSourceModel {
  assertEditableShapeFill(fill, "setShapeFill");
  if (handle.nodeId === undefined) {
    throw new Error("setShapeFill: shape fill edit requires a node id");
  }

  let matched = false;
  let changed = false;

  const slides = source.slides.map((slide) => {
    let slideChanged = false;
    const shapes = slide.shapes.map((shape) => {
      if (!sourceHandlesEqual(shape.handle, handle)) return shape;
      matched = true;
      if (shape.kind !== "shape") {
        throw new Error("setShapeFill: only top-level sp shapes support fill edits");
      }
      if (hasAlternateContentSidecar(shape)) {
        throw new Error("setShapeFill: shapes inside AlternateContent are not supported");
      }
      const nextFill = toSourceFill(fill);
      if (sourceFillEqual(shape.fill, nextFill)) return shape;
      changed = true;
      slideChanged = true;
      return {
        ...shape,
        fill: nextFill,
      } satisfies SourceShape;
    });
    return slideChanged ? { ...slide, shapes } : slide;
  });

  if (!matched) {
    if (source.slides.some((slide) => hasNestedShapeNodeWithHandle(slide.shapes, handle))) {
      throw new Error("setShapeFill: nested group shape editing is not supported");
    }
    throw new Error("setShapeFill: shape handle was not found in PptxSourceModel source");
  }
  if (!changed) return source;

  return {
    ...source,
    slides,
    edits: appendShapeFillEdit(source.edits ?? [], handle, fill),
  };
}

export function setShapeOutline(
  source: PptxSourceModel,
  handle: SourceHandle,
  outline: EditableShapeOutline,
): PptxSourceModel {
  assertEditableShapeOutline(outline, "setShapeOutline");
  if (handle.nodeId === undefined) {
    throw new Error("setShapeOutline: shape outline edit requires a node id");
  }

  let matched = false;
  let changed = false;

  const slides = source.slides.map((slide) => {
    let slideChanged = false;
    const shapes = slide.shapes.map((shape) => {
      if (!sourceHandlesEqual(shape.handle, handle)) return shape;
      matched = true;
      if (shape.kind !== "shape" && shape.kind !== "connector") {
        throw new Error(
          "setShapeOutline: only top-level sp and cxnSp shapes support outline edits",
        );
      }
      if (hasAlternateContentSidecar(shape)) {
        throw new Error("setShapeOutline: shapes inside AlternateContent are not supported");
      }
      const nextOutline = patchSourceOutline(shape.outline, outline);
      if (sourceOutlineEqual(shape.outline, nextOutline)) return shape;
      changed = true;
      slideChanged = true;
      return {
        ...shape,
        outline: nextOutline,
      } satisfies StyleEditableShapeNode;
    });
    return slideChanged ? { ...slide, shapes } : slide;
  });

  if (!matched) {
    if (source.slides.some((slide) => hasNestedShapeNodeWithHandle(slide.shapes, handle))) {
      throw new Error("setShapeOutline: nested group shape editing is not supported");
    }
    throw new Error("setShapeOutline: shape handle was not found in PptxSourceModel source");
  }
  if (!changed) return source;

  return {
    ...source,
    slides,
    edits: appendShapeOutlineEdit(source.edits ?? [], handle, outline),
  };
}

export function addTextBox(
  source: PptxSourceModel,
  slideHandle: SourceHandle,
  input: AddTextBoxInput,
): PptxSourceModel {
  assertTextBoxInput(input);
  const slideIndex = source.slides.findIndex((slide) =>
    sourceHandlesEqual(slide.handle, slideHandle),
  );
  if (slideIndex === -1) {
    throw new Error("addTextBox: slide handle was not found in PptxSourceModel source");
  }

  const slide = source.slides[slideIndex];
  const shapeId = nextShapeId(slide.shapes, source.edits ?? [], slide.partPath);
  const shapeIdValue = String(shapeId);
  const name = input.name?.trim() || `TextBox ${shapeIdValue}`;
  const orderingSlot = nextOrderingSlot(slide.shapes);
  const xml = buildTextBoxXml({
    shapeId: shapeIdValue,
    name,
    offsetX: input.offsetX,
    offsetY: input.offsetY,
    width: input.width,
    height: input.height,
    ...(input.rotation !== undefined ? { rotation: input.rotation } : {}),
    ...(input.paragraphs !== undefined ? { paragraphs: input.paragraphs } : { text: input.text }),
    ...(input.body !== undefined ? { body: input.body } : {}),
  });
  const shape = parseShapeNodeXml(xml, slide.partPath, orderingSlot);
  const slides = source.slides.map((candidate, index) =>
    index === slideIndex
      ? {
          ...candidate,
          shapes: [...candidate.shapes, shape],
        }
      : candidate,
  );

  return {
    ...source,
    slides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "addTextBox",
        slidePartPath: slide.partPath,
        shapeId: shapeIdValue,
        xml,
      },
    ],
  };
}

export function addShape(
  source: PptxSourceModel,
  slideHandle: SourceHandle,
  input: AddShapeInput,
): PptxSourceModel {
  assertShapeInput(input);
  const slideIndex = source.slides.findIndex((slide) =>
    sourceHandlesEqual(slide.handle, slideHandle),
  );
  if (slideIndex === -1) {
    throw new Error("addShape: slide handle was not found in PptxSourceModel source");
  }

  const slide = source.slides[slideIndex];
  const shapeId = nextShapeId(slide.shapes, source.edits ?? [], slide.partPath);
  const shapeIdValue = String(shapeId);
  const name = input.name?.trim() || `Shape ${shapeIdValue}`;
  const preset = input.preset.trim();
  const orderingSlot = nextOrderingSlot(slide.shapes);
  const xml = buildShapeXml({
    shapeId: shapeIdValue,
    name,
    preset,
    offsetX: input.offsetX,
    offsetY: input.offsetY,
    width: input.width,
    height: input.height,
    ...(input.rotation !== undefined ? { rotation: input.rotation } : {}),
    ...(input.fill !== undefined ? { fill: input.fill } : {}),
    ...(input.outline !== undefined ? { outline: input.outline } : {}),
    ...(input.effects !== undefined ? { effects: input.effects } : {}),
    ...(input.paragraphs !== undefined ? { paragraphs: input.paragraphs } : {}),
    ...(input.text !== undefined ? { text: input.text } : {}),
    ...(input.body !== undefined ? { body: input.body } : {}),
  });
  const shape = parseShapeNodeXml(xml, slide.partPath, orderingSlot);
  const slides = source.slides.map((candidate, index) =>
    index === slideIndex
      ? {
          ...candidate,
          shapes: [...candidate.shapes, shape],
        }
      : candidate,
  );

  return {
    ...source,
    slides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "addShape",
        slidePartPath: slide.partPath,
        shapeId: shapeIdValue,
        xml,
      } satisfies PptxSourceModelAddShapeEdit,
    ],
  };
}

export function addConnector(
  source: PptxSourceModel,
  slideHandle: SourceHandle,
  input: AddConnectorInput,
): PptxSourceModel {
  assertConnectorInput(input);
  const slideIndex = source.slides.findIndex((slide) =>
    sourceHandlesEqual(slide.handle, slideHandle),
  );
  if (slideIndex === -1) {
    throw new Error("addConnector: slide handle was not found in PptxSourceModel source");
  }

  const slide = source.slides[slideIndex];
  const startShape =
    input.start !== undefined
      ? requireConnectorTargetShape(slide, input.start.shapeHandle, "start")
      : undefined;
  const endShape =
    input.end !== undefined
      ? requireConnectorTargetShape(slide, input.end.shapeHandle, "end")
      : undefined;
  const shapeId = nextShapeId(slide.shapes, source.edits ?? [], slide.partPath);
  const shapeIdValue = String(shapeId);
  const name = input.name?.trim() || `Connector ${shapeIdValue}`;
  const orderingSlot = nextOrderingSlot(slide.shapes);
  const startShapeId = startShape !== undefined ? String(startShape.nodeId) : undefined;
  const endShapeId = endShape !== undefined ? String(endShape.nodeId) : undefined;
  const xml = buildConnectorXml({
    shapeId: shapeIdValue,
    name,
    preset: input.preset,
    offsetX: input.offsetX,
    offsetY: input.offsetY,
    width: input.width,
    height: input.height,
    ...(startShapeId !== undefined && input.start !== undefined
      ? {
          startShapeId,
          startConnectionSiteIndex: input.start.connectionSiteIndex,
        }
      : {}),
    ...(endShapeId !== undefined && input.end !== undefined
      ? {
          endShapeId,
          endConnectionSiteIndex: input.end.connectionSiteIndex,
        }
      : {}),
    ...(input.outline !== undefined ? { outline: input.outline } : {}),
  });
  const connector = parseShapeNodeXml(xml, slide.partPath, orderingSlot);
  const slides = source.slides.map((candidate, index) =>
    index === slideIndex
      ? {
          ...candidate,
          shapes: [...candidate.shapes, connector],
        }
      : candidate,
  );

  return {
    ...source,
    slides,
    edits: [
      ...(source.edits ?? []),
      {
        kind: "addConnector",
        slidePartPath: slide.partPath,
        shapeId: shapeIdValue,
        ...(startShapeId !== undefined ? { startShapeId } : {}),
        ...(endShapeId !== undefined ? { endShapeId } : {}),
        xml,
      } satisfies PptxSourceModelAddConnectorEdit,
    ],
  };
}

export function deleteShape(source: PptxSourceModel, handle: SourceHandle): PptxSourceModel {
  if (handle.nodeId === undefined) {
    throw new Error("deleteShape: shape delete requires a node id");
  }

  let found: SourceShapeNode | undefined;
  let deleted = false;
  const slides = source.slides.map((slide) => {
    let slideChanged = false;
    const nextShapes = slide.shapes.filter((shape) => {
      if (!sourceHandlesEqual(shape.handle, handle)) return true;
      found = shape;
      if (shape.kind !== "shape" && shape.kind !== "connector") {
        throw new Error("deleteShape: only top-level sp or cxnSp shapes can be deleted");
      }
      if (hasAlternateContentSidecar(shape)) {
        throw new Error("deleteShape: shapes inside AlternateContent are not supported");
      }
      deleted = true;
      slideChanged = true;
      return false;
    });
    return slideChanged ? { ...slide, shapes: nextShapes } : slide;
  });

  if (!deleted) {
    if (
      found === undefined &&
      source.slides.some((slide) => hasNestedShapeNodeWithHandle(slide.shapes, handle))
    ) {
      throw new Error("deleteShape: nested group shape deletion is not supported");
    }
    throw new Error("deleteShape: shape handle was not found in PptxSourceModel source");
  }

  const referencingConnector = findConnectorReferencingShape(source, handle);
  if (referencingConnector !== undefined) {
    throw new Error(
      `deleteShape: shape is referenced by connector '${referencingConnector.name ?? referencingConnector.nodeId ?? "unknown"}'`,
    );
  }

  const retainedEdits = (source.edits ?? []).filter((edit) => !editTargetsShape(edit, handle));
  const deletedInsertedShape = (source.edits ?? []).some((edit) => {
    const inserted = editInsertedShape(edit);
    return (
      inserted !== undefined &&
      inserted.slidePartPath === handle.partPath &&
      inserted.shapeId === String(handle.nodeId)
    );
  });

  return {
    ...source,
    slides,
    edits: deletedInsertedShape
      ? retainedEdits
      : [
          ...retainedEdits,
          {
            kind: "deleteShape",
            handle,
          },
        ],
  };
}

function shapeTransformPositionAndSizeEqual(
  current: TransformableShapeNode["transform"],
  next: UpdateShapeTransformInput,
): boolean {
  return (
    current?.offsetX === next.offsetX &&
    current.offsetY === next.offsetY &&
    current.width === next.width &&
    current.height === next.height
  );
}

function patchSourceOutline(
  current: SourceOutline | undefined,
  patch: EditableShapeOutline,
): SourceOutline {
  return {
    ...(current ?? {}),
    ...(patch.width !== undefined ? { width: patch.width } : {}),
    ...(patch.fill !== undefined ? { fill: toSourceFill(patch.fill) } : {}),
  };
}

function appendShapeFillEdit(
  edits: readonly PptxSourceModelEdit[],
  handle: SourceHandle,
  fill: EditableShapeFill,
): PptxSourceModelEdit[] {
  const retainedEdits = edits.filter(
    (edit) => edit.kind !== "updateShapeFill" || !sourceHandlesEqual(edit.handle, handle),
  );
  return [...retainedEdits, { kind: "updateShapeFill", handle, fill }];
}

function appendShapeOutlineEdit(
  edits: readonly PptxSourceModelEdit[],
  handle: SourceHandle,
  outline: EditableShapeOutline,
): PptxSourceModelEdit[] {
  let outlineEdit: PptxSourceModelShapeOutlineEdit = {
    kind: "updateShapeOutline",
    handle,
    outline,
  };
  const retainedEdits: PptxSourceModelEdit[] = [];

  for (const edit of edits) {
    if (edit.kind !== "updateShapeOutline" || !sourceHandlesEqual(edit.handle, handle)) {
      retainedEdits.push(edit);
      continue;
    }
    outlineEdit = {
      ...outlineEdit,
      outline: mergeEditableShapeOutline(edit.outline, outlineEdit.outline),
    };
  }

  return [...retainedEdits, outlineEdit];
}

function mergeEditableShapeOutline(
  base: EditableShapeOutline,
  patch: EditableShapeOutline,
): EditableShapeOutline {
  return {
    ...(base.width !== undefined ? { width: base.width } : {}),
    ...(base.fill !== undefined ? { fill: base.fill } : {}),
    ...(patch.width !== undefined ? { width: patch.width } : {}),
    ...(patch.fill !== undefined ? { fill: patch.fill } : {}),
  };
}

function toSourceFill(fill: EditableShapeFill): SourceFill {
  if (fill.kind === "none") return { kind: "none" };
  return {
    kind: "solid",
    color: { kind: "srgb", hex: fill.color.hex },
  };
}

function assertEditableShapeFill(fill: EditableShapeFill, operationName: string): void {
  if (fill.kind === "none") return;
  if (fill.kind !== "solid") {
    throw new Error(`${operationName}: only solid and none fills are supported`);
  }
  if (fill.color.kind !== "srgb") {
    throw new Error(`${operationName}: only srgb solid fill colors are supported`);
  }
  if (!/^[0-9A-Fa-f]{6}$/.test(fill.color.hex)) {
    throw new Error(`${operationName}: srgb fill color must be a 6-digit hex value`);
  }
}

function assertEditableShapeOutline(outline: EditableShapeOutline, operationName: string): void {
  if (outline.width === undefined && outline.fill === undefined) {
    throw new Error(`${operationName}: outline must set width or fill`);
  }
  if (outline.width !== undefined) {
    assertPositiveFiniteEmu(outline.width, operationName, "width");
  }
  if (outline.fill !== undefined) assertEditableShapeFill(outline.fill, operationName);
}

function sourceFillEqual(left: SourceFill | undefined, right: SourceFill | undefined): boolean {
  return stableValueEqual(left ?? {}, right ?? {});
}

function sourceOutlineEqual(
  left: SourceOutline | undefined,
  right: SourceOutline | undefined,
): boolean {
  return stableValueEqual(left ?? {}, right ?? {});
}

function stableValueEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right)) return false;
    if (left.length !== right.length) return false;
    return left.every((value, index) => stableValueEqual(value, right[index]));
  }
  if (isPlainRecord(left) || isPlainRecord(right)) {
    if (!isPlainRecord(left) || !isPlainRecord(right)) return false;
    const leftKeys = Object.keys(left).sort();
    const rightKeys = Object.keys(right).sort();
    if (!stableValueEqual(leftKeys, rightKeys)) return false;
    return leftKeys.every((key) => stableValueEqual(left[key], right[key]));
  }
  return false;
}

function isPlainRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function hasEditableTransform(shape: SourceShapeNode): shape is TransformableShapeNode & {
  readonly transform: NonNullable<TransformableShapeNode["transform"]>;
} {
  return shape.kind !== "raw" && shape.transform !== undefined;
}

function assertTextBoxInput(input: AddTextBoxInput): void {
  assertFiniteEmu(input.offsetX, "addTextBox", "offsetX");
  assertFiniteEmu(input.offsetY, "addTextBox", "offsetY");
  assertPositiveFiniteEmu(input.width, "addTextBox", "width");
  assertPositiveFiniteEmu(input.height, "addTextBox", "height");
  if (input.rotation !== undefined) {
    assertFiniteIntegerNumber(input.rotation, "addTextBox", "rotation");
  }
  if (input.text !== undefined && input.paragraphs !== undefined) {
    throw new Error("addTextBox: specify either text or paragraphs, not both");
  }
  if (input.text === undefined && input.paragraphs === undefined) {
    throw new Error("addTextBox: text or paragraphs must be provided");
  }
  if (input.text !== undefined && typeof input.text !== "string") {
    throw new Error("addTextBox: text must be a string when provided");
  }
  if (input.paragraphs !== undefined) {
    assertTextBoxParagraphs(input.paragraphs);
  }
  if (input.body !== undefined) {
    assertTextBoxBody(input.body, "body");
  }
  if (input.name !== undefined && input.name.trim() === "") {
    throw new Error("addTextBox: name must be a non-empty string when provided");
  }
}

function assertShapeInput(input: AddShapeInput): void {
  assertFiniteEmu(input.offsetX, "addShape", "offsetX");
  assertFiniteEmu(input.offsetY, "addShape", "offsetY");
  assertPositiveFiniteEmu(input.width, "addShape", "width");
  assertPositiveFiniteEmu(input.height, "addShape", "height");
  if (typeof input.preset !== "string" || input.preset.trim() === "") {
    throw new Error("addShape: preset must be a non-empty string");
  }
  if (input.rotation !== undefined) {
    assertFiniteIntegerNumber(input.rotation, "addShape", "rotation");
  }
  if (input.text !== undefined && input.paragraphs !== undefined) {
    throw new Error("addShape: specify either text or paragraphs, not both");
  }
  if (input.text !== undefined && typeof input.text !== "string") {
    throw new Error("addShape: text must be a string when provided");
  }
  if (input.body !== undefined && input.text === undefined && input.paragraphs === undefined) {
    throw new Error("addShape: body requires text or paragraphs");
  }
  if (input.paragraphs !== undefined) {
    assertTextBoxParagraphs(input.paragraphs, "addShape");
  }
  if (input.body !== undefined) {
    assertTextBoxBody(input.body, "body", "addShape");
  }
  if (input.name !== undefined && input.name.trim() === "") {
    throw new Error("addShape: name must be a non-empty string when provided");
  }
  if (input.fill !== undefined) assertShapeFill(input.fill, "fill");
  if (input.outline !== undefined) assertShapeOutline(input.outline);
  if (input.effects !== undefined) assertShapeEffects(input.effects);
}

function assertTextBoxParagraphs(
  paragraphs: readonly TextBoxParagraphInput[],
  operationName = "addTextBox",
): void {
  if (!isArrayValue(paragraphs) || paragraphs.length === 0) {
    throw new Error(`${operationName}: paragraphs must contain at least one paragraph`);
  }
  paragraphs.forEach((paragraph: unknown, paragraphIndex: number) => {
    if (!isPlainRecord(paragraph)) {
      throw new Error(`${operationName}: paragraphs[${paragraphIndex}] must be an object`);
    }
    const runs = paragraph.runs;
    if (!isArrayValue(runs) || runs.length === 0) {
      throw new Error(
        `${operationName}: paragraphs[${paragraphIndex}].runs must contain at least one run`,
      );
    }
    assertTextBoxParagraphProperties(paragraph.properties, paragraphIndex, operationName);
    runs.forEach((run, runIndex) => assertTextBoxRun(run, paragraphIndex, runIndex, operationName));
  });
}

function assertTextBoxParagraphProperties(
  properties: unknown,
  paragraphIndex: number,
  operationName = "addTextBox",
): void {
  if (properties === undefined) return;
  if (!isPlainRecord(properties)) {
    throw new Error(`${operationName}: paragraphs[${paragraphIndex}].properties must be an object`);
  }
  const align = properties.align;
  const lineSpacing = properties.lineSpacing;
  if (
    align !== undefined &&
    align !== "left" &&
    align !== "center" &&
    align !== "right" &&
    align !== "justify"
  ) {
    throw new Error(
      `${operationName}: paragraphs[${paragraphIndex}].properties.align is not supported`,
    );
  }
  if (lineSpacing !== undefined) {
    assertPositiveFiniteIntegerNumber(
      lineSpacing,
      operationName,
      `paragraphs[${paragraphIndex}].properties.lineSpacing`,
    );
  }
}

function assertTextBoxRun(
  run: unknown,
  paragraphIndex: number,
  runIndex: number,
  operationName = "addTextBox",
): void {
  const path = `paragraphs[${paragraphIndex}].runs[${runIndex}]`;
  if (!isPlainRecord(run)) {
    throw new Error(`${operationName}: ${path} must be an object`);
  }
  if (typeof run.text !== "string") {
    throw new Error(`${operationName}: ${path}.text must be a string`);
  }
  assertTextBoxRunProperties(run.properties, path, operationName);
}

function assertTextBoxRunProperties(
  properties: unknown,
  path: string,
  operationName = "addTextBox",
): void {
  if (properties === undefined) return;
  if (!isPlainRecord(properties)) {
    throw new Error(`${operationName}: ${path}.properties must be an object`);
  }
  const fontFace = properties.fontFace;
  const fontSize = properties.fontSize;
  const color = properties.color;
  const gradientFill = properties.gradientFill;
  const underline = properties.underline;
  const baseline = properties.baseline;
  const highlight = properties.highlight;
  const glow = properties.glow;
  const outline = properties.outline;
  const charSpacing = properties.charSpacing;
  if (fontFace !== undefined && (typeof fontFace !== "string" || fontFace.trim() === "")) {
    throw new Error(`${operationName}: ${path}.properties.fontFace must be a non-empty string`);
  }
  if (fontSize !== undefined) {
    assertPositiveFiniteNumber(fontSize, operationName, `${path}.properties.fontSize`);
  }
  assertBooleanOrUndefined(properties.bold, operationName, `${path}.properties.bold`);
  assertBooleanOrUndefined(properties.italic, operationName, `${path}.properties.italic`);
  assertBooleanOrUndefined(properties.strike, operationName, `${path}.properties.strike`);
  if (color !== undefined) assertTextBoxColor(color, `${path}.properties.color`, operationName);
  if (gradientFill !== undefined) {
    if (color !== undefined) {
      throw new Error(
        `${operationName}: ${path}.properties cannot set both color and gradientFill`,
      );
    }
    assertTextBoxGradientFill(gradientFill, `${path}.properties.gradientFill`, operationName);
  }
  if (underline !== undefined) {
    assertTextBoxUnderline(underline, `${path}.properties.underline`, operationName);
  }
  if (baseline !== undefined) {
    assertTextBoxBaseline(baseline, `${path}.properties.baseline`, operationName);
  }
  if (highlight !== undefined) {
    assertTextBoxColor(highlight, `${path}.properties.highlight`, operationName);
  }
  if (glow !== undefined) assertTextBoxGlow(glow, `${path}.properties.glow`, operationName);
  if (outline !== undefined) {
    assertTextBoxOutline(outline, `${path}.properties.outline`, operationName);
  }
  if (charSpacing !== undefined) {
    assertTextPointNumber(charSpacing, operationName, `${path}.properties.charSpacing`);
  }
}

function assertTextBoxColor(color: unknown, path: string, operationName = "addTextBox"): void {
  if (!isPlainRecord(color)) {
    throw new Error(`${operationName}: ${path} must be an srgb 6-digit hex color`);
  }
  if (
    color.kind !== "srgb" ||
    typeof color.hex !== "string" ||
    !/^[0-9A-Fa-f]{6}$/.test(color.hex)
  ) {
    throw new Error(`${operationName}: ${path} must be an srgb 6-digit hex color`);
  }
}

function assertTextBoxGradientFill(
  fill: unknown,
  path: string,
  operationName = "addTextBox",
): void {
  if (!isPlainRecord(fill)) {
    throw new Error(`${operationName}: ${path} must be a gradient fill object`);
  }
  const stops = fill.stops;
  if (!isArrayValue(stops) || stops.length < 2) {
    throw new Error(`${operationName}: ${path}.stops must contain at least two stops`);
  }
  if (fill.angle !== undefined)
    assertFiniteIntegerNumber(fill.angle, operationName, `${path}.angle`);
  stops.forEach((stop: unknown, index: number) => {
    if (!isPlainRecord(stop)) {
      throw new Error(`${operationName}: ${path}.stops[${index}] must be an object`);
    }
    const position = stop.position;
    assertFiniteIntegerNumber(position, operationName, `${path}.stops[${index}].position`);
    if (typeof position !== "number" || position < 0 || position > 100000) {
      throw new Error(
        `${operationName}: ${path}.stops[${index}].position must be between 0 and 100000`,
      );
    }
    assertTextBoxColor(stop.color, `${path}.stops[${index}].color`, operationName);
  });
}

function isArrayValue(value: unknown): value is readonly unknown[] {
  return Array.isArray(value);
}

function assertTextBoxUnderline(
  underline: unknown,
  path: string,
  operationName = "addTextBox",
): void {
  if (typeof underline === "boolean") return;
  if (!isPlainRecord(underline)) {
    throw new Error(`${operationName}: ${path} must be a boolean or underline options object`);
  }
  const style = underline.style ?? "sng";
  if (typeof style !== "string" || !UNDERLINE_STYLES.has(style)) {
    throw new Error(`${operationName}: ${path}.style is not supported`);
  }
  if (underline.color !== undefined) {
    assertTextBoxColor(underline.color, `${path}.color`, operationName);
  }
}

function assertTextBoxBaseline(
  baseline: unknown,
  path: string,
  operationName = "addTextBox",
): void {
  if (baseline === "subscript" || baseline === "superscript") return;
  throw new Error(`${operationName}: ${path} must be subscript or superscript`);
}

function assertTextBoxGlow(glow: unknown, path: string, operationName = "addTextBox"): void {
  if (!isPlainRecord(glow)) {
    throw new Error(`${operationName}: ${path} must be an object`);
  }
  assertPositiveFiniteEmu(glow.radius, operationName, `${path}.radius`);
  assertTextBoxColor(glow.color, `${path}.color`, operationName);
}

function assertTextBoxOutline(outline: unknown, path: string, operationName = "addTextBox"): void {
  if (!isPlainRecord(outline)) {
    throw new Error(`${operationName}: ${path} must be an object`);
  }
  if (outline.width === undefined && outline.color === undefined) {
    throw new Error(`${operationName}: ${path} must set width or color`);
  }
  if (outline.width !== undefined) {
    assertPositiveFiniteEmu(outline.width, operationName, `${path}.width`);
  }
  if (outline.color !== undefined)
    assertTextBoxColor(outline.color, `${path}.color`, operationName);
}

function assertTextBoxBody(body: unknown, path: string, operationName = "addTextBox"): void {
  if (!isPlainRecord(body)) {
    throw new Error(`${operationName}: ${path} must be an object`);
  }
  if (
    body.anchor !== undefined &&
    body.anchor !== "top" &&
    body.anchor !== "middle" &&
    body.anchor !== "bottom"
  ) {
    throw new Error(`${operationName}: body.anchor is not supported`);
  }
  if (body.marginLeft !== undefined)
    assertFiniteEmu(body.marginLeft, operationName, "body.marginLeft");
  if (body.marginRight !== undefined) {
    assertFiniteEmu(body.marginRight, operationName, "body.marginRight");
  }
  if (body.marginTop !== undefined)
    assertFiniteEmu(body.marginTop, operationName, "body.marginTop");
  if (body.marginBottom !== undefined) {
    assertFiniteEmu(body.marginBottom, operationName, "body.marginBottom");
  }
}

function assertConnectorInput(input: AddConnectorInput): void {
  assertFiniteEmu(input.offsetX, "addConnector", "offsetX");
  assertFiniteEmu(input.offsetY, "addConnector", "offsetY");
  assertPositiveFiniteEmu(input.width, "addConnector", "width");
  assertPositiveFiniteEmu(input.height, "addConnector", "height");
  if (!CONNECTOR_PRESETS.has(input.preset)) {
    throw new Error(
      "addConnector: preset must be straightConnector1, bentConnector3, or curvedConnector3",
    );
  }
  if (input.start !== undefined)
    assertConnectionSiteIndex(input.start.connectionSiteIndex, "start");
  if (input.end !== undefined) assertConnectionSiteIndex(input.end.connectionSiteIndex, "end");
  if (input.name !== undefined && input.name.trim() === "") {
    throw new Error("addConnector: name must be a non-empty string when provided");
  }
  assertArrowEndpoint(input.outline?.headEnd, "headEnd", "addConnector");
  assertArrowEndpoint(input.outline?.tailEnd, "tailEnd", "addConnector");
}

function assertShapeFill(fill: unknown, path: string): void {
  if (!isPlainRecord(fill)) {
    throw new Error(`addShape: ${path} must be a fill object`);
  }
  switch (fill.kind) {
    case "none":
      return;
    case "solid":
      assertTextBoxColor(fill.color, `${path}.color`, "addShape");
      return;
    case "gradient":
      assertTextBoxGradientFill(fill, path, "addShape");
      return;
    default:
      throw new Error(`addShape: ${path}.kind is not supported`);
  }
}

function assertShapeOutline(outline: unknown): void {
  if (!isPlainRecord(outline)) {
    throw new Error("addShape: outline must be an object");
  }
  if (
    outline.width === undefined &&
    outline.fill === undefined &&
    outline.dash === undefined &&
    outline.headEnd === undefined &&
    outline.tailEnd === undefined
  ) {
    throw new Error("addShape: outline must set width, fill, dash, headEnd, or tailEnd");
  }
  if (outline.width !== undefined) {
    assertPositiveFiniteEmu(outline.width, "addShape", "outline.width");
  }
  if (outline.fill !== undefined) assertShapeFill(outline.fill, "outline.fill");
  if (outline.dash !== undefined) {
    if (typeof outline.dash !== "string" || !DASH_STYLES.has(outline.dash)) {
      throw new Error("addShape: outline.dash is not supported");
    }
  }
  assertArrowEndpoint(outline.headEnd, "headEnd", "addShape");
  assertArrowEndpoint(outline.tailEnd, "tailEnd", "addShape");
}

function assertShapeEffects(effects: unknown): void {
  if (!isPlainRecord(effects)) {
    throw new Error("addShape: effects must be an object");
  }
  if (effects.glow === undefined) {
    throw new Error("addShape: effects must set glow");
  }
  if (effects.glow !== undefined) assertTextBoxGlow(effects.glow, "effects.glow", "addShape");
}

function assertConnectionSiteIndex(value: number, fieldName: "start" | "end"): void {
  if (!Number.isInteger(value) || value < 0) {
    throw new Error(
      `addConnector: ${fieldName}.connectionSiteIndex must be a non-negative integer`,
    );
  }
}

function assertArrowEndpoint(
  endpoint: unknown,
  fieldName: string,
  operationName: "addConnector" | "addShape",
): void {
  if (endpoint === undefined) return;
  if (!isPlainRecord(endpoint)) {
    throw new Error(`${operationName}: outline.${fieldName} must be an object`);
  }
  if (typeof endpoint.type !== "string" || !ARROW_TYPES.has(endpoint.type)) {
    throw new Error(`${operationName}: outline.${fieldName}.type is not supported`);
  }
  if (typeof endpoint.width !== "string" || !ARROW_SIZES.has(endpoint.width)) {
    throw new Error(`${operationName}: outline.${fieldName}.width is not supported`);
  }
  if (typeof endpoint.length !== "string" || !ARROW_SIZES.has(endpoint.length)) {
    throw new Error(`${operationName}: outline.${fieldName}.length is not supported`);
  }
}

function assertBooleanOrUndefined(value: unknown, operationName: string, fieldName: string): void {
  if (value !== undefined && typeof value !== "boolean") {
    throw new Error(`${operationName}: ${fieldName} must be a boolean value`);
  }
}

function assertFiniteIntegerNumber(value: unknown, operationName: string, fieldName: string): void {
  if (!Number.isInteger(value)) {
    throw new Error(`${operationName}: ${fieldName} must be a finite integer`);
  }
}

function assertPositiveFiniteNumber(
  value: unknown,
  operationName: string,
  fieldName: string,
): void {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw new Error(`${operationName}: ${fieldName} must be a finite positive number`);
  }
}

function assertPositiveFiniteIntegerNumber(
  value: unknown,
  operationName: string,
  fieldName: string,
): void {
  if (typeof value !== "number" || !Number.isInteger(value) || value <= 0) {
    throw new Error(`${operationName}: ${fieldName} must be a finite positive integer`);
  }
}

function assertTextPointNumber(value: unknown, operationName: string, fieldName: string): void {
  assertFiniteIntegerNumber(value, operationName, fieldName);
  if (typeof value !== "number" || value < -400000 || value > 400000) {
    throw new Error(`${operationName}: ${fieldName} must be between -400000 and 400000`);
  }
}

function assertFiniteEmu(value: unknown, operationName: string, fieldName: string): void {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${operationName}: ${fieldName} must be a finite EMU value`);
  }
}

function assertPositiveFiniteEmu(value: unknown, operationName: string, fieldName: string): void {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw new Error(`${operationName}: ${fieldName} must be a finite positive EMU value`);
  }
}

function requireConnectorTargetShape(
  slide: SourceSlide,
  handle: SourceHandle,
  endpointName: "start" | "end",
): SourceShape & { readonly nodeId: SourceNodeId } {
  const shape = slide.shapes.find((candidate) => sourceHandlesEqual(candidate.handle, handle));
  if (shape === undefined) {
    throw new Error(`addConnector: ${endpointName} shape handle was not found on the target slide`);
  }
  if (shape.kind !== "shape") {
    throw new Error(`addConnector: ${endpointName} target must be a top-level sp shape`);
  }
  const nodeId = shape.nodeId;
  if (nodeId === undefined) {
    throw new Error(`addConnector: ${endpointName} target shape requires a node id`);
  }
  return { ...shape, nodeId };
}

function nextShapeId(
  shapes: readonly SourceShapeNode[],
  edits: readonly PptxSourceModelEdit[],
  slidePartPath: PartPath,
): SourceNodeId {
  const used = new Set<number>();
  collectNumericShapeIds(shapes, used);
  collectNumericShapeEditIds(edits, slidePartPath, used);
  const usedNames = new Set([...used].map(String));
  return asSourceNodeId(nextNumberedName(usedNames, /^(\d+)$/, String));
}

function collectNumericShapeIds(shapes: readonly SourceShapeNode[], used: Set<number>): void {
  for (const shape of shapes) {
    const numericId = Number(shape.nodeId);
    if (Number.isInteger(numericId) && numericId > 0) used.add(numericId);
    if (shape.kind === "group") collectNumericShapeIds(shape.children, used);
  }
}

function collectNumericShapeEditIds(
  edits: readonly PptxSourceModelEdit[],
  slidePartPath: PartPath,
  used: Set<number>,
): void {
  for (const edit of edits) {
    const numericId = Number(editReservedShapeId(edit, slidePartPath));
    if (Number.isInteger(numericId) && numericId > 0) used.add(numericId);
  }
}

function nextOrderingSlot(shapes: readonly SourceShapeNode[]): number {
  return (
    shapes.reduce((current, shape) => {
      const slot = shape.handle?.orderingSlot ?? -1;
      return Math.max(current, slot);
    }, -1) + 1
  );
}

function findShapeNodeInTree(
  shapes: readonly SourceShapeNode[],
  handle: SourceHandle,
): SourceShapeNode | undefined {
  for (const shape of shapes) {
    if (sourceHandlesEqual(shape.handle, handle)) return shape;
    if (shape.kind === "group") {
      const child = findShapeNodeInTree(shape.children, handle);
      if (child !== undefined) return child;
    }
  }
  return undefined;
}

function hasNestedShapeNodeWithHandle(
  shapes: readonly SourceShapeNode[],
  handle: SourceHandle,
): boolean {
  return shapes.some(
    (shape) =>
      shape.kind === "group" &&
      (findShapeNodeInTree(shape.children, handle) !== undefined ||
        hasNestedShapeNodeWithHandle(shape.children, handle)),
  );
}

function hasAlternateContentSidecar(shape: SourceShapeNode): boolean {
  if (shape.kind === "raw") return false;
  return shape.rawSidecars?.some((sidecar) => sidecar.node.name === "mc:AlternateContent") ?? false;
}

function findConnectorReferencingShape(
  source: PptxSourceModel,
  handle: SourceHandle,
): SourceConnector | undefined {
  if (handle.nodeId === undefined) return undefined;
  for (const slide of source.slides) {
    if (slide.partPath !== handle.partPath) continue;
    for (const shape of slide.shapes) {
      if (shape.kind !== "connector") continue;
      if (
        shape.connection?.start?.shapeId === handle.nodeId ||
        shape.connection?.end?.shapeId === handle.nodeId
      ) {
        return shape;
      }
    }
  }
  return undefined;
}
