/**
 * Shape editing and authoring operations for PptxSourceModel.
 *
 * Shape authoring uses the same edit-time XML finalization model as text boxes and
 * connectors. Preset names remain pass-through strings because OOXML preset coverage
 * is broad, while preset adjustments and the supported custom path commands stay
 * inside the typed subset that the source reader can parse back into `SourceShape`.
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
 *
 * Paragraph spacing and explicit baselines use discriminated percentage inputs so
 * callers cannot accidentally interchange point and percentage units. Bullet options
 * form a discriminated union because character, automatic numbering, and explicit
 * absence serialize to different DrawingML elements. Bullet font and size stay nested
 * with the selected bullet kind while the parsed source model keeps the corresponding
 * sibling OOXML properties. Shape auto-fit is exposed as a semantic `"shape"` option
 * and is read back as the source model's `spAutofit` value.
 */

import { nextDrawingOrderingSlot, nextDrawingShapeId } from "./drawing-authoring-allocation.js";
import { editInsertedShape, editTargetsShape, sourceHandlesEqual } from "./edit-descriptors.js";
import {
  assertShadowEffectsInput,
  type InnerShadowInput,
  type OuterShadowInput,
} from "./effect-authoring.js";
import type { RelationshipId } from "./handles.js";
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
  Relationship,
  SourceArrowEndpoint,
  SourceConnector,
  SourceFill,
  SourceHandle,
  SourceNodeId,
  SourceOutline,
  SourceShape,
  SourceShapeNode,
} from "./index.js";
import { nextRelationshipId } from "./package-graph-mutations.js";
import type {
  AuthoringColorTransformInput,
  ShapeColorInput,
  ShapeCustomGeometryInput,
  ShapeCustomGeometryPathCommandInput,
  ShapeCustomGeometryPathInput,
  ShapeEffectsInput,
  ShapeFillInput,
  ShapeGeometryInput,
  ShapeGlowInput,
  ShapeGradientFillInput,
  ShapeOutlineInput,
  ShapePresetGeometryInput,
  TextBoxBaselineInput,
  TextBoxBodyPropertiesInput,
  TextBoxBulletInput,
  TextBoxColorInput,
  TextBoxGlowInput,
  TextBoxGradientFillInput,
  TextBoxGradientStopInput,
  TextBoxLineSpacingInput,
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
  buildSlideNumberXml,
  buildTextBoxXml,
  parseShapeNodeXml,
} from "./shape-xml.js";

export type AddTextBoxBaselineInput = TextBoxBaselineInput;
export type AddTextBoxBulletInput = TextBoxBulletInput;
export type AddTextBoxLineSpacingInput = TextBoxLineSpacingInput;
export type AddTextBoxBodyPropertiesInput = TextBoxBodyPropertiesInput;
export type AddTextBoxColorInput = TextBoxColorInput;
export type AddTextBoxColorTransformInput = AuthoringColorTransformInput;
export type AddTextBoxGlowInput = TextBoxGlowInput;
export type AddTextBoxGradientFillInput = TextBoxGradientFillInput;
export type AddTextBoxGradientStopInput = TextBoxGradientStopInput;
export type AddTextBoxOutlineInput = TextBoxOutlineInput;
export type AddTextBoxParagraphInput = TextBoxParagraphInput;
export type AddTextBoxParagraphPropertiesInput = TextBoxParagraphPropertiesInput;
export type AddTextBoxRunInput = TextBoxRunInput;
export type AddTextBoxRunPropertiesInput = TextBoxRunPropertiesInput;
export type AddTextBoxUnderlineInput = TextBoxUnderlineInput;
export type AddTextBoxUnderlineStyle = TextBoxUnderlineStyle;
export type AddShapeBodyPropertiesInput = TextBoxBodyPropertiesInput;
export type AddShapeColorInput = ShapeColorInput;
export type AddShapeColorTransformInput = AuthoringColorTransformInput;
export type AddShapeEffectsInput = ShapeEffectsInput;
export type AddInnerShadowInput = InnerShadowInput;
export type AddOuterShadowInput = OuterShadowInput;
export type AddShapeFillInput = ShapeFillInput;
export type AddShapeGlowInput = ShapeGlowInput;
export type AddShapeGradientFillInput = ShapeGradientFillInput;
export type AddShapeGradientStopInput = TextBoxGradientStopInput;
export type AddShapeGeometryInput = ShapeGeometryInput;
export type AddShapePresetGeometryInput = ShapePresetGeometryInput;
export type AddShapeCustomGeometryInput = ShapeCustomGeometryInput;
export type AddShapeCustomGeometryPathInput = ShapeCustomGeometryPathInput;
export type AddShapeCustomGeometryPathCommandInput = ShapeCustomGeometryPathCommandInput;
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
  "lgDashDotDot",
  "sysDash",
  "sysDot",
]);
const UNDERLINE_STYLES: ReadonlySet<string> = new Set<AddTextBoxUnderlineStyle>([
  "sng",
  "words",
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
const AUTO_NUMBER_SCHEMES = new Set([
  "arabicPeriod",
  "arabicParenR",
  "romanUcPeriod",
  "romanLcPeriod",
  "alphaUcPeriod",
  "alphaLcPeriod",
  "alphaLcParenR",
  "alphaUcParenR",
  "arabicPlain",
]);
const HYPERLINK_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

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
  readonly geometry: AddShapeGeometryInput;
  readonly fill?: AddShapeFillInput;
  readonly outline?: AddShapeOutlineInput;
  readonly effects?: AddShapeEffectsInput;
  readonly text?: string;
  readonly paragraphs?: readonly AddShapeParagraphInput[];
  readonly body?: AddShapeBodyPropertiesInput;
  readonly rotation?: NonNullable<SourceShape["transform"]>["rotation"];
  readonly flipHorizontal?: boolean;
  readonly flipVertical?: boolean;
  readonly name?: string;
}

export interface AddSlideNumberInput extends UpdateShapeTransformInput {
  readonly properties?: AddTextBoxRunPropertiesInput;
  readonly body?: AddTextBoxBodyPropertiesInput;
  readonly align?: "left" | "center" | "right";
  readonly name?: string;
}

export interface AddConnectorConnectionEndpointInput {
  readonly shapeHandle: SourceHandle;
  /** Unsigned index into the target shape's OOXML connection-site table. */
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

interface AuthoringTarget {
  readonly kind: "slide" | "layout" | "master";
  readonly index: number;
  readonly partPath: PartPath;
  readonly shapes: readonly SourceShapeNode[];
}

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
  const target = findAuthoringTarget(source, slideHandle);
  if (target === undefined) {
    throw new Error(
      "addTextBox: slide, layout, or master handle was not found in PptxSourceModel source",
    );
  }
  const body = mergeTextBodyProperties(defaultTextBodyProperties(source, target), input.body);
  const shapeId = nextDrawingShapeId(source, target.shapes, target.partPath);
  const shapeIdValue = String(shapeId);
  const name = input.name?.trim() || `TextBox ${shapeIdValue}`;
  const orderingSlot = nextDrawingOrderingSlot(target.shapes);
  const hyperlinkAllocation = allocateRunHyperlinks(
    source,
    target.partPath,
    input.paragraphs ?? [],
  );
  const xml = buildTextBoxXml({
    shapeId: shapeIdValue,
    name,
    offsetX: input.offsetX,
    offsetY: input.offsetY,
    width: input.width,
    height: input.height,
    ...(input.rotation !== undefined ? { rotation: input.rotation } : {}),
    ...(input.paragraphs !== undefined ? { paragraphs: input.paragraphs } : { text: input.text }),
    ...(body !== undefined ? { body } : {}),
    hyperlinkIds: hyperlinkAllocation.ids,
  });
  const shape = parseShapeNodeXml(xml, target.partPath, orderingSlot);

  return {
    ...withAuthoringTargetShapes(
      { ...source, packageGraph: hyperlinkAllocation.packageGraph },
      target,
      [...target.shapes, shape],
    ),
    edits: [
      ...(source.edits ?? []),
      {
        kind: "addTextBox",
        slidePartPath: target.partPath,
        shapeId: shapeIdValue,
        xml,
      },
    ],
  };
}

export function addSlideNumber(
  source: PptxSourceModel,
  targetHandle: SourceHandle,
  input: AddSlideNumberInput,
): PptxSourceModel {
  assertSlideNumberInput(input);
  const target = findAuthoringTarget(source, targetHandle);
  if (target === undefined || target.kind === "slide") {
    throw new Error(
      "addSlideNumber: layout or master handle was not found in PptxSourceModel source",
    );
  }
  const shapeId = nextDrawingShapeId(source, target.shapes, target.partPath);
  const shapeIdValue = String(shapeId);
  const name = input.name?.trim() || `Slide Number ${shapeIdValue}`;
  const orderingSlot = nextDrawingOrderingSlot(target.shapes);
  const xml = buildSlideNumberXml({
    partPath: target.partPath,
    shapeId: shapeIdValue,
    name,
    offsetX: input.offsetX,
    offsetY: input.offsetY,
    width: input.width,
    height: input.height,
    ...(input.properties !== undefined ? { properties: input.properties } : {}),
    ...(input.body !== undefined ? { body: input.body } : {}),
    ...(input.align !== undefined ? { align: input.align } : {}),
  });
  const shape = parseShapeNodeXml(xml, target.partPath, orderingSlot);
  return {
    ...withAuthoringTargetShapes(source, target, [...target.shapes, shape]),
    edits: [
      ...(source.edits ?? []),
      { kind: "addTextBox", slidePartPath: target.partPath, shapeId: shapeIdValue, xml },
    ],
  };
}

export function addShape(
  source: PptxSourceModel,
  slideHandle: SourceHandle,
  input: AddShapeInput,
): PptxSourceModel {
  assertShapeInput(input);
  const target = findAuthoringTarget(source, slideHandle);
  if (target === undefined) {
    throw new Error(
      "addShape: slide, layout, or master handle was not found in PptxSourceModel source",
    );
  }
  const body = mergeTextBodyProperties(defaultTextBodyProperties(source, target), input.body);
  const shapeId = nextDrawingShapeId(source, target.shapes, target.partPath);
  const shapeIdValue = String(shapeId);
  const name = input.name?.trim() || `Shape ${shapeIdValue}`;
  const geometry =
    input.geometry.kind === "preset"
      ? { ...input.geometry, preset: input.geometry.preset.trim() }
      : input.geometry;
  const orderingSlot = nextDrawingOrderingSlot(target.shapes);
  const hyperlinkAllocation = allocateRunHyperlinks(
    source,
    target.partPath,
    input.paragraphs ?? [],
  );
  const xml = buildShapeXml({
    shapeId: shapeIdValue,
    name,
    geometry,
    offsetX: input.offsetX,
    offsetY: input.offsetY,
    width: input.width,
    height: input.height,
    ...(input.rotation !== undefined ? { rotation: input.rotation } : {}),
    ...(input.flipHorizontal !== undefined ? { flipHorizontal: input.flipHorizontal } : {}),
    ...(input.flipVertical !== undefined ? { flipVertical: input.flipVertical } : {}),
    ...(input.fill !== undefined ? { fill: input.fill } : {}),
    ...(input.outline !== undefined ? { outline: input.outline } : {}),
    ...(input.effects !== undefined ? { effects: input.effects } : {}),
    ...(input.paragraphs !== undefined ? { paragraphs: input.paragraphs } : {}),
    ...(input.text !== undefined ? { text: input.text } : {}),
    ...(body !== undefined ? { body } : {}),
    hyperlinkIds: hyperlinkAllocation.ids,
  });
  const shape = parseShapeNodeXml(xml, target.partPath, orderingSlot);

  return {
    ...withAuthoringTargetShapes(
      { ...source, packageGraph: hyperlinkAllocation.packageGraph },
      target,
      [...target.shapes, shape],
    ),
    edits: [
      ...(source.edits ?? []),
      {
        kind: "addShape",
        slidePartPath: target.partPath,
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
  const target = findAuthoringTarget(source, slideHandle);
  if (target === undefined) {
    throw new Error(
      "addConnector: slide, layout, or master handle was not found in PptxSourceModel source",
    );
  }
  const startShape =
    input.start !== undefined
      ? requireConnectorTargetShape(target, input.start.shapeHandle, "start")
      : undefined;
  const endShape =
    input.end !== undefined
      ? requireConnectorTargetShape(target, input.end.shapeHandle, "end")
      : undefined;
  const shapeId = nextDrawingShapeId(source, target.shapes, target.partPath);
  const shapeIdValue = String(shapeId);
  const name = input.name?.trim() || `Connector ${shapeIdValue}`;
  const orderingSlot = nextDrawingOrderingSlot(target.shapes);
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
  const connector = parseShapeNodeXml(xml, target.partPath, orderingSlot);

  return {
    ...withAuthoringTargetShapes(source, target, [...target.shapes, connector]),
    edits: [
      ...(source.edits ?? []),
      {
        kind: "addConnector",
        slidePartPath: target.partPath,
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

function assertSlideNumberInput(input: AddSlideNumberInput): void {
  assertFiniteEmu(input.offsetX, "addSlideNumber", "offsetX");
  assertFiniteEmu(input.offsetY, "addSlideNumber", "offsetY");
  assertPositiveFiniteEmu(input.width, "addSlideNumber", "width");
  assertPositiveFiniteEmu(input.height, "addSlideNumber", "height");
  if (input.properties !== undefined) {
    assertTextBoxRunProperties(input.properties, "properties", "addSlideNumber");
  }
  if (input.body !== undefined) assertTextBoxBody(input.body, "body", "addSlideNumber");
  if (
    input.align !== undefined &&
    input.align !== "left" &&
    input.align !== "center" &&
    input.align !== "right"
  ) {
    throw new Error("addSlideNumber: align is not supported");
  }
  if (input.name !== undefined && input.name.trim() === "") {
    throw new Error("addSlideNumber: name must be a non-empty string when provided");
  }
}

function assertShapeInput(input: AddShapeInput): void {
  assertFiniteEmu(input.offsetX, "addShape", "offsetX");
  assertFiniteEmu(input.offsetY, "addShape", "offsetY");
  assertShapeGeometry(input.geometry);
  const isLine = input.geometry.kind === "preset" && input.geometry.preset.trim() === "line";
  if (isLine) {
    assertNonNegativeFiniteEmu(input.width, "addShape", "width");
    assertNonNegativeFiniteEmu(input.height, "addShape", "height");
    if (input.width === 0 && input.height === 0) {
      throw new Error("addShape: line width and height must not both be zero");
    }
  } else {
    assertPositiveFiniteEmu(input.width, "addShape", "width");
    assertPositiveFiniteEmu(input.height, "addShape", "height");
  }
  if (input.rotation !== undefined) {
    assertFiniteIntegerNumber(input.rotation, "addShape", "rotation");
  }
  assertBooleanOrUndefined(input.flipHorizontal, "addShape", "flipHorizontal");
  assertBooleanOrUndefined(input.flipVertical, "addShape", "flipVertical");
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

function assertShapeGeometry(geometry: unknown): asserts geometry is AddShapeGeometryInput {
  if (!isPlainRecord(geometry)) throw new Error("addShape: geometry must be an object");
  if (geometry.kind === "preset") {
    if (typeof geometry.preset !== "string" || geometry.preset.trim() === "") {
      throw new Error("addShape: geometry.preset must be a non-empty string");
    }
    if (geometry.adjustValues !== undefined) {
      if (!isPlainRecord(geometry.adjustValues)) {
        throw new Error("addShape: geometry.adjustValues must be an object");
      }
      for (const [name, value] of Object.entries(geometry.adjustValues)) {
        if (name.trim() === "" || name !== name.trim()) {
          throw new Error(
            "addShape: geometry.adjustValues names must be non-empty and have no surrounding whitespace",
          );
        }
        assertFiniteIntegerNumber(value, "addShape", `geometry.adjustValues.${name}`);
      }
    }
    return;
  }
  if (geometry.kind !== "custom") throw new Error("addShape: geometry.kind is not supported");
  if (!isArrayValue(geometry.paths) || geometry.paths.length === 0) {
    throw new Error("addShape: geometry.paths must contain at least one path");
  }
  geometry.paths.forEach((path, pathIndex) => assertCustomGeometryPath(path, pathIndex));
}

function assertCustomGeometryPath(path: unknown, pathIndex: number): void {
  const field = `geometry.paths[${pathIndex}]`;
  if (!isPlainRecord(path)) throw new Error(`addShape: ${field} must be an object`);
  assertPositiveFiniteIntegerNumber(path.width, "addShape", `${field}.width`);
  assertPositiveFiniteIntegerNumber(path.height, "addShape", `${field}.height`);
  if (!isArrayValue(path.commands) || path.commands.length === 0) {
    throw new Error(`addShape: ${field}.commands must contain at least one command`);
  }
  const commands = path.commands;
  commands.forEach((command, commandIndex) => {
    const commandField = `${field}.commands[${commandIndex}]`;
    if (!isPlainRecord(command)) throw new Error(`addShape: ${commandField} must be an object`);
    if (commandIndex === 0 ? command.kind !== "moveTo" : command.kind === "moveTo") {
      throw new Error(`addShape: ${field}.commands must start with one moveTo command`);
    }
    if (command.kind !== "moveTo" && command.kind !== "lineTo" && command.kind !== "close") {
      throw new Error(`addShape: ${commandField}.kind is not supported`);
    }
    if (command.kind === "close") {
      if (commandIndex !== commands.length - 1) {
        throw new Error(`addShape: ${field}.close must be the final command`);
      }
    } else {
      assertFiniteIntegerNumber(command.x, "addShape", `${commandField}.x`);
      assertFiniteIntegerNumber(command.y, "addShape", `${commandField}.y`);
    }
  });
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
  const path = `paragraphs[${paragraphIndex}].properties`;
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
    assertTextBoxLineSpacing(lineSpacing, `${path}.lineSpacing`, operationName);
  }
  if (properties.marginLeft !== undefined) {
    assertFiniteEmu(properties.marginLeft, operationName, `${path}.marginLeft`);
  }
  if (properties.indent !== undefined) {
    assertFiniteEmu(properties.indent, operationName, `${path}.indent`);
  }
  if (properties.bullet !== undefined) {
    assertTextBoxBullet(properties.bullet, `${path}.bullet`, operationName);
  }
}

function assertTextBoxLineSpacing(spacing: unknown, path: string, operationName: string): void {
  if (typeof spacing === "number") {
    assertTextBoxLineSpacingValue(spacing, "points", path, operationName);
    return;
  }
  if (!isPlainRecord(spacing) || (spacing.type !== "points" && spacing.type !== "percent")) {
    throw new Error(`${operationName}: ${path} must be a points or percent spacing object`);
  }
  assertTextBoxLineSpacingValue(spacing.value, spacing.type, `${path}.value`, operationName);
}

function assertTextBoxLineSpacingValue(
  value: unknown,
  type: "points" | "percent",
  path: string,
  operationName: string,
): void {
  assertFiniteIntegerNumber(value, operationName, path);
  const maximum = type === "points" ? 158400 : 13200000;
  if (typeof value !== "number" || value < 0 || value > maximum) {
    throw new Error(`${operationName}: ${path} must be between 0 and ${maximum}`);
  }
}

function assertTextBoxBullet(bullet: unknown, path: string, operationName: string): void {
  if (!isPlainRecord(bullet)) {
    throw new Error(`${operationName}: ${path} must be a bullet object`);
  }
  if (bullet.type === "none") return;
  if (bullet.type !== "character" && bullet.type !== "auto-number") {
    throw new Error(`${operationName}: ${path}.type is not supported`);
  }
  if (bullet.fontFace !== undefined) {
    if (typeof bullet.fontFace !== "string" || bullet.fontFace.trim() === "") {
      throw new Error(`${operationName}: ${path}.fontFace must be a non-empty string`);
    }
  }
  if (bullet.size !== undefined) {
    assertFiniteIntegerNumber(bullet.size, operationName, `${path}.size`);
    if (typeof bullet.size !== "number" || bullet.size < 25000 || bullet.size > 400000) {
      throw new Error(`${operationName}: ${path}.size must be between 25000 and 400000`);
    }
  }
  if (bullet.type === "character") {
    if (typeof bullet.character !== "string" || bullet.character.length === 0) {
      throw new Error(`${operationName}: ${path}.character must be a non-empty string`);
    }
    return;
  }
  if (typeof bullet.scheme !== "string" || !AUTO_NUMBER_SCHEMES.has(bullet.scheme)) {
    throw new Error(`${operationName}: ${path}.scheme is not supported`);
  }
  if (bullet.startAt !== undefined) {
    assertFiniteIntegerNumber(bullet.startAt, operationName, `${path}.startAt`);
    if (typeof bullet.startAt !== "number" || bullet.startAt < 1 || bullet.startAt > 32767) {
      throw new Error(`${operationName}: ${path}.startAt must be between 1 and 32767`);
    }
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
  if (run.hyperlink !== undefined) {
    assertExternalHttpHyperlink(run.hyperlink, `${path}.hyperlink`, operationName);
  }
  assertTextBoxRunProperties(run.properties, path, operationName);
}

function assertExternalHttpHyperlink(
  hyperlink: unknown,
  path: string,
  operationName: string,
): void {
  if (typeof hyperlink !== "string" || hyperlink.trim() !== hyperlink || hyperlink === "") {
    throw new Error(`${operationName}: ${path} must be an absolute HTTP(S) URL`);
  }
  try {
    const url = new URL(hyperlink);
    if (url.protocol !== "http:" && url.protocol !== "https:") {
      throw new Error("unsupported protocol");
    }
  } catch {
    throw new Error(`${operationName}: ${path} must be an absolute HTTP(S) URL`);
  }
}

interface RunHyperlinkAllocation {
  readonly packageGraph: PptxSourceModel["packageGraph"];
  readonly ids: ReadonlyMap<string, RelationshipId>;
}

function allocateRunHyperlinks(
  source: PptxSourceModel,
  slidePartPath: PartPath,
  paragraphs: readonly TextBoxParagraphInput[],
): RunHyperlinkAllocation {
  const group = source.packageGraph.relationships.find(
    (relationships) => relationships.sourcePartPath === slidePartPath,
  );
  let relationships = group?.relationships ?? [];
  const ids = new Map<string, RelationshipId>();
  for (const paragraph of paragraphs) {
    for (const run of paragraph.runs) {
      if (run.hyperlink === undefined || ids.has(run.hyperlink)) continue;
      const id = nextRelationshipId(relationships);
      ids.set(run.hyperlink, id);
      relationships = [
        ...relationships,
        {
          id,
          type: HYPERLINK_REL_TYPE,
          target: run.hyperlink,
          targetMode: "External",
        } satisfies Relationship,
      ];
    }
  }
  if (ids.size === 0) return { packageGraph: source.packageGraph, ids };

  return {
    packageGraph: {
      ...source.packageGraph,
      relationships:
        group === undefined
          ? [...source.packageGraph.relationships, { sourcePartPath: slidePartPath, relationships }]
          : source.packageGraph.relationships.map((candidate) =>
              candidate === group ? { ...candidate, relationships } : candidate,
            ),
    },
    ids,
  };
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
  if (color.transforms !== undefined) {
    if (!isArrayValue(color.transforms)) {
      throw new Error(`${operationName}: ${path}.transforms must be an array`);
    }
    color.transforms.forEach((transform: unknown, index: number) => {
      if (!isPlainRecord(transform) || transform.kind !== "alpha") {
        throw new Error(
          `${operationName}: ${path}.transforms[${index}] must be an alpha transform`,
        );
      }
      assertOoxmlPercent(transform.value, operationName, `${path}.transforms[${index}].value`);
    });
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
  if (fill.gradientType !== "linear" && fill.gradientType !== "radial") {
    throw new Error(`${operationName}: ${path}.gradientType must be linear or radial`);
  }
  const stops = fill.stops;
  if (!isArrayValue(stops) || stops.length < 2) {
    throw new Error(`${operationName}: ${path}.stops must contain at least two stops`);
  }
  if (fill.gradientType === "linear") {
    if (fill.angle !== undefined)
      assertFiniteIntegerNumber(fill.angle, operationName, `${path}.angle`);
  } else {
    assertOoxmlPercent(fill.centerX, operationName, `${path}.centerX`);
    assertOoxmlPercent(fill.centerY, operationName, `${path}.centerY`);
  }
  stops.forEach((stop: unknown, index: number) => {
    if (!isPlainRecord(stop)) {
      throw new Error(`${operationName}: ${path}.stops[${index}] must be an object`);
    }
    const position = stop.position;
    assertOoxmlPercent(position, operationName, `${path}.stops[${index}].position`);
    assertTextBoxColor(stop.color, `${path}.stops[${index}].color`, operationName);
  });
}

function assertOoxmlPercent(value: unknown, operationName: string, path: string): void {
  assertFiniteIntegerNumber(value, operationName, path);
  if (typeof value !== "number" || value < 0 || value > 100000) {
    throw new Error(`${operationName}: ${path} must be between 0 and 100000`);
  }
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
  if (isPlainRecord(baseline) && baseline.type === "percent") {
    assertFiniteIntegerNumber(baseline.value, operationName, `${path}.value`);
    if (
      typeof baseline.value === "number" &&
      baseline.value >= -400000 &&
      baseline.value <= 400000
    ) {
      return;
    }
  }
  throw new Error(
    `${operationName}: ${path} must be subscript, superscript, or a percent between -400000 and 400000`,
  );
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
  if (body.autoFit !== undefined && body.autoFit !== "shape") {
    throw new Error(`${operationName}: body.autoFit is not supported`);
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
  assertShadowEffectsInput(effects, "addShape");
  if (
    effects.glow === undefined &&
    effects.outerShadow === undefined &&
    effects.innerShadow === undefined
  ) {
    throw new Error("addShape: effects must set glow, outerShadow, or innerShadow");
  }
  if (effects.glow !== undefined) assertTextBoxGlow(effects.glow, "effects.glow", "addShape");
}

function assertConnectionSiteIndex(value: number, fieldName: "start" | "end"): void {
  if (!Number.isInteger(value) || value < 0 || value > 0xffffffff) {
    throw new Error(
      `addConnector: ${fieldName}.connectionSiteIndex must be an unsigned 32-bit integer`,
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

function assertPositiveFiniteIntegerNumber(
  value: unknown,
  operationName: string,
  fieldName: string,
): void {
  assertFiniteIntegerNumber(value, operationName, fieldName);
  if (typeof value !== "number" || value <= 0) {
    throw new Error(`${operationName}: ${fieldName} must be a finite positive integer`);
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

function assertNonNegativeFiniteEmu(
  value: unknown,
  operationName: string,
  fieldName: string,
): void {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    throw new Error(`${operationName}: ${fieldName} must be a finite non-negative EMU value`);
  }
}

function requireConnectorTargetShape(
  target: AuthoringTarget,
  handle: SourceHandle,
  endpointName: "start" | "end",
): SourceShape & { readonly nodeId: SourceNodeId } {
  if (handle.partPath !== target.partPath) {
    throw new Error(
      `addConnector: ${endpointName} target shape belongs to a different drawing part`,
    );
  }
  const shape = target.shapes.find((candidate) => sourceHandlesEqual(candidate.handle, handle));
  if (shape === undefined) {
    throw new Error(
      `addConnector: ${endpointName} shape handle was not found in the target drawing part`,
    );
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

function findAuthoringTarget(
  source: PptxSourceModel,
  handle: SourceHandle,
): AuthoringTarget | undefined {
  const slideIndex = source.slides.findIndex((candidate) =>
    sourceHandlesEqual(candidate.handle, handle),
  );
  if (slideIndex >= 0) {
    const slide = source.slides[slideIndex];
    return { kind: "slide", index: slideIndex, partPath: slide.partPath, shapes: slide.shapes };
  }
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

function withAuthoringTargetShapes(
  source: PptxSourceModel,
  target: AuthoringTarget,
  shapes: readonly SourceShapeNode[],
): PptxSourceModel {
  switch (target.kind) {
    case "slide":
      return {
        ...source,
        slides: source.slides.map((candidate, index) =>
          index === target.index ? { ...candidate, shapes } : candidate,
        ),
      };
    case "layout":
      return {
        ...source,
        slideLayouts: source.slideLayouts.map((candidate, index) =>
          index === target.index ? { ...candidate, shapes } : candidate,
        ),
      };
    case "master":
      return {
        ...source,
        slideMasters: source.slideMasters.map((candidate, index) =>
          index === target.index ? { ...candidate, shapes } : candidate,
        ),
      };
  }
}

function defaultTextBodyProperties(
  source: PptxSourceModel,
  target: AuthoringTarget,
): AddTextBoxBodyPropertiesInput | undefined {
  if (target.kind !== "slide") return undefined;
  const slide = source.slides[target.index];
  const defaults = source.slideLayouts.find(
    (layout) => layout.partPath === slide.layoutPartPath,
  )?.defaultTextBodyProperties;
  if (defaults === undefined) return undefined;
  return {
    ...(defaults.anchor !== undefined ? { anchor: defaults.anchor } : {}),
    ...(defaults.marginLeft !== undefined ? { marginLeft: defaults.marginLeft } : {}),
    ...(defaults.marginRight !== undefined ? { marginRight: defaults.marginRight } : {}),
    ...(defaults.marginTop !== undefined ? { marginTop: defaults.marginTop } : {}),
    ...(defaults.marginBottom !== undefined ? { marginBottom: defaults.marginBottom } : {}),
    ...(defaults.autoFit === "spAutofit" ? { autoFit: "shape" as const } : {}),
  };
}

function mergeTextBodyProperties(
  defaults: AddTextBoxBodyPropertiesInput | undefined,
  explicit: AddTextBoxBodyPropertiesInput | undefined,
): AddTextBoxBodyPropertiesInput | undefined {
  if (defaults === undefined) return explicit;
  if (explicit === undefined) return defaults;
  return { ...defaults, ...explicit };
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
