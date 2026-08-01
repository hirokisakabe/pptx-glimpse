/**
 * Mutable orchestration for consecutive PptxSourceModel authoring operations.
 *
 * The session owns only the current immutable source value and a target handle. Each
 * operation delegates validation, XML generation, and allocation to the existing
 * function API, then exposes the handle of the typed node produced by that operation.
 */

import type { AddChartInput } from "./chart-authoring.js";
import { addChart } from "./chart-authoring.js";
import {
  editInsertedShape,
  editInsertedSlidePartPath,
  sourceHandlesEqual,
} from "./edit-descriptors.js";
import type { SourceHandle } from "./handles.js";
import type { AddPictureInput } from "./picture-authoring.js";
import { addPicture } from "./picture-authoring.js";
import type { PptxSourceModel } from "./pptx-source-model.js";
import type {
  AddConnectorInput,
  AddShapeInput,
  AddSlideNumberInput,
  AddTextBoxInput,
} from "./shape-authoring.js";
import { addConnector, addShape, addSlideNumber, addTextBox } from "./shape-authoring.js";
import { groupShapes } from "./shape-grouping.js";
import { reorderShapes } from "./shape-ordering.js";
import type { SourceGroup, SourceShapeNode } from "./shapes.js";
import type { SetBackgroundInput, SetSlideBackgroundInput } from "./slide-background-authoring.js";
import {
  clearBackground,
  setBackground,
  setSlideBackground,
} from "./slide-background-authoring.js";
import type { AddSlideLayoutInput, CloneSlideLayoutInput } from "./slide-layout-authoring.js";
import { addSlideLayout, cloneSlideLayout } from "./slide-layout-authoring.js";
import type { AddEmptySlideFromLayoutInput } from "./slide-topology.js";
import { addEmptySlideFromLayout } from "./slide-topology.js";
import type { AddTableInput } from "./table-authoring.js";
import { addTable } from "./table-authoring.js";

/**
 * Consecutive authoring operations bound to one slide, layout, or master handle. A native group
 * handle is supported only by `groupShapes`, which groups its consecutive direct children.
 */
export interface PptxAuthoringTarget {
  addTextBox(input: AddTextBoxInput): SourceHandle;
  addSlideNumber(input: AddSlideNumberInput): SourceHandle;
  addShape(input: AddShapeInput): SourceHandle;
  addConnector(input: AddConnectorInput): SourceHandle;
  addPicture(input: AddPictureInput): SourceHandle;
  addTable(input: AddTableInput): SourceHandle;
  addChart(input: AddChartInput): SourceHandle;
  /**
   * Group consecutive direct children and return the new native group handle.
   *
   * Supported from-scratch children are handles returned by `addShape`, `addPicture`,
   * `addConnector`, `addTable`, `addChart`, or an earlier `groupShapes` call. The target may be
   * a slide/layout/master or an authored native group; every supplied handle must be its direct
   * child. SmartArt, raw preserved nodes, and consumer-specific composite inputs are not part of
   * the from-scratch contract.
   */
  groupShapes(shapeHandles: readonly SourceHandle[]): SourceHandle;
  setBackground(input: SetBackgroundInput): void;
  clearBackground(): void;
  /** @deprecated Use `setBackground`. */
  setSlideBackground(input: SetSlideBackgroundInput): void;
  reorderShapes(orderedShapeHandles: readonly SourceHandle[]): void;
}

type DrawingOperationInput =
  | AddTextBoxInput
  | AddSlideNumberInput
  | AddShapeInput
  | AddConnectorInput
  | AddPictureInput
  | AddTableInput
  | AddChartInput;

type DrawingOperation<Input extends DrawingOperationInput> = (
  source: PptxSourceModel,
  targetHandle: SourceHandle,
  input: Input,
) => PptxSourceModel;

/** Owns the latest immutable source while consecutive authoring operations are applied. */
export class PptxAuthoringSession {
  #source: PptxSourceModel;

  constructor(source: PptxSourceModel) {
    this.#source = source;
  }

  /** Latest source including every operation that completed successfully. */
  get source(): PptxSourceModel {
    return this.#source;
  }

  /**
   * Creates an authoring scope bound to a slide, layout, or master handle. A native group handle
   * is also accepted for the `groupShapes` direct-child operation.
   */
  target(targetHandle: SourceHandle): PptxAuthoringTarget {
    return {
      addTextBox: (input) => this.#addDrawing("addTextBox", targetHandle, input, addTextBox),
      addSlideNumber: (input) =>
        this.#addDrawing("addSlideNumber", targetHandle, input, addSlideNumber),
      addShape: (input) => this.#addDrawing("addShape", targetHandle, input, addShape),
      addConnector: (input) => this.#addDrawing("addConnector", targetHandle, input, addConnector),
      addPicture: (input) => this.#addDrawing("addPicture", targetHandle, input, addPicture),
      addTable: (input) => this.#addDrawing("addTable", targetHandle, input, addTable),
      addChart: (input) => this.#addDrawing("addChart", targetHandle, input, addChart),
      groupShapes: (shapeHandles) => {
        this.#assertDirectChildren(targetHandle, shapeHandles);
        const updated = groupShapes(this.#source, shapeHandles);
        const edit = updated.edits?.at(-1);
        const inserted = edit === undefined ? undefined : editInsertedShape(edit);
        const handle =
          inserted === undefined
            ? undefined
            : findShapeHandle(updated, inserted.slidePartPath, inserted.shapeId);
        if (handle === undefined) {
          throw new Error(
            "PptxAuthoringSession.groupShapes: operation did not produce a group handle",
          );
        }
        this.#source = updated;
        return handle;
      },
      setBackground: (input) => {
        this.#source = setBackground(this.#source, targetHandle, input);
      },
      clearBackground: () => {
        this.#source = clearBackground(this.#source, targetHandle);
      },
      setSlideBackground: (input) => {
        this.#source = setSlideBackground(this.#source, targetHandle, input);
      },
      reorderShapes: (orderedShapeHandles) => {
        this.#source = reorderShapes(this.#source, targetHandle, orderedShapeHandles);
      },
    };
  }

  #assertDirectChildren(targetHandle: SourceHandle, shapeHandles: readonly SourceHandle[]): void {
    const root = [
      ...this.#source.slides,
      ...this.#source.slideLayouts,
      ...this.#source.slideMasters,
    ].find((candidate) => sourceHandlesEqual(candidate.handle, targetHandle));
    const group = root === undefined ? findGroupByHandle(this.#source, targetHandle) : undefined;
    const children = root?.shapes ?? group?.children;
    if (children === undefined) {
      throw new Error(
        "PptxAuthoringSession.groupShapes: target must be a slide, layout, master, or native group",
      );
    }
    if (
      shapeHandles.some(
        (handle) => !children.some((child) => sourceHandlesEqual(child.handle, handle)),
      )
    ) {
      throw new Error(
        "PptxAuthoringSession.groupShapes: every shape must be a direct child of the target",
      );
    }
  }

  /** Adds a slide and returns its stable source handle directly. */
  addEmptySlideFromLayout(input: AddEmptySlideFromLayoutInput): SourceHandle {
    const updated = addEmptySlideFromLayout(this.#source, input);
    const edit = updated.edits?.at(-1);
    const insertedPartPath = edit === undefined ? undefined : editInsertedSlidePartPath(edit);
    const handle = updated.slides.find((slide) => slide.partPath === insertedPartPath)?.handle;
    if (handle === undefined) {
      throw new Error(
        "PptxAuthoringSession.addEmptySlideFromLayout: operation did not produce a slide handle",
      );
    }
    this.#source = updated;
    return handle;
  }

  /** Adds an empty layout below one existing master and returns its authoring target handle. */
  addSlideLayout(masterHandle: SourceHandle, input: AddSlideLayoutInput): SourceHandle {
    const updated = addSlideLayout(this.#source, masterHandle, input);
    const edit = updated.edits?.at(-1);
    const handle =
      edit?.kind === "addSlideLayout"
        ? updated.slideLayouts.find((layout) => layout.partPath === edit.newLayoutPartPath)?.handle
        : undefined;
    if (handle === undefined) {
      throw new Error(
        "PptxAuthoringSession.addSlideLayout: operation did not produce a layout handle",
      );
    }
    this.#source = updated;
    return handle;
  }

  /** Clones a preserved layout within its master and returns its authoring target handle. */
  cloneSlideLayout(layoutHandle: SourceHandle, input: CloneSlideLayoutInput): SourceHandle {
    const updated = cloneSlideLayout(this.#source, layoutHandle, input);
    const edit = updated.edits?.at(-1);
    const handle =
      edit?.kind === "cloneSlideLayout"
        ? updated.slideLayouts.find((layout) => layout.partPath === edit.newLayoutPartPath)?.handle
        : undefined;
    if (handle === undefined) {
      throw new Error(
        "PptxAuthoringSession.cloneSlideLayout: operation did not produce a layout handle",
      );
    }
    this.#source = updated;
    return handle;
  }

  #addDrawing<Input extends DrawingOperationInput>(
    operationName: string,
    targetHandle: SourceHandle,
    input: Input,
    operation: DrawingOperation<Input>,
  ): SourceHandle {
    const updated = operation(this.#source, targetHandle, input);
    const edit = updated.edits?.at(-1);
    const inserted = edit === undefined ? undefined : editInsertedShape(edit);
    const target = [...updated.slides, ...updated.slideLayouts, ...updated.slideMasters].find(
      (candidate) => candidate.partPath === inserted?.slidePartPath,
    );
    const handle = target?.shapes.find(
      (shape) => String(shape.nodeId) === inserted?.shapeId,
    )?.handle;
    if (handle === undefined) {
      throw new Error(
        `PptxAuthoringSession.${operationName}: operation did not produce a drawing handle`,
      );
    }
    this.#source = updated;
    return handle;
  }
}

function findShapeHandle(
  source: PptxSourceModel,
  partPath: string,
  shapeId: string,
): SourceHandle | undefined {
  const visit = (nodes: readonly SourceShapeNode[]): SourceHandle | undefined => {
    for (const node of nodes) {
      if (node.handle?.partPath === partPath && String(node.nodeId) === shapeId) return node.handle;
      if (node.kind === "group") {
        const nested = visit(node.children);
        if (nested !== undefined) return nested;
      }
    }
    return undefined;
  };
  for (const target of [...source.slides, ...source.slideLayouts, ...source.slideMasters]) {
    const handle = visit(target.shapes);
    if (handle !== undefined) return handle;
  }
  return undefined;
}

function findGroupByHandle(source: PptxSourceModel, handle: SourceHandle): SourceGroup | undefined {
  const visit = (nodes: readonly SourceShapeNode[]): SourceGroup | undefined => {
    for (const node of nodes) {
      if (node.kind !== "group") continue;
      if (sourceHandlesEqual(node.handle, handle)) return node;
      const nested = visit(node.children);
      if (nested !== undefined) return nested;
    }
    return undefined;
  };
  for (const target of [...source.slides, ...source.slideLayouts, ...source.slideMasters]) {
    const group = visit(target.shapes);
    if (group !== undefined) return group;
  }
  return undefined;
}

/** Creates a mutable authoring session around any PptxSourceModel. */
export function createPptxAuthoringSession(source: PptxSourceModel): PptxAuthoringSession {
  return new PptxAuthoringSession(source);
}
