/**
 * Mutable orchestration for consecutive PptxSourceModel authoring operations.
 *
 * The session owns only the current immutable source value and a target handle. Each
 * operation delegates validation, XML generation, and allocation to the existing
 * function API, then exposes the handle of the typed node produced by that operation.
 */

import type { AddChartInput } from "./chart-authoring.js";
import { addChart } from "./chart-authoring.js";
import { editInsertedShape, editInsertedSlidePartPath } from "./edit-descriptors.js";
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
import { reorderShapes } from "./shape-ordering.js";
import type { SetSlideBackgroundInput } from "./slide-background-authoring.js";
import { setSlideBackground } from "./slide-background-authoring.js";
import type { AddEmptySlideFromLayoutInput } from "./slide-topology.js";
import { addEmptySlideFromLayout } from "./slide-topology.js";
import type { AddTableInput } from "./table-authoring.js";
import { addTable } from "./table-authoring.js";

/** Consecutive authoring operations bound to one slide, layout, or master handle. */
export interface PptxAuthoringTarget {
  addTextBox(input: AddTextBoxInput): SourceHandle;
  addSlideNumber(input: AddSlideNumberInput): SourceHandle;
  addShape(input: AddShapeInput): SourceHandle;
  addConnector(input: AddConnectorInput): SourceHandle;
  addPicture(input: AddPictureInput): SourceHandle;
  addTable(input: AddTableInput): SourceHandle;
  addChart(input: AddChartInput): SourceHandle;
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

  /** Creates an authoring scope bound to a slide, layout, or master handle. */
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
      setSlideBackground: (input) => {
        this.#source = setSlideBackground(this.#source, targetHandle, input);
      },
      reorderShapes: (orderedShapeHandles) => {
        this.#source = reorderShapes(this.#source, targetHandle, orderedShapeHandles);
      },
    };
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

/** Creates a mutable authoring session around any PptxSourceModel. */
export function createPptxAuthoringSession(source: PptxSourceModel): PptxAuthoringSession {
  return new PptxAuthoringSession(source);
}
