import {
  asEmu,
  findShapeNodeBySourceHandle,
  type PptxSourceModel,
  readPptx,
  type SourceHandle,
  type SourceShapeNode,
  type SourceTextBody,
  type SourceTextRun,
  writePptx,
} from "@pptx-glimpse/document";
import {
  createEditorSession,
  type EditorCommand,
  type EditorCommandWarning,
  type PptxTextBodyProseMirrorDocJson,
  proseMirrorDocJsonToEditorCommands,
  textBodyToProseMirrorDocJson,
} from "@pptx-glimpse/editor-core";

import { type ConvertOptions, renderPptxSourceModelToSvg, type SlideSvg } from "./svg-converter.js";

const EMU_PER_INCH = 914400;
const DEFAULT_DPI = 96;
const EMU_PER_PIXEL = EMU_PER_INCH / DEFAULT_DPI;

const DEFAULT_TEXT_BOX_BOUNDS_PX = {
  x: 96,
  y: 96,
  width: 288,
  height: 72,
};
const DEFAULT_TEXT_BOX_TEXT = "New text box";

export interface BrowserEditorHistoryState {
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly undoDepth: number;
  readonly redoDepth: number;
}

export interface BrowserEditorSelectionInfo {
  readonly shapeHandle: SourceHandle;
}

export interface BrowserEditorTextRunInfo {
  readonly text: string;
  readonly handle: SourceHandle;
}

export interface BrowserEditorTextBodyInfo {
  readonly docJson: PptxTextBodyProseMirrorDocJson;
}

export interface BrowserEditorShapeBoundsPx {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface BrowserEditorShapeInfo {
  readonly id: string;
  readonly kind: SourceShapeNode["kind"];
  readonly name?: string;
  readonly handle?: SourceHandle;
  readonly bounds?: BrowserEditorShapeBoundsPx;
  readonly editableTransform?: boolean;
  readonly editableDelete?: boolean;
  readonly textRuns?: readonly BrowserEditorTextRunInfo[];
  readonly editableTextBody?: BrowserEditorTextBodyInfo;
}

export interface BrowserEditorAddTextBoxOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly text?: string;
  readonly name?: string;
}

export interface BrowserEditorSlidesResponse {
  readonly slides: readonly SlideSvg[];
  readonly history: BrowserEditorHistoryState;
  readonly selection?: BrowserEditorSelectionInfo;
  readonly warnings?: readonly EditorCommandWarning[];
}

export interface BrowserEditorSaveResponse {
  readonly ok: true;
  readonly pptx: Uint8Array;
  readonly history: BrowserEditorHistoryState;
}

export type BrowserEditorRenderOptions = Omit<ConvertOptions, "slides">;

export class BrowserPptxEditorSession {
  #session: ReturnType<typeof createEditorSession>;
  #slides: readonly SlideSvg[] = [];
  readonly #renderOptions: BrowserEditorRenderOptions;

  private constructor(source: PptxSourceModel, renderOptions: BrowserEditorRenderOptions) {
    this.#session = createEditorSession(source);
    this.#renderOptions = renderOptions;
  }

  static async create(
    input: Uint8Array,
    renderOptions: BrowserEditorRenderOptions = {},
  ): Promise<BrowserPptxEditorSession> {
    const editor = new BrowserPptxEditorSession(readPptx(input), renderOptions);
    await editor.renderCurrentSlides();
    return editor;
  }

  get document(): PptxSourceModel {
    return this.#session.document;
  }

  get slides(): readonly SlideSvg[] {
    return this.#slides;
  }

  get history(): BrowserEditorHistoryState {
    return {
      canUndo: this.#session.canUndo,
      canRedo: this.#session.canRedo,
      undoDepth: this.#session.undoDepth,
      redoDepth: this.#session.redoDepth,
    };
  }

  get selection(): BrowserEditorSelectionInfo | undefined {
    return this.#session.selection;
  }

  response(warnings?: readonly EditorCommandWarning[]): BrowserEditorSlidesResponse {
    return {
      slides: this.#slides,
      history: this.history,
      ...(this.selection !== undefined ? { selection: this.selection } : {}),
      ...(warnings !== undefined && warnings.length > 0 ? { warnings } : {}),
    };
  }

  shapes(slideNumber: number): readonly BrowserEditorShapeInfo[] {
    const slide = this.#session.document.slides[slideNumber - 1];
    if (slide === undefined) return [];
    return slide.shapes.flatMap((shape, index) => shapeInfo(shape, index, true, slide.shapes));
  }

  async renderCurrentSlides(): Promise<readonly SlideSvg[]> {
    const report = await renderPptxSourceModelToSvg(this.#session.document, {
      textOutput: "text",
      skipSystemFonts: true,
      ...this.#renderOptions,
    });
    this.#slides = report.slides;
    return this.#slides;
  }

  async apply(command: EditorCommand): Promise<BrowserEditorSlidesResponse> {
    const result = this.#session.apply(command);
    if (!result.ok) {
      throw new Error(result.message);
    }
    await this.renderCurrentSlides();
    return this.response(result.warnings);
  }

  async addTextBox(
    slideNumber = 1,
    options: BrowserEditorAddTextBoxOptions = {},
  ): Promise<BrowserEditorSlidesResponse> {
    const slide = this.#session.document.slides[slideNumber - 1];
    if (slide?.handle === undefined) {
      throw new Error("addTextBox: slide handle was not found in PptxSourceModel source");
    }
    const existingShapeKeys = new Set(slide.shapes.map(shapeSourceKey));
    const result = this.#session.apply({
      kind: "addTextBox",
      slideHandle: slide.handle,
      offsetX: pxToEmu(options.x ?? DEFAULT_TEXT_BOX_BOUNDS_PX.x),
      offsetY: pxToEmu(options.y ?? DEFAULT_TEXT_BOX_BOUNDS_PX.y),
      width: pxToEmu(options.width ?? DEFAULT_TEXT_BOX_BOUNDS_PX.width),
      height: pxToEmu(options.height ?? DEFAULT_TEXT_BOX_BOUNDS_PX.height),
      text: options.text ?? DEFAULT_TEXT_BOX_TEXT,
      ...(options.name !== undefined ? { name: options.name } : {}),
    });
    if (!result.ok) {
      throw new Error(result.message);
    }
    this.#selectNewShape(slideNumber, existingShapeKeys);
    await this.renderCurrentSlides();
    return this.response(result.warnings);
  }

  async deleteShape(handle: SourceHandle): Promise<BrowserEditorSlidesResponse> {
    const result = this.#session.apply({ kind: "deleteShape", handle });
    if (!result.ok) {
      throw new Error(result.message);
    }
    await this.renderCurrentSlides();
    return this.response(result.warnings);
  }

  async deleteSelectedShape(): Promise<BrowserEditorSlidesResponse> {
    const selection = this.#session.selection;
    if (selection === undefined) {
      throw new Error("deleteShape: no selected shape");
    }
    return this.deleteShape(selection.shapeHandle);
  }

  async applyTextBodyDocJson(
    handle: SourceHandle,
    docJson: unknown,
  ): Promise<BrowserEditorSlidesResponse> {
    const textBody = this.#requireEditableShapeTextBody(handle);
    const commands = proseMirrorDocJsonToEditorCommands(textBody, docJson);
    if (commands.length === 0) return this.response();

    const result = this.#session.applyAll(commands);
    if (!result.ok) {
      throw new Error(result.message);
    }
    await this.renderCurrentSlides();
    return this.response(result.warnings);
  }

  selectShape(handle: SourceHandle): BrowserEditorSlidesResponse {
    const result = this.#session.selectShape(handle);
    if (!result.ok) {
      throw new Error(result.message);
    }
    return this.response();
  }

  async undo(): Promise<BrowserEditorSlidesResponse> {
    const result = this.#session.undo();
    if (!result.ok) {
      throw new Error(result.reason);
    }
    await this.renderCurrentSlides();
    return this.response();
  }

  async redo(): Promise<BrowserEditorSlidesResponse> {
    const result = this.#session.redo();
    if (!result.ok) {
      throw new Error(result.reason);
    }
    await this.renderCurrentSlides();
    return this.response();
  }

  save(): BrowserEditorSaveResponse {
    const output = writePptx(this.#session.document);
    readPptx(output);
    return { ok: true, pptx: output, history: this.history };
  }

  #requireEditableShapeTextBody(handle: SourceHandle): SourceTextBody {
    const shape = findShapeNodeBySourceHandle(this.#session.document, handle);
    if (shape === undefined) {
      throw new Error("text body edit: shape handle was not found in PptxSourceModel source");
    }
    if (shape.kind !== "shape" || shape.textBody === undefined) {
      throw new Error("text body edit: shape does not have editable text body");
    }
    return shape.textBody;
  }

  #selectNewShape(slideNumber: number, existingShapeKeys: ReadonlySet<string>): void {
    const slide = this.#session.document.slides[slideNumber - 1];
    const addedShape = slide?.shapes.find((shape) => !existingShapeKeys.has(shapeSourceKey(shape)));
    if (addedShape?.handle === undefined) return;
    this.#session.selectShape(addedShape.handle);
  }
}

export function createBrowserPptxEditorSession(
  input: Uint8Array,
  renderOptions?: BrowserEditorRenderOptions,
): Promise<BrowserPptxEditorSession> {
  return BrowserPptxEditorSession.create(input, renderOptions);
}

function shapeInfo(
  shape: SourceShapeNode,
  index: number,
  editableTransform = true,
  slideShapes: readonly SourceShapeNode[] = [],
): BrowserEditorShapeInfo[] {
  const canEditTransform =
    shape.kind !== "raw" &&
    shape.transform !== undefined &&
    editableTransform &&
    isEditableTransformShape(shape);
  const base: BrowserEditorShapeInfo = {
    id: String(shape.nodeId ?? shape.handle?.nodeId ?? `${shape.kind}:${String(index)}`),
    kind: shape.kind,
    ...(shapeName(shape) !== undefined ? { name: shapeName(shape) } : {}),
    ...(shape.handle !== undefined ? { handle: shape.handle } : {}),
    ...(canEditTransform
      ? {
          bounds: transformBoundsPx(shape.transform),
          editableTransform: true,
        }
      : {}),
    ...(editableTransform && isDeletableShape(shape, slideShapes) ? { editableDelete: true } : {}),
    ...("textBody" in shape && shape.textBody !== undefined
      ? {
          textRuns: collectTextRuns(
            shape.textBody.paragraphs.flatMap((paragraph) => paragraph.runs),
          ),
        }
      : {}),
    ...(shape.kind === "shape" && canEditTransform && shape.textBody !== undefined
      ? editableTextBody(shape.textBody)
      : {}),
  };

  if (shape.kind !== "group") return [base];
  return [
    base,
    ...shape.children.flatMap((child, childIndex) =>
      shapeInfo(child, childIndex, false, slideShapes),
    ),
  ];
}

function shapeName(shape: SourceShapeNode): string | undefined {
  return "name" in shape ? shape.name : undefined;
}

function isEditableTransformShape(shape: SourceShapeNode): boolean {
  if (shape.kind === "raw" || shape.transform === undefined || shape.handle?.nodeId === undefined) {
    return false;
  }
  return !shape.rawSidecars?.some((sidecar) => sidecar.node.name === "mc:AlternateContent");
}

function isDeletableShape(
  shape: SourceShapeNode,
  slideShapes: readonly SourceShapeNode[],
): boolean {
  if (shape.kind !== "shape" || shape.handle?.nodeId === undefined) {
    return false;
  }
  if (isShapeReferencedByConnector(shape, slideShapes)) {
    return false;
  }
  return !shape.rawSidecars?.some((sidecar) => sidecar.node.name === "mc:AlternateContent");
}

function isShapeReferencedByConnector(
  shape: SourceShapeNode,
  slideShapes: readonly SourceShapeNode[],
): boolean {
  return slideShapes.some(
    (candidate) =>
      candidate.kind === "connector" &&
      (candidate.connection?.start?.shapeId === shape.nodeId ||
        candidate.connection?.end?.shapeId === shape.nodeId),
  );
}

function collectTextRuns(runs: readonly SourceTextRun[]): BrowserEditorTextRunInfo[] {
  return runs.flatMap((run) => {
    if (run.handle === undefined) return [];
    return [{ text: run.text, handle: run.handle }];
  });
}

function editableTextBody(
  textBody: SourceTextBody,
): Partial<Pick<BrowserEditorShapeInfo, "editableTextBody">> {
  try {
    return { editableTextBody: { docJson: textBodyToProseMirrorDocJson(textBody) } };
  } catch {
    return {};
  }
}

function transformBoundsPx(transform: {
  readonly offsetX: number;
  readonly offsetY: number;
  readonly width: number;
  readonly height: number;
}): BrowserEditorShapeBoundsPx {
  return {
    x: emuToPixels(transform.offsetX),
    y: emuToPixels(transform.offsetY),
    width: emuToPixels(transform.width),
    height: emuToPixels(transform.height),
  };
}

function emuToPixels(value: number): number {
  return (value / EMU_PER_INCH) * DEFAULT_DPI;
}

function pxToEmu(value: number): ReturnType<typeof asEmu> {
  return asEmu(Math.round(value * EMU_PER_PIXEL));
}

function shapeSourceKey(shape: SourceShapeNode): string {
  const handle = shape.handle;
  return [
    shape.kind,
    handle?.partPath ?? "",
    handle?.nodeId ?? "",
    handle?.relationshipId ?? "",
    handle?.orderingSlot ?? "",
  ].join("\u0000");
}
