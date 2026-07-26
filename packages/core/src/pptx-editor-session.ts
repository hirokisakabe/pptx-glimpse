/**
 * High-level read/edit/render/write session shared by the Node and browser entries.
 *
 * Headless expected failures are unwrapped into PptxEditorError without changing their
 * code, message, or cause. Read, render, and write integrations are wrapped only at
 * their owning boundaries. See the repository's
 * [editor error contract](../../../docs/editor-error-contract.md).
 */

import {
  asEmu,
  findShapeNodeBySourceHandle,
  type MediaPart,
  type PartPath,
  type PptxSourceModel,
  readPptx,
  type Relationship,
  type RelationshipId,
  type SourceHandle,
  type SourceImage,
  type SourceParagraphProperties,
  type SourceRunProperties,
  type SourceShapeNode,
  type SourceTextBody,
  type SourceTextBodyProperties,
  type SourceTextRun,
  writePptx,
} from "@pptx-glimpse/document";
import {
  createEditorSession,
  type EditorCommand,
  type EditorCommandWarning,
  type EditorOperationErrorCode,
  type EditorOperationFailure,
} from "@pptx-glimpse/editor";

import {
  type PptxTextBodyProseMirrorDocJson,
  proseMirrorDocJsonToEditorCommands,
  textBodyToProseMirrorDocJson,
} from "./prosemirror-text-body-compat.js";
import {
  type ConvertOptions,
  renderPptxSourceModelToSvg,
  type SlideSvg,
  type SvgConversionReport,
} from "./svg-converter.js";
import { unsafeBrandAssertion } from "./unsafe-type-assertion.js";

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
const DEFAULT_CONNECTOR_BOUNDS_PX = {
  x: 144,
  y: 144,
  width: 288,
  height: 96,
};
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const IMAGE_ACCEPT_BY_CONTENT_TYPE: Readonly<Record<string, string>> = {
  "image/png": "image/png,.png",
  "image/jpeg": "image/jpeg,.jpg,.jpeg",
  "image/gif": "image/gif,.gif",
  "image/bmp": "image/bmp,.bmp",
  "image/tiff": "image/tiff,.tif,.tiff",
  "image/webp": "image/webp,.webp",
};

/**
 * Undo/redo availability and stack depth after the latest editor operation.
 */
export interface PptxEditorHistoryState {
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly undoDepth: number;
  readonly redoDepth: number;
}

/**
 * Current shape selection.
 */
export interface PptxEditorSelectionInfo {
  readonly shapeHandle: SourceHandle;
}

/**
 * Legacy plain-text run view with a required editable source handle.
 */
export interface PptxEditorTextRunInfo {
  readonly text: string;
  readonly handle: SourceHandle;
}

/**
 * Legacy ProseMirror-compatible text body representation.
 *
 * @deprecated Use {@link PptxEditorTextBodyView} and {@link PptxEditorSession.applyAll}.
 */
export interface PptxEditorTextBodyInfo {
  /**
   * @deprecated Use `PptxEditorShapeInfo.textBody` and `applyAll()` with editor commands.
   * This ProseMirror-compatible JSON bridge is retained temporarily for compatibility.
   */
  readonly docJson: PptxTextBodyProseMirrorDocJson;
}

/**
 * Editable text body properties and ordered paragraphs for one shape.
 */
export interface PptxEditorTextBodyView {
  readonly handle?: SourceHandle;
  readonly properties?: SourceTextBodyProperties;
  readonly paragraphs: readonly PptxEditorTextParagraphView[];
}

/**
 * Editable paragraph properties and ordered text runs.
 */
export interface PptxEditorTextParagraphView {
  readonly handle?: SourceHandle;
  readonly properties?: SourceParagraphProperties;
  readonly runs: readonly PptxEditorTextRunView[];
}

/**
 * Text and editable run properties for one run.
 */
export interface PptxEditorTextRunView {
  readonly handle?: SourceHandle;
  readonly properties?: SourceRunProperties;
  readonly text: string;
}

/**
 * Media metadata needed to validate and replace an image.
 */
export interface PptxEditorImageReplacementInfo {
  readonly contentType: string;
  readonly accept: string;
  readonly mediaPartPath: string;
  readonly sharedReferenceCount: number;
}

/**
 * Shape bounds in CSS pixels at 96 DPI.
 */
export interface PptxEditorShapeBoundsPx {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/**
 * Editable capabilities and source handles exposed for one slide shape.
 *
 * Capability flags are present and `true` only when the corresponding edit is supported.
 */
export interface PptxEditorShapeInfo {
  readonly id: string;
  readonly kind: SourceShapeNode["kind"];
  readonly name?: string;
  readonly handle?: SourceHandle;
  readonly bounds?: PptxEditorShapeBoundsPx;
  readonly editableTransform?: boolean;
  readonly editableDelete?: boolean;
  readonly textRuns?: readonly PptxEditorTextRunInfo[];
  readonly textBody?: PptxEditorTextBodyView;
  /**
   * @deprecated Use `textBody` and `applyAll()` with `@pptx-glimpse/editor` commands.
   * This ProseMirror-compatible view is retained temporarily for compatibility.
   */
  readonly editableTextBody?: PptxEditorTextBodyInfo;
  readonly editableImageReplacement?: PptxEditorImageReplacementInfo;
}

export interface PptxEditorSlideSvg extends SlideSvg {
  readonly handle?: SourceHandle;
}

/**
 * Pixel bounds and initial content for {@link PptxEditorSession.addTextBox}.
 */
export interface PptxEditorAddTextBoxOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly text?: string;
  readonly name?: string;
}

/**
 * Pixel bounds and optional name for {@link PptxEditorSession.addConnector}.
 */
export interface PptxEditorAddConnectorOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly name?: string;
}

/**
 * Rendered editor state returned after a successful operation.
 *
 * `warnings` reports successful edits with important side effects, such as replacing a media
 * part shared by multiple shapes. Warnings do not indicate an operation failure.
 */
export interface PptxEditorSlidesResponse {
  readonly slides: readonly PptxEditorSlideSvg[];
  readonly history: PptxEditorHistoryState;
  readonly selection?: PptxEditorSelectionInfo;
  readonly warnings?: readonly EditorCommandWarning[];
}

/**
 * Serialized PPTX bytes and current history state returned by a successful save.
 */
export interface PptxEditorSaveResponse {
  readonly ok: true;
  readonly pptx: Uint8Array;
  readonly history: PptxEditorHistoryState;
}

/**
 * Conversion options applied when the editor renders its complete presentation.
 *
 * Slide selection is managed internally and therefore cannot be supplied.
 */
export type PptxEditorRenderOptions = Omit<ConvertOptions, "slides">;

/**
 * Stable failure codes thrown by the high-level editor.
 */
export type PptxEditorErrorCode =
  | EditorOperationErrorCode
  | "read-failed"
  | "render-failed"
  | "write-failed";

/**
 * Typed high-level failure for expected operation rejections and integration/runtime failures.
 *
 * Use {@link isPptxEditorError} to narrow caught values and branch on {@link code}. Successful
 * commands may still return non-fatal warnings in {@link PptxEditorSlidesResponse.warnings}.
 */
export class PptxEditorError extends Error {
  readonly code: PptxEditorErrorCode;
  declare readonly cause?: unknown;

  constructor(code: PptxEditorErrorCode, message: string, options?: { readonly cause?: unknown }) {
    super(message, options);
    this.name = "PptxEditorError";
    this.code = code;
  }
}

/**
 * Test whether an unknown caught value follows the {@link PptxEditorError} contract.
 *
 * @param value Caught value to inspect.
 * @returns `true` when the value contains a known editor error code.
 */
export function isPptxEditorError(value: unknown): value is PptxEditorError {
  return (
    isRecord(value) &&
    value.name === "PptxEditorError" &&
    typeof value.message === "string" &&
    isPptxEditorErrorCode(value.code)
  );
}

type PptxEditorSvgRenderer = (
  source: PptxSourceModel,
  options?: ConvertOptions,
) => Promise<SvgConversionReport>;
type PptxEditorAffectedSlidesResolver = (
  before: PptxSourceModel,
  after: PptxSourceModel,
) => ReadonlySet<string> | undefined;

let defaultPptxEditorSvgRenderer: PptxEditorSvgRenderer = renderPptxSourceModelToSvg;
let defaultAffectedSlidesResolver: PptxEditorAffectedSlidesResolver = affectedSlidePartPaths;

/** @internal Configures the renderer selected by the active package conditional entry. */
export function configurePptxEditorSessionRenderer(renderer: PptxEditorSvgRenderer): void {
  defaultPptxEditorSvgRenderer = renderer;
}

/** @internal Configures affected-slide resolution for integration tests. */
export function configurePptxEditorSessionAffectedSlidesResolver(
  resolver: PptxEditorAffectedSlidesResolver,
): void {
  defaultAffectedSlidesResolver = resolver;
}

/**
 * Headless read/edit/render/write session for one PPTX presentation.
 *
 * Create sessions with {@link createPptxEditorSession}. Mutating methods update history and
 * rerender affected slides. Expected operation failures throw {@link PptxEditorError}; successful
 * edits can return warnings through {@link PptxEditorSlidesResponse}.
 *
 * Operation rejections before commit leave the document, selection, and history unchanged. A
 * rendering failure after a successful mutation does not roll back the committed document or
 * history; the cached {@link slides} remain unchanged until {@link renderCurrentSlides} succeeds.
 */
export class PptxEditorSession {
  #session: ReturnType<typeof createEditorSession>;
  #slides: readonly PptxEditorSlideSvg[] = [];
  readonly #renderOptions: PptxEditorRenderOptions;
  readonly #renderToSvg: PptxEditorSvgRenderer;
  readonly #resolveAffectedSlides: PptxEditorAffectedSlidesResolver;

  private constructor(source: PptxSourceModel, renderOptions: PptxEditorRenderOptions) {
    this.#session = createEditorSession(source);
    this.#renderOptions = renderOptions;
    this.#renderToSvg = defaultPptxEditorSvgRenderer;
    this.#resolveAffectedSlides = defaultAffectedSlidesResolver;
  }

  /**
   * Parse PPTX bytes, create a session, and render its initial slides.
   *
   * @param input PPTX binary data.
   * @param renderOptions Options reused for editor SVG rendering.
   * @returns An initialized editor session.
   * @throws {@link PptxEditorError} with `read-failed` or `render-failed`.
   */
  static async create(
    input: Uint8Array,
    renderOptions: PptxEditorRenderOptions = {},
  ): Promise<PptxEditorSession> {
    let source: PptxSourceModel;
    try {
      source = readPptx(input);
    } catch (cause) {
      throw integrationError("read-failed", "Failed to read PPTX input", cause);
    }
    const editor = new PptxEditorSession(source, renderOptions);
    await editor.renderCurrentSlides();
    return editor;
  }

  /** Current immutable PPTX source model after all applied history entries. */
  get document(): PptxSourceModel {
    return this.#session.document;
  }

  /** Latest rendered SVG result for every slide. */
  get slides(): readonly PptxEditorSlideSvg[] {
    return this.#slides;
  }

  /** Current undo/redo availability and stack depths. */
  get history(): PptxEditorHistoryState {
    return {
      canUndo: this.#session.canUndo,
      canRedo: this.#session.canRedo,
      undoDepth: this.#session.undoDepth,
      redoDepth: this.#session.redoDepth,
    };
  }

  /** Currently selected shape, or `undefined` when no shape is selected. */
  get selection(): PptxEditorSelectionInfo | undefined {
    return this.#session.selection;
  }

  /**
   * Build a snapshot of the current rendered, history, selection, and warning state.
   *
   * @param warnings Optional warnings emitted by the operation that produced the state.
   * @returns The current public editor response.
   */
  response(warnings?: readonly EditorCommandWarning[]): PptxEditorSlidesResponse {
    return {
      slides: this.#slides,
      history: this.history,
      ...(this.selection !== undefined ? { selection: this.selection } : {}),
      ...(warnings !== undefined && warnings.length > 0 ? { warnings } : {}),
    };
  }

  /**
   * Inspect editable shape capabilities for one slide.
   *
   * @param slideNumber PowerPoint-style 1-based slide number.
   * @returns Shape information, or an empty array when the slide does not exist.
   */
  shapes(slideNumber: number): readonly PptxEditorShapeInfo[] {
    const slide = this.#session.document.slides[slideNumber - 1];
    if (slide === undefined) return [];
    return slide.shapes.flatMap((shape, index) =>
      shapeInfo(this.#session.document, shape, index, true, slide.shapes),
    );
  }

  /**
   * Rerender all current slides using the session render options.
   *
   * @returns The updated SVG slide cache.
   * @throws {@link PptxEditorError} with `render-failed`.
   */
  async renderCurrentSlides(): Promise<readonly PptxEditorSlideSvg[]> {
    return this.#renderSlides();
  }

  async #renderChangedSlides(before: PptxSourceModel): Promise<void> {
    const affectedPartPaths = this.#resolveAffectedSlides(before, this.#session.document);
    await this.#renderSlides(affectedPartPaths);
  }

  async #renderSlides(
    affectedPartPaths?: ReadonlySet<string>,
  ): Promise<readonly PptxEditorSlideSvg[]> {
    const document = this.#session.document;
    const cachedByPartPath = new Map(
      this.#slides.flatMap((slide) =>
        slide.handle?.partPath === undefined ? [] : [[slide.handle.partPath, slide] as const],
      ),
    );
    const canReuseCache =
      affectedPartPaths !== undefined &&
      document.slides.every(
        (slide) =>
          slide.handle?.partPath !== undefined &&
          (affectedPartPaths.has(slide.partPath) || cachedByPartPath.has(slide.partPath)),
      );
    const slideNumbers = canReuseCache
      ? document.slides.flatMap((slide, index) =>
          affectedPartPaths.has(slide.partPath) ? [index + 1] : [],
        )
      : undefined;

    let renderedSlides: SvgConversionReport["slides"] = [];
    if (slideNumbers === undefined || slideNumbers.length > 0) {
      try {
        const report = await this.#renderToSvg(document, {
          textOutput: "text",
          skipSystemFonts: true,
          ...this.#renderOptions,
          ...(slideNumbers !== undefined ? { slides: slideNumbers } : {}),
        });
        renderedSlides = report.slides;
      } catch (cause) {
        throw integrationError("render-failed", "Failed to render editor slides", cause);
      }
    }
    const renderedBySlideNumber = new Map(
      renderedSlides.map((slide) => [slide.slideNumber, slide]),
    );
    this.#slides = document.slides.map((sourceSlide, index) => {
      const slideNumber = index + 1;
      const rendered = renderedBySlideNumber.get(slideNumber);
      const cached = cachedByPartPath.get(sourceSlide.partPath);
      const mayUseCachedSlide =
        canReuseCache &&
        affectedPartPaths !== undefined &&
        !affectedPartPaths.has(sourceSlide.partPath);
      const slide = rendered ?? (mayUseCachedSlide ? cached : undefined);
      if (slide === undefined) {
        throw new Error(
          `PptxEditorSession: renderer did not return requested slide ${String(slideNumber)}`,
        );
      }
      return {
        ...slide,
        slideNumber,
        ...(sourceSlide.handle !== undefined ? { handle: sourceSlide.handle } : {}),
      };
    });
    return this.#slides;
  }

  /**
   * Apply one {@link EditorCommand} as one atomic history entry.
   *
   * @param command Discriminated command containing its `kind` and payload.
   * @returns Updated slides, history, selection, and any non-fatal warnings.
   * @throws {@link PptxEditorError} with `invalid-command` or `render-failed`.
   */
  async apply(command: EditorCommand): Promise<PptxEditorSlidesResponse> {
    return this.applyAll([command]);
  }

  /**
   * Apply commands atomically as one history entry and rerender affected slides.
   *
   * @param commands Ordered editor commands. An empty array returns the current state unchanged.
   * @returns Updated slides, history, selection, and any non-fatal warnings.
   * @throws {@link PptxEditorError} with `invalid-command` or `render-failed`.
   */
  async applyAll(commands: readonly EditorCommand[]): Promise<PptxEditorSlidesResponse> {
    const before = this.#session.document;
    const result = unwrapEditorOperation(this.#session.applyAll(commands));
    if (commands.length === 0) return this.response(result.warnings);
    await this.#renderChangedSlides(before);
    return this.response(result.warnings);
  }

  /**
   * Add and select a text box on one slide.
   *
   * @param slideNumber PowerPoint-style 1-based slide number.
   * @param options Pixel bounds, initial text, and optional shape name.
   * @returns Updated editor state.
   * @throws {@link PptxEditorError} with `invalid-command` or `render-failed`.
   */
  async addTextBox(
    slideNumber = 1,
    options: PptxEditorAddTextBoxOptions = {},
  ): Promise<PptxEditorSlidesResponse> {
    const slide = this.#session.document.slides[slideNumber - 1];
    if (slide?.handle === undefined) {
      throw new PptxEditorError(
        "invalid-command",
        "addTextBox: slide handle was not found in PptxSourceModel source",
      );
    }
    const existingShapeKeys = new Set(slide.shapes.map(shapeSourceKey));
    const before = this.#session.document;
    const result = unwrapEditorOperation(
      this.#session.apply({
        kind: "addTextBox",
        slideHandle: slide.handle,
        offsetX: pxToEmu(options.x ?? DEFAULT_TEXT_BOX_BOUNDS_PX.x),
        offsetY: pxToEmu(options.y ?? DEFAULT_TEXT_BOX_BOUNDS_PX.y),
        width: pxToEmu(options.width ?? DEFAULT_TEXT_BOX_BOUNDS_PX.width),
        height: pxToEmu(options.height ?? DEFAULT_TEXT_BOX_BOUNDS_PX.height),
        text: options.text ?? DEFAULT_TEXT_BOX_TEXT,
        ...(options.name !== undefined ? { name: options.name } : {}),
      }),
    );
    this.#selectNewShape(slideNumber, existingShapeKeys);
    await this.#renderChangedSlides(before);
    return this.response(result.warnings);
  }

  /**
   * Add and select a straight connector on one slide.
   *
   * @param slideNumber PowerPoint-style 1-based slide number.
   * @param options Pixel bounds and optional shape name.
   * @returns Updated editor state.
   * @throws {@link PptxEditorError} with `invalid-command` or `render-failed`.
   */
  async addConnector(
    slideNumber = 1,
    options: PptxEditorAddConnectorOptions = {},
  ): Promise<PptxEditorSlidesResponse> {
    const slide = this.#session.document.slides[slideNumber - 1];
    if (slide?.handle === undefined) {
      throw new PptxEditorError(
        "invalid-command",
        "addConnector: slide handle was not found in PptxSourceModel source",
      );
    }
    const existingShapeKeys = new Set(slide.shapes.map(shapeSourceKey));
    const before = this.#session.document;
    const result = unwrapEditorOperation(
      this.#session.apply({
        kind: "addConnector",
        slideHandle: slide.handle,
        preset: "straightConnector1",
        offsetX: pxToEmu(options.x ?? DEFAULT_CONNECTOR_BOUNDS_PX.x),
        offsetY: pxToEmu(options.y ?? DEFAULT_CONNECTOR_BOUNDS_PX.y),
        width: pxToEmu(options.width ?? DEFAULT_CONNECTOR_BOUNDS_PX.width),
        height: pxToEmu(options.height ?? DEFAULT_CONNECTOR_BOUNDS_PX.height),
        outline: {
          tailEnd: { type: "triangle", width: "med", length: "med" },
        },
        ...(options.name !== undefined ? { name: options.name } : {}),
      }),
    );
    this.#selectNewShape(slideNumber, existingShapeKeys);
    await this.#renderChangedSlides(before);
    return this.response(result.warnings);
  }

  /**
   * Delete a shape by stable source handle.
   *
   * @param handle Handle returned by {@link shapes}.
   * @returns Updated editor state.
   * @throws {@link PptxEditorError} with `invalid-command` or `render-failed`.
   */
  async deleteShape(handle: SourceHandle): Promise<PptxEditorSlidesResponse> {
    const before = this.#session.document;
    const result = unwrapEditorOperation(this.#session.apply({ kind: "deleteShape", handle }));
    await this.#renderChangedSlides(before);
    return this.response(result.warnings);
  }

  /**
   * Delete the currently selected shape.
   *
   * @returns Updated editor state.
   * @throws {@link PptxEditorError} with `invalid-selection`, `invalid-command`, or
   * `render-failed`.
   */
  async deleteSelectedShape(): Promise<PptxEditorSlidesResponse> {
    const selection = this.#session.selection;
    if (selection === undefined) {
      throw new PptxEditorError("invalid-selection", "deleteShape: no selected shape");
    }
    return this.deleteShape(selection.shapeHandle);
  }

  /**
   * @deprecated Use `textBody` from `shapes()` and pass generated commands to `applyAll()`.
   * This ProseMirror-compatible bridge is retained temporarily for compatibility.
   */
  async applyTextBodyDocJson(
    handle: SourceHandle,
    docJson: unknown,
  ): Promise<PptxEditorSlidesResponse> {
    const textBody = this.#requireEditableShapeTextBody(handle);
    const commands = proseMirrorDocJsonToEditorCommands(textBody, docJson);
    if (commands.length === 0) return this.response();

    return this.applyAll(commands);
  }

  /**
   * Select a shape without changing the document or history.
   *
   * @param handle Handle returned by {@link shapes}.
   * @returns Updated editor state.
   * @throws {@link PptxEditorError} with `invalid-selection`.
   */
  selectShape(handle: SourceHandle): PptxEditorSlidesResponse {
    unwrapEditorOperation(this.#session.selectShape(handle));
    return this.response();
  }

  /**
   * Restore the document before the latest history entry.
   *
   * @returns Updated editor state.
   * @throws {@link PptxEditorError} with `empty-undo-stack` or `render-failed`.
   */
  async undo(): Promise<PptxEditorSlidesResponse> {
    const before = this.#session.document;
    unwrapEditorOperation(this.#session.undo());
    await this.#renderChangedSlides(before);
    return this.response();
  }

  /**
   * Reapply the latest undone history entry.
   *
   * @returns Updated editor state.
   * @throws {@link PptxEditorError} with `empty-redo-stack` or `render-failed`.
   */
  async redo(): Promise<PptxEditorSlidesResponse> {
    const before = this.#session.document;
    unwrapEditorOperation(this.#session.redo());
    await this.#renderChangedSlides(before);
    return this.response();
  }

  /**
   * Serialize and validate the current document as PPTX bytes.
   *
   * @returns Serialized bytes and current history state.
   * @throws {@link PptxEditorError} with `write-failed`.
   */
  save(): PptxEditorSaveResponse {
    let output: Uint8Array;
    try {
      output = writePptx(this.#session.document);
      readPptx(output);
    } catch (cause) {
      throw integrationError("write-failed", "Failed to write or validate PPTX output", cause);
    }
    return { ok: true, pptx: output, history: this.history };
  }

  #requireEditableShapeTextBody(handle: SourceHandle): SourceTextBody {
    const shape = findShapeNodeBySourceHandle(this.#session.document, handle);
    if (shape === undefined) {
      throw new PptxEditorError(
        "invalid-command",
        "text body edit: shape handle was not found in PptxSourceModel source",
      );
    }
    if (shape.kind !== "shape" || shape.textBody === undefined) {
      throw new PptxEditorError(
        "invalid-command",
        "text body edit: shape does not have editable text body",
      );
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

/**
 * Parse PPTX bytes and create a rendered high-level editor session.
 *
 * @param input PPTX binary data.
 * @param renderOptions Options reused for SVG rendering after edits.
 * @returns An initialized editor session.
 * @throws {@link PptxEditorError} with `read-failed` or `render-failed`.
 */
export function createPptxEditorSession(
  input: Uint8Array,
  renderOptions?: PptxEditorRenderOptions,
): Promise<PptxEditorSession> {
  return PptxEditorSession.create(input, renderOptions);
}

/**
 * Finds the slide parts whose rendered output may differ between two editor documents.
 *
 * An empty set means that topology-only changes can reuse every cached SVG. `undefined`
 * means that the change cannot be scoped safely and all slides must be rendered.
 *
 * @internal
 */
export function affectedSlidePartPaths(
  before: PptxSourceModel,
  after: PptxSourceModel,
): ReadonlySet<string> | undefined {
  if (before === after) return new Set();
  if (nonSlideRenderingInputsChanged(before, after)) return undefined;

  const beforeByPartPath = new Map(before.slides.map((slide) => [slide.partPath, slide]));
  const affected = new Set<string>();
  for (const slide of after.slides) {
    if (beforeByPartPath.get(slide.partPath) !== slide) affected.add(slide.partPath);
  }

  const beforePartPaths = before.slides.map((slide) => slide.partPath);
  const afterPartPaths = after.slides.map((slide) => slide.partPath);
  const afterPartPathSet: ReadonlySet<string> = new Set(afterPartPaths);
  const topologyChanged =
    beforePartPaths.length !== afterPartPaths.length ||
    beforePartPaths.some((partPath, index) => partPath !== afterPartPaths[index]);

  const changedMediaPartPaths = findChangedMediaPartPaths(before, after);
  let mediaChangeWasScoped = true;
  for (const mediaPartPath of changedMediaPartPaths) {
    const beforeReferences = slidePartPathsReferencingMedia(before, mediaPartPath);
    const afterReferences = slidePartPathsReferencingMedia(after, mediaPartPath);
    if (beforeReferences === undefined || afterReferences === undefined) {
      mediaChangeWasScoped = false;
      continue;
    }
    for (const slidePartPath of [...beforeReferences, ...afterReferences]) {
      if (afterPartPathSet.has(slidePartPath)) affected.add(slidePartPath);
    }
  }

  if (changedMediaPartPaths.size > 0 && !mediaChangeWasScoped) return undefined;
  if (affected.size > 0 || topologyChanged) return affected;
  if (changedMediaPartPaths.size > 0 && mediaChangeWasScoped) return affected;
  return undefined;
}

function nonSlideRenderingInputsChanged(before: PptxSourceModel, after: PptxSourceModel): boolean {
  return (
    !sameReferences(before.slideLayouts, after.slideLayouts) ||
    !sameReferences(before.slideMasters, after.slideMasters) ||
    !sameReferences(before.themes, after.themes) ||
    before.diagnostics !== after.diagnostics ||
    before.presentation.partPath !== after.presentation.partPath ||
    before.presentation.slideSize !== after.presentation.slideSize ||
    before.presentation.defaultTextStyle !== after.presentation.defaultTextStyle ||
    before.presentation.handle !== after.presentation.handle ||
    before.presentation.rawSidecars !== after.presentation.rawSidecars
  );
}

function sameReferences<T>(before: readonly T[], after: readonly T[]): boolean {
  return (
    before === after ||
    (before.length === after.length && before.every((value, index) => value === after[index]))
  );
}

function findChangedMediaPartPaths(
  before: PptxSourceModel,
  after: PptxSourceModel,
): ReadonlySet<string> {
  const beforeMedia = new Map(
    before.packageGraph.media.map((media) => [media.partPath, media.bytes]),
  );
  const afterMedia = new Map(
    after.packageGraph.media.map((media) => [media.partPath, media.bytes]),
  );
  const partPaths = new Set([...beforeMedia.keys(), ...afterMedia.keys()]);
  return new Set(
    [...partPaths].filter(
      (partPath) => !bytesEqual(beforeMedia.get(partPath), afterMedia.get(partPath)),
    ),
  );
}

function bytesEqual(left: Uint8Array | undefined, right: Uint8Array | undefined): boolean {
  if (left === right) return true;
  if (left === undefined || right === undefined || left.length !== right.length) return false;
  return left.every((value, index) => value === right[index]);
}

function slidePartPathsReferencingMedia(
  source: PptxSourceModel,
  mediaPartPath: string,
): ReadonlySet<string> | undefined {
  const directReferences = new Set<string>();
  let inheritedReference = false;
  let unknownReference = false;
  const slidePartPaths = new Set(source.slides.map((slide) => slide.partPath));
  const inheritedPartPaths = new Set([
    ...source.slideLayouts.map((layout) => layout.partPath),
    ...source.slideMasters.map((master) => master.partPath),
  ]);

  for (const relationships of source.packageGraph.relationships) {
    if (
      !relationships.relationships.some(
        (relationship) =>
          relationship.targetMode !== "External" &&
          resolvePartPath(relationships.sourcePartPath, relationship.target) === mediaPartPath,
      )
    ) {
      continue;
    }
    if (slidePartPaths.has(relationships.sourcePartPath)) {
      directReferences.add(relationships.sourcePartPath);
    } else if (inheritedPartPaths.has(relationships.sourcePartPath)) {
      inheritedReference = true;
    } else {
      unknownReference = true;
    }
  }

  if (inheritedReference) return new Set(source.slides.map((slide) => slide.partPath));
  if (unknownReference) return undefined;
  return directReferences;
}

function resolvePartPath(sourcePartPath: string, target: string): string {
  if (target.startsWith("/")) return normalizePartPath(target.slice(1));
  const parent = sourcePartPath.slice(0, Math.max(0, sourcePartPath.lastIndexOf("/") + 1));
  return normalizePartPath(`${parent}${target}`);
}

function normalizePartPath(partPath: string): string {
  const segments: string[] = [];
  for (const segment of partPath.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") {
      segments.pop();
    } else {
      segments.push(segment);
    }
  }
  return segments.join("/");
}

function unwrapEditorOperation<Success extends { readonly ok: true }>(
  result: Success | EditorOperationFailure,
): Success {
  if (!result.ok) {
    throw new PptxEditorError(
      result.code,
      result.message,
      "cause" in result ? { cause: result.cause } : undefined,
    );
  }
  return result;
}

function integrationError(
  code: "read-failed" | "render-failed" | "write-failed",
  message: string,
  cause: unknown,
): PptxEditorError {
  return new PptxEditorError(code, errorMessage(message, cause), { cause });
}

function errorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0
    ? `${message}: ${cause.message}`
    : message;
}

function isRecord(value: unknown): value is { readonly [key: string]: unknown } {
  return typeof value === "object" && value !== null;
}

function isPptxEditorErrorCode(value: unknown): value is PptxEditorErrorCode {
  switch (value) {
    case "invalid-command":
    case "invalid-selection":
    case "empty-undo-stack":
    case "empty-redo-stack":
    case "read-failed":
    case "render-failed":
    case "write-failed":
      return true;
    default:
      return false;
  }
}

function shapeInfo(
  source: PptxSourceModel,
  shape: SourceShapeNode,
  index: number,
  editableTransform = true,
  slideShapes: readonly SourceShapeNode[] = [],
): PptxEditorShapeInfo[] {
  const canEditTransform =
    shape.kind !== "raw" &&
    shape.transform !== undefined &&
    editableTransform &&
    isEditableTransformShape(shape);
  const base: PptxEditorShapeInfo = {
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
          textBody: textBodyView(shape.textBody),
        }
      : {}),
    ...(shape.kind === "shape" && canEditTransform && shape.textBody !== undefined
      ? editableTextBody(shape.textBody)
      : {}),
    ...(shape.kind === "image" ? editableImageReplacement(source, shape) : {}),
  };

  if (shape.kind !== "group") return [base];
  return [
    base,
    ...shape.children.flatMap((child, childIndex) =>
      shapeInfo(source, child, childIndex, false, slideShapes),
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
  if (
    (shape.kind !== "shape" && shape.kind !== "connector") ||
    shape.handle?.nodeId === undefined
  ) {
    return false;
  }
  if (shape.kind === "shape" && isShapeReferencedByConnector(shape, slideShapes)) {
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

function collectTextRuns(runs: readonly SourceTextRun[]): PptxEditorTextRunInfo[] {
  return runs.flatMap((run) => {
    if (run.handle === undefined) return [];
    return [{ text: run.text, handle: run.handle }];
  });
}

function textBodyView(textBody: SourceTextBody): PptxEditorTextBodyView {
  return {
    ...(textBody.handle !== undefined ? { handle: textBody.handle } : {}),
    ...(textBody.properties !== undefined ? { properties: textBody.properties } : {}),
    paragraphs: textBody.paragraphs.map((paragraph) => ({
      ...(paragraph.handle !== undefined ? { handle: paragraph.handle } : {}),
      ...(paragraph.properties !== undefined ? { properties: paragraph.properties } : {}),
      runs: paragraph.runs.map((run) => ({
        text: run.text,
        ...(run.handle !== undefined ? { handle: run.handle } : {}),
        ...(run.properties !== undefined ? { properties: run.properties } : {}),
      })),
    })),
  };
}

function editableTextBody(
  textBody: SourceTextBody,
): Partial<Pick<PptxEditorShapeInfo, "editableTextBody">> {
  try {
    return { editableTextBody: { docJson: textBodyToProseMirrorDocJson(textBody) } };
  } catch {
    return {};
  }
}

function editableImageReplacement(
  source: PptxSourceModel,
  image: SourceImage,
): Partial<Pick<PptxEditorShapeInfo, "editableImageReplacement">> {
  const media = imageMediaPart(source, image);
  if (media === undefined) return {};
  const accept = IMAGE_ACCEPT_BY_CONTENT_TYPE[media.contentType];
  if (accept === undefined) return {};

  return {
    editableImageReplacement: {
      contentType: media.contentType,
      accept,
      mediaPartPath: media.partPath,
      sharedReferenceCount: countImageReferencesToMedia(source, media.partPath),
    },
  };
}

function imageMediaPart(source: PptxSourceModel, image: SourceImage): MediaPart | undefined {
  if (image.blipRelationshipId === undefined || image.handle?.partPath === undefined) {
    return undefined;
  }
  const mediaPartPath = imageMediaPartPath(source, image.handle.partPath, image.blipRelationshipId);
  if (mediaPartPath === undefined) return undefined;
  return source.packageGraph.media.find((part) => part.partPath === mediaPartPath);
}

function imageMediaPartPath(
  source: PptxSourceModel,
  sourcePartPath: PartPath,
  relationshipId: RelationshipId,
): PartPath | undefined {
  const relationships = source.packageGraph.relationships.find(
    (candidate) => candidate.sourcePartPath === sourcePartPath,
  );
  const relationship = relationships?.relationships.find(
    (candidate) => candidate.id === relationshipId && candidate.type === IMAGE_REL_TYPE,
  );
  if (relationship === undefined) return undefined;
  return resolveInternalRelationshipTarget(sourcePartPath, relationship);
}

function countImageReferencesToMedia(source: PptxSourceModel, mediaPartPath: PartPath): number {
  const parsedImageRelationshipKeys = new Set<string>();
  let count = 0;
  for (const slide of source.slides) {
    count += countImageReferencesInTree(
      source,
      slide.partPath,
      slide.shapes,
      mediaPartPath,
      parsedImageRelationshipKeys,
    );
  }
  for (const layout of source.slideLayouts) {
    count += countImageReferencesInTree(
      source,
      layout.partPath,
      layout.shapes,
      mediaPartPath,
      parsedImageRelationshipKeys,
    );
  }
  for (const master of source.slideMasters) {
    count += countImageReferencesInTree(
      source,
      master.partPath,
      master.shapes,
      mediaPartPath,
      parsedImageRelationshipKeys,
    );
  }

  for (const relationships of source.packageGraph.relationships) {
    for (const relationship of relationships.relationships) {
      if (relationship.type !== IMAGE_REL_TYPE) continue;
      if (
        parsedImageRelationshipKeys.has(
          imageRelationshipKey(relationships.sourcePartPath, relationship.id),
        )
      ) {
        continue;
      }
      if (
        resolveInternalRelationshipTarget(relationships.sourcePartPath, relationship) ===
        mediaPartPath
      ) {
        count += 1;
      }
    }
  }
  return count;
}

function countImageReferencesInTree(
  source: PptxSourceModel,
  sourcePartPath: PartPath,
  shapes: readonly SourceShapeNode[],
  mediaPartPath: PartPath,
  parsedImageRelationshipKeys: Set<string>,
): number {
  let count = 0;
  for (const shape of shapes) {
    if (shape.kind === "group") {
      count += countImageReferencesInTree(
        source,
        sourcePartPath,
        shape.children,
        mediaPartPath,
        parsedImageRelationshipKeys,
      );
      continue;
    }
    if (shape.kind !== "image" || shape.blipRelationshipId === undefined) continue;
    if (imageMediaPartPath(source, sourcePartPath, shape.blipRelationshipId) === mediaPartPath) {
      count += 1;
      parsedImageRelationshipKeys.add(
        imageRelationshipKey(sourcePartPath, shape.blipRelationshipId),
      );
    }
  }
  return count;
}

function imageRelationshipKey(partPath: PartPath, relationshipId: RelationshipId): string {
  return `${partPath}\0${relationshipId}`;
}

function resolveInternalRelationshipTarget(
  sourcePartPath: PartPath,
  relationship: Relationship,
): PartPath | undefined {
  if (relationship.targetMode === "External") return undefined;
  return unsafeBrandAssertion<PartPath>(
    normalizePackagePath(
      relationship.target.startsWith("/")
        ? relationship.target.slice(1)
        : joinPackageRelativeTarget(sourcePartPath, relationship.target),
    ),
  );
}

function joinPackageRelativeTarget(sourcePartPath: string, target: string): string {
  const slash = sourcePartPath.lastIndexOf("/");
  const baseDir = slash === -1 ? "" : sourcePartPath.slice(0, slash);
  return baseDir === "" ? target : `${baseDir}/${target}`;
}

function normalizePackagePath(path: string): string {
  const segments: string[] = [];
  for (const segment of path.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") {
      segments.pop();
      continue;
    }
    segments.push(segment);
  }
  return segments.join("/");
}

function transformBoundsPx(transform: {
  readonly offsetX: number;
  readonly offsetY: number;
  readonly width: number;
  readonly height: number;
}): PptxEditorShapeBoundsPx {
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
