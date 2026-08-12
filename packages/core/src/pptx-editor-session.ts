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
  countImageReferencesToMedia,
  createComputedTemplateView,
  type MediaPart,
  type PartPath,
  type PptxComputedView,
  type PptxSourceModel,
  readPptx,
  type Relationship,
  type RelationshipId,
  type SourceBackground,
  type SourceHandle,
  type SourceImage,
  type SourceParagraphProperties,
  type SourceRunProperties,
  type SourceShapeNode,
  type SourceSlideLayout,
  type SourceSlideMaster,
  type SourceTextBody,
  type SourceTextBodyProperties,
  type SourceTextRun,
  type UpdateThemeSchemeInput,
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
  type ConversionDiagnostic,
  type ConvertOptions,
  renderPptxComputedViewToSvg,
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
 * One slide layout in the authoring order of its parent master.
 *
 * `handle` identifies the layout root by its OOXML part path. `hidden` is `true` only when
 * `p:sldLayout@show` is explicitly false; an omitted attribute is visible. The reference count
 * includes slides whose direct layout relationship resolves to this layout.
 */
export interface PptxEditorSlideLayoutCatalogEntry {
  readonly handle: SourceHandle;
  readonly name?: string;
  readonly type?: string;
  readonly hidden: boolean;
  readonly slideReferenceCount: number;
}

/**
 * One slide master and its layouts in presentation authoring order.
 *
 * `handle` identifies the master root by its OOXML part path. `layouts` follows the master's
 * `p:sldLayoutIdLst` order.
 */
export interface PptxEditorSlideMasterCatalogEntry {
  readonly handle: SourceHandle;
  readonly name?: string;
  readonly layouts: readonly PptxEditorSlideLayoutCatalogEntry[];
}

/** Kind of ordered catalog entry rendered by a template preview request. */
export type PptxEditorTemplatePreviewTargetKind = "master" | "layout";

/** Stable expected-failure code returned without throwing by template preview lookup. */
export type PptxEditorTemplatePreviewErrorCode =
  | "preview-handle-not-found"
  | "preview-handle-ambiguous";

/** Successful one-target master/layout SVG preview. */
export interface PptxEditorTemplatePreviewSuccess {
  readonly ok: true;
  readonly targetKind: PptxEditorTemplatePreviewTargetKind;
  readonly handle: SourceHandle;
  readonly svg: string;
  readonly diagnostics: readonly ConversionDiagnostic[];
}

/** Expected failure when a catalog handle cannot identify exactly one template target. */
export interface PptxEditorTemplatePreviewFailure {
  readonly ok: false;
  readonly code: PptxEditorTemplatePreviewErrorCode;
  readonly message: string;
  readonly handle: SourceHandle;
}

/** Result of one non-mutating master/layout preview request. */
export type PptxEditorTemplatePreviewResult =
  | PptxEditorTemplatePreviewSuccess
  | PptxEditorTemplatePreviewFailure;

/**
 * Legacy plain-text run view with a required editable source handle.
 */
export interface PptxEditorTextRunInfo {
  readonly text: string;
  readonly handle: SourceHandle;
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
  readonly editableImageReplacement?: PptxEditorImageReplacementInfo;
}

/** SVG rendering result for one slide in an editor session. */
export interface PptxEditorSlideSvg extends SlideSvg {
  /** Stable source handle for the slide, or `undefined` when no handle is available. */
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
type PptxEditorComputedSvgRenderer = (
  source: PptxSourceModel,
  computed: PptxComputedView,
  options?: ConvertOptions,
) => Promise<SvgConversionReport>;
type PptxEditorAffectedSlidesResolver = (
  before: PptxSourceModel,
  after: PptxSourceModel,
) => ReadonlySet<string> | undefined;

interface PptxEditorSessionDependencies {
  readonly renderToSvg: PptxEditorSvgRenderer;
  readonly renderComputedToSvg: PptxEditorComputedSvgRenderer;
  readonly resolveAffectedSlides: PptxEditorAffectedSlidesResolver;
}

const DEFAULT_PPTX_EDITOR_SESSION_DEPENDENCIES: PptxEditorSessionDependencies = {
  renderToSvg: renderPptxSourceModelToSvg,
  renderComputedToSvg: renderPptxComputedViewToSvg,
  resolveAffectedSlides: affectedSlidePartPaths,
};

let createPptxEditorSessionWithDependencies: (
  input: Uint8Array,
  renderOptions: PptxEditorRenderOptions,
  dependencies: PptxEditorSessionDependencies,
) => Promise<PptxEditorSession>;

/** @internal Reads input and renders a session created by an entry-specific class. */
export async function initializePptxEditorSession<T extends PptxEditorSession>(
  input: Uint8Array,
  createSession: (source: PptxSourceModel) => T,
): Promise<T> {
  let source: PptxSourceModel;
  try {
    source = readPptx(input);
  } catch (cause) {
    throw integrationError("read-failed", "Failed to read PPTX input", cause);
  }
  const editor = createSession(source);
  await editor.renderCurrentSlides();
  return editor;
}

/**
 * Headless read/edit/render/write session for one PPTX presentation.
 *
 * Create sessions with {@link PptxEditorSession.create}. Mutating methods update history and
 * rerender affected slides. Expected operation failures throw {@link PptxEditorError}; successful
 * edits can return warnings through {@link PptxEditorSlidesResponse}.
 *
 * Operation rejections before commit leave the document, selection, and history unchanged. A
 * rendering failure after a successful mutation does not roll back the committed document or
 * history, or a selection updated by that mutation. The failed render leaves cached {@link slides}
 * unchanged; a later successful render updates them, and {@link renderCurrentSlides} can retry
 * explicitly.
 */
export class PptxEditorSession {
  #session: ReturnType<typeof createEditorSession>;
  #slides: readonly PptxEditorSlideSvg[] = [];
  readonly #renderOptions: PptxEditorRenderOptions;
  readonly #renderToSvg: PptxEditorSvgRenderer;
  readonly #renderComputedToSvg: PptxEditorComputedSvgRenderer;
  readonly #resolveAffectedSlides: PptxEditorAffectedSlidesResolver;

  static {
    createPptxEditorSessionWithDependencies = async (input, renderOptions, dependencies) => {
      return initializePptxEditorSession(
        input,
        (source) => new PptxEditorSession(source, renderOptions, dependencies),
      );
    };
  }

  protected constructor(
    source: PptxSourceModel,
    renderOptions: PptxEditorRenderOptions,
    dependencies: PptxEditorSessionDependencies,
  ) {
    this.#session = createEditorSession(source);
    this.#renderOptions = renderOptions;
    this.#renderToSvg = dependencies.renderToSvg;
    this.#renderComputedToSvg = dependencies.renderComputedToSvg;
    this.#resolveAffectedSlides = dependencies.resolveAffectedSlides;
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
    return createPptxEditorSessionWithDependencies(
      input,
      renderOptions,
      DEFAULT_PPTX_EDITOR_SESSION_DEPENDENCIES,
    );
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
   * Slide masters and their layouts in PowerPoint authoring order.
   *
   * The outer array follows `p:sldMasterIdLst`; each nested array follows the corresponding
   * `p:sldLayoutIdLst`. Only resolved catalog entries are returned. Inspect
   * {@link document}.diagnostics for unresolved relationships.
   */
  get layoutCatalog(): readonly PptxEditorSlideMasterCatalogEntry[] {
    return buildLayoutCatalog(this.#session.document);
  }

  /**
   * Render exactly one ordered-catalog master or layout without modifying editor state.
   *
   * Missing and ambiguous handles are returned as stable failures. Unsupported render content is
   * skipped and reported through stable conversion diagnostics. Runtime renderer failures use the
   * existing {@link PptxEditorError} `render-failed` integration contract.
   */
  async previewLayoutCatalogTarget(handle: SourceHandle): Promise<PptxEditorTemplatePreviewResult> {
    const matches = findTemplatePreviewTargets(this.#session.document, handle);
    if (matches.length === 0) {
      return {
        ok: false,
        code: "preview-handle-not-found",
        message: `No slide master or layout matches handle '${formatSourceHandle(handle)}'.`,
        handle,
      };
    }
    if (matches.length > 1) {
      return {
        ok: false,
        code: "preview-handle-ambiguous",
        message: `Multiple slide masters or layouts match handle '${formatSourceHandle(handle)}'.`,
        handle,
      };
    }
    const target = matches[0];
    if (target === undefined) throw new Error("template preview target invariant failed");
    const computed = createComputedTemplateView(this.#session.document, {
      kind: target.kind === "master" ? "slideMaster" : "slideLayout",
      partPath: target.handle.partPath,
    });
    let report: SvgConversionReport;
    try {
      report = await this.#renderComputedToSvg(this.#session.document, computed, {
        textOutput: "text",
        skipSystemFonts: true,
        ...this.#renderOptions,
      });
    } catch (cause) {
      throw integrationError("render-failed", "Failed to render master or layout preview", cause);
    }
    const rendered = report.slides[0];
    if (rendered === undefined) {
      throw new Error("PptxEditorSession: renderer did not return template preview target");
    }
    return {
      ok: true,
      targetKind: target.kind,
      handle: target.handle,
      svg: rendered.svg,
      diagnostics: report.diagnostics
        .filter((diagnostic) =>
          templatePreviewDiagnosticApplies(diagnostic, this.#session.document, target),
        )
        .map(withoutSyntheticSlideNumber),
    };
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
   * Delete a supported root or nested drawing by stable source handle.
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
   * Group consecutive sibling drawings and select the new native group.
   *
   * @param shapeHandles Stable handles in any order; document order determines group child order.
   * @returns Updated editor state with the new group selected.
   * @throws {@link PptxEditorError} with `invalid-command` or `render-failed`.
   */
  async groupShapes(shapeHandles: readonly SourceHandle[]): Promise<PptxEditorSlidesResponse> {
    const before = this.#session.document;
    const result = unwrapEditorOperation(
      this.#session.apply({ kind: "groupShapes", shapeHandles }),
    );
    await this.#renderChangedSlides(before);
    return this.response(result.warnings);
  }

  /**
   * Move consecutive sibling drawings to a root/native-group destination in the same part.
   * Identity-mapped ancestor chains preserve each moved drawing's rendered transform.
   */
  async moveShapes(
    shapeHandles: readonly SourceHandle[],
    destinationHandle: SourceHandle,
    options: { readonly beforeShapeHandle?: SourceHandle } = {},
  ): Promise<PptxEditorSlidesResponse> {
    const before = this.#session.document;
    const result = unwrapEditorOperation(
      this.#session.apply({
        kind: "moveShapes",
        shapeHandles,
        destinationHandle,
        ...(options.beforeShapeHandle !== undefined
          ? { beforeShapeHandle: options.beforeShapeHandle }
          : {}),
      }),
    );
    await this.#renderChangedSlides(before);
    return this.response(result.warnings);
  }

  /** Move consecutive non-placeholder shape/picture/connector roots to another slide. */
  async moveShapesAcrossSlides(
    shapeHandles: readonly SourceHandle[],
    destinationSlideHandle: SourceHandle,
    options: { readonly beforeShapeHandle?: SourceHandle } = {},
  ): Promise<PptxEditorSlidesResponse> {
    const before = this.#session.document;
    const result = unwrapEditorOperation(
      this.#session.apply({
        kind: "moveShapesAcrossSlides",
        shapeHandles,
        destinationSlideHandle,
        ...(options.beforeShapeHandle !== undefined
          ? { beforeShapeHandle: options.beforeShapeHandle }
          : {}),
      }),
    );
    await this.#renderChangedSlides(before);
    return this.response(result.warnings);
  }

  /**
   * Expand one losslessly ungroupable native group and select its first child in document order.
   *
   * @param groupHandle Stable handle of the native group.
   * @returns Updated editor state with the first expanded child selected.
   * @throws {@link PptxEditorError} with `invalid-command` or `render-failed`.
   */
  async ungroupShape(groupHandle: SourceHandle): Promise<PptxEditorSlidesResponse> {
    const before = this.#session.document;
    const result = unwrapEditorOperation(
      this.#session.apply({ kind: "ungroupShape", groupHandle }),
    );
    await this.#renderChangedSlides(before);
    return this.response(result.warnings);
  }

  /** Update selected color/font fields on one existing theme and rerender its master slides. */
  async updateThemeScheme(
    themeHandle: SourceHandle,
    input: UpdateThemeSchemeInput,
  ): Promise<PptxEditorSlidesResponse> {
    return this.apply({
      kind: "updateThemeScheme",
      handle: themeHandle,
      colorScheme: input.colorScheme,
      fontScheme: input.fontScheme,
    });
  }

  /**
   * Delete the currently selected supported root or nested drawing.
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

  #selectNewShape(slideNumber: number, existingShapeKeys: ReadonlySet<string>): void {
    const slide = this.#session.document.slides[slideNumber - 1];
    const addedShape = slide?.shapes.find((shape) => !existingShapeKeys.has(shapeSourceKey(shape)));
    if (addedShape?.handle === undefined) return;
    this.#session.selectShape(addedShape.handle);
  }
}

function buildLayoutCatalog(source: PptxSourceModel): readonly PptxEditorSlideMasterCatalogEntry[] {
  const mastersByPartPath = new Map(
    source.slideMasters.map((master) => [master.partPath, master] as const),
  );
  const layoutsByPartPath = new Map(
    source.slideLayouts.map((layout) => [layout.partPath, layout] as const),
  );
  const slideReferencesByLayoutPartPath = new Map<PartPath, number>();
  for (const slide of source.slides) {
    slideReferencesByLayoutPartPath.set(
      slide.layoutPartPath,
      (slideReferencesByLayoutPartPath.get(slide.layoutPartPath) ?? 0) + 1,
    );
  }

  return source.presentation.slideMasterPartPaths.flatMap((masterPartPath) => {
    const master = mastersByPartPath.get(masterPartPath);
    if (master === undefined) return [];
    return [
      {
        handle: master.handle ?? { partPath: master.partPath },
        ...(master.name !== undefined ? { name: master.name } : {}),
        layouts: master.layoutPartPaths.flatMap((layoutPartPath) => {
          const layout = layoutsByPartPath.get(layoutPartPath);
          if (layout === undefined) return [];
          return [
            {
              handle: layout.handle ?? { partPath: layout.partPath },
              ...(layout.name !== undefined ? { name: layout.name } : {}),
              ...(layout.type !== undefined ? { type: layout.type } : {}),
              hidden: layout.show === false,
              slideReferenceCount: slideReferencesByLayoutPartPath.get(layout.partPath) ?? 0,
            },
          ];
        }),
      },
    ];
  });
}

interface TemplatePreviewTarget {
  readonly kind: PptxEditorTemplatePreviewTargetKind;
  readonly handle: SourceHandle;
}

function findTemplatePreviewTargets(
  source: PptxSourceModel,
  handle: SourceHandle,
): readonly TemplatePreviewTarget[] {
  return buildLayoutCatalog(source).flatMap((master) => [
    ...(sourceHandlesMatch(master.handle, handle)
      ? [{ kind: "master" as const, handle: master.handle }]
      : []),
    ...master.layouts.flatMap((layout) =>
      sourceHandlesMatch(layout.handle, handle)
        ? [{ kind: "layout" as const, handle: layout.handle }]
        : [],
    ),
  ]);
}

function sourceHandlesMatch(left: SourceHandle, right: SourceHandle): boolean {
  return (
    left.partPath === right.partPath &&
    left.nodeId === right.nodeId &&
    left.relationshipId === right.relationshipId &&
    left.orderingSlot === right.orderingSlot &&
    arraysEqual(left.rawSidecarIds, right.rawSidecarIds)
  );
}

function arraysEqual<T>(left: readonly T[] | undefined, right: readonly T[] | undefined): boolean {
  if (left === undefined || right === undefined) return left === right;
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

function formatSourceHandle(handle: SourceHandle): string {
  return handle.nodeId === undefined
    ? handle.partPath
    : `${handle.partPath}#${String(handle.nodeId)}`;
}

function withoutSyntheticSlideNumber(diagnostic: ConversionDiagnostic): ConversionDiagnostic {
  const { slideNumber, ...previewDiagnostic } = diagnostic;
  void slideNumber;
  return previewDiagnostic;
}

function templatePreviewDiagnosticApplies(
  diagnostic: ConversionDiagnostic,
  source: PptxSourceModel,
  target: TemplatePreviewTarget,
): boolean {
  if (diagnostic.source !== "document" || diagnostic.sourcePartPath === undefined) return true;
  return templatePreviewDependencyPartPaths(source, target).has(diagnostic.sourcePartPath);
}

function templatePreviewDependencyPartPaths(
  source: PptxSourceModel,
  target: TemplatePreviewTarget,
): ReadonlySet<string> {
  const partPaths = new Set<string>([target.handle.partPath]);
  let master: SourceSlideMaster | undefined;
  if (target.kind === "master") {
    master = source.slideMasters.find((candidate) => candidate.partPath === target.handle.partPath);
  } else {
    const layout = source.slideLayouts.find(
      (candidate) => candidate.partPath === target.handle.partPath,
    );
    if (layout !== undefined) partPaths.add(layout.masterPartPath);
    master = source.slideMasters.find((candidate) => candidate.partPath === layout?.masterPartPath);
  }
  if (master?.themePartPath !== undefined) partPaths.add(master.themePartPath);
  return partPaths;
}

/**
 * Creates an entry-specific editor-session factory with immutable internal dependencies.
 *
 * @internal
 */
export function createPptxEditorSessionFactory(
  renderToSvg: PptxEditorSvgRenderer,
  resolveAffectedSlides: PptxEditorAffectedSlidesResolver = affectedSlidePartPaths,
  renderComputedToSvg: PptxEditorComputedSvgRenderer = renderPptxComputedViewToSvg,
): (input: Uint8Array, renderOptions?: PptxEditorRenderOptions) => Promise<PptxEditorSession> {
  const dependencies: PptxEditorSessionDependencies = {
    renderToSvg,
    renderComputedToSvg,
    resolveAffectedSlides,
  };
  return (input, renderOptions = {}) =>
    createPptxEditorSessionWithDependencies(input, renderOptions, dependencies);
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
  const changedThemePartPaths = findChangedThemePartPaths(before, after);
  if (changedThemePartPaths === undefined) return undefined;
  if (nonDrawingRenderingInputsChanged(before, after)) return undefined;

  const inheritedDrawingChanges = changedInheritedDrawingPartPaths(before, after);
  if (inheritedDrawingChanges === undefined) return undefined;

  const beforeByPartPath = new Map(before.slides.map((slide) => [slide.partPath, slide]));
  const affected = new Set<string>();
  for (const slide of after.slides) {
    if (beforeByPartPath.get(slide.partPath) !== slide) affected.add(slide.partPath);
    if (inheritedDrawingChanges.layoutPartPaths.has(slide.layoutPartPath)) {
      affected.add(slide.partPath);
      continue;
    }
    const layout = after.slideLayouts.find(
      (candidate) => candidate.partPath === slide.layoutPartPath,
    );
    if (
      layout !== undefined &&
      inheritedDrawingChanges.masterPartPaths.has(layout.masterPartPath)
    ) {
      affected.add(slide.partPath);
    }
    const master = after.slideMasters.find(
      (candidate) => candidate.partPath === layout?.masterPartPath,
    );
    if (master?.themePartPath !== undefined && changedThemePartPaths.has(master.themePartPath)) {
      affected.add(slide.partPath);
    }
  }

  const beforePartPaths = before.slides.map((slide) => slide.partPath);
  const afterPartPaths = after.slides.map((slide) => slide.partPath);
  const afterPartPathSet: ReadonlySet<string> = new Set(afterPartPaths);

  const beforeLayouts = new Map(before.slideLayouts.map((layout) => [layout.partPath, layout]));
  const afterLayouts = new Map(after.slideLayouts.map((layout) => [layout.partPath, layout]));
  const changedLayoutPartPaths = changedHierarchyPartPaths(beforeLayouts, afterLayouts);
  for (const slide of after.slides) {
    if (!changedLayoutPartPaths.has(slide.layoutPartPath)) continue;
    const previous = beforeLayouts.get(slide.layoutPartPath);
    const current = afterLayouts.get(slide.layoutPartPath);
    const backgroundOnly =
      previous !== undefined &&
      current !== undefined &&
      onlyLayoutBackgroundChanged(previous, current);
    if (!backgroundOnly || slide.background === undefined) {
      affected.add(slide.partPath);
    }
  }

  const beforeMasters = new Map(before.slideMasters.map((master) => [master.partPath, master]));
  const afterMasters = new Map(after.slideMasters.map((master) => [master.partPath, master]));
  const changedMasterPartPaths = changedHierarchyPartPaths(beforeMasters, afterMasters);
  for (const slide of after.slides) {
    const layout = afterLayouts.get(slide.layoutPartPath);
    if (layout === undefined || !changedMasterPartPaths.has(layout.masterPartPath)) continue;
    const previous = beforeMasters.get(layout.masterPartPath);
    const current = afterMasters.get(layout.masterPartPath);
    const backgroundOnly =
      previous !== undefined &&
      current !== undefined &&
      onlyMasterBackgroundChanged(previous, current);
    if (!backgroundOnly || (slide.background === undefined && layout.background === undefined)) {
      affected.add(slide.partPath);
    }
  }
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
  if (changedThemePartPaths.size > 0) return affected;
  return undefined;
}

function nonDrawingRenderingInputsChanged(
  before: PptxSourceModel,
  after: PptxSourceModel,
): boolean {
  return (
    before.diagnostics !== after.diagnostics ||
    before.presentation.partPath !== after.presentation.partPath ||
    before.presentation.slideSize !== after.presentation.slideSize ||
    before.presentation.defaultTextStyle !== after.presentation.defaultTextStyle ||
    before.presentation.handle !== after.presentation.handle ||
    before.presentation.rawSidecars !== after.presentation.rawSidecars
  );
}

function findChangedThemePartPaths(
  before: PptxSourceModel,
  after: PptxSourceModel,
): ReadonlySet<string> | undefined {
  if (before.themes.length !== after.themes.length) return undefined;
  const beforeByPartPath = new Map(before.themes.map((theme) => [theme.partPath, theme]));
  const changed = new Set<string>();
  for (const theme of after.themes) {
    const previous = beforeByPartPath.get(theme.partPath);
    if (previous === undefined) return undefined;
    if (previous !== theme) changed.add(theme.partPath);
  }
  return changed;
}

function changedHierarchyPartPaths<T>(
  before: ReadonlyMap<string, T>,
  after: ReadonlyMap<string, T>,
): ReadonlySet<string> {
  const partPaths = new Set([...before.keys(), ...after.keys()]);
  return new Set([...partPaths].filter((partPath) => before.get(partPath) !== after.get(partPath)));
}

function onlyLayoutBackgroundChanged(before: SourceSlideLayout, after: SourceSlideLayout): boolean {
  return (
    before.background !== after.background &&
    before.masterPartPath === after.masterPartPath &&
    before.colorMapOverride === after.colorMapOverride &&
    before.showMasterShapes === after.showMasterShapes &&
    before.shapes === after.shapes &&
    before.rawSidecars === after.rawSidecars
  );
}

function onlyMasterBackgroundChanged(before: SourceSlideMaster, after: SourceSlideMaster): boolean {
  return (
    before.background !== after.background &&
    before.themePartPath === after.themePartPath &&
    before.colorMap === after.colorMap &&
    before.txStyles === after.txStyles &&
    before.shapes === after.shapes &&
    before.rawSidecars === after.rawSidecars
  );
}

interface ChangedInheritedDrawingPartPaths {
  readonly layoutPartPaths: ReadonlySet<string>;
  readonly masterPartPaths: ReadonlySet<string>;
}

function changedInheritedDrawingPartPaths(
  before: PptxSourceModel,
  after: PptxSourceModel,
): ChangedInheritedDrawingPartPaths | undefined {
  const layoutPartPaths = changedShapeOnlyPartPaths(before.slideLayouts, after.slideLayouts);
  const masterPartPaths = changedShapeOnlyPartPaths(before.slideMasters, after.slideMasters);
  if (layoutPartPaths === undefined || masterPartPaths === undefined) return undefined;
  return { layoutPartPaths, masterPartPaths };
}

function changedShapeOnlyPartPaths<
  T extends { readonly partPath: string; readonly shapes: readonly unknown[] },
>(before: readonly T[], after: readonly T[]): ReadonlySet<string> | undefined {
  if (before.length !== after.length) return undefined;
  const beforeByPartPath = new Map(before.map((part) => [part.partPath, part]));
  const changed = new Set<string>();
  for (const part of after) {
    const previous = beforeByPartPath.get(part.partPath);
    if (previous === undefined) return undefined;
    if (previous === part) continue;
    if (!sameReferencesExceptShapesAndBackground(previous, part)) return undefined;
    if (previous.shapes !== part.shapes) changed.add(part.partPath);
  }
  return changed;
}

function sameReferencesExceptShapesAndBackground(before: object, after: object): boolean {
  const keys = new Set([...Object.keys(before), ...Object.keys(after)]);
  for (const key of keys) {
    if (key === "shapes" || key === "background") continue;
    if (Reflect.get(before, key) !== Reflect.get(after, key)) return false;
  }
  return true;
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
  let unknownReference = false;
  const slidePartPaths = new Set(source.slides.map((slide) => slide.partPath));
  const layouts = new Map(source.slideLayouts.map((layout) => [layout.partPath, layout]));
  const masters = new Map(source.slideMasters.map((master) => [master.partPath, master]));

  for (const relationships of source.packageGraph.relationships) {
    const relationshipIds = new Set(
      relationships.relationships.flatMap((relationship) =>
        relationship.targetMode !== "External" &&
        resolvePartPath(relationships.sourcePartPath, relationship.target) === mediaPartPath
          ? [relationship.id]
          : [],
      ),
    );
    if (relationshipIds.size === 0) continue;
    if (slidePartPaths.has(relationships.sourcePartPath)) {
      directReferences.add(relationships.sourcePartPath);
      continue;
    }
    const layout = layouts.get(relationships.sourcePartPath);
    if (layout !== undefined) {
      const backgroundOnly =
        backgroundUsesRelationship(layout.background, relationshipIds) &&
        backgroundEditOwnsRelationship(source, layout.partPath, mediaPartPath, relationshipIds);
      for (const slide of source.slides) {
        if (
          slide.layoutPartPath === layout.partPath &&
          (!backgroundOnly || slide.background === undefined)
        ) {
          directReferences.add(slide.partPath);
        }
      }
      continue;
    }
    const master = masters.get(relationships.sourcePartPath);
    if (master !== undefined) {
      const backgroundOnly =
        backgroundUsesRelationship(master.background, relationshipIds) &&
        backgroundEditOwnsRelationship(source, master.partPath, mediaPartPath, relationshipIds);
      for (const slide of source.slides) {
        const slideLayout = layouts.get(slide.layoutPartPath);
        if (
          slideLayout?.masterPartPath === master.partPath &&
          (!backgroundOnly ||
            (slide.background === undefined && slideLayout.background === undefined))
        ) {
          directReferences.add(slide.partPath);
        }
      }
      continue;
    }
    unknownReference = true;
  }

  if (unknownReference) return undefined;
  return directReferences;
}

/**
 * Newly authored background image relationships are target-local and cannot also be used by a
 * shape. Existing package relationships may be shared with arbitrary drawing content, so they
 * must conservatively invalidate the whole layout/master hierarchy.
 */
function backgroundEditOwnsRelationship(
  source: PptxSourceModel,
  ownerPartPath: string,
  mediaPartPath: string,
  relationshipIds: ReadonlySet<string>,
): boolean {
  return (source.edits ?? []).some((edit) => {
    if (edit.kind === "setBackground") {
      return (
        edit.targetPartPath === ownerPartPath &&
        edit.mediaPartPath === mediaPartPath &&
        edit.relationshipId !== undefined &&
        relationshipIds.has(edit.relationshipId)
      );
    }
    return (
      edit.kind === "setSlideBackground" &&
      edit.slidePartPath === ownerPartPath &&
      edit.mediaPartPath === mediaPartPath &&
      edit.relationshipId !== undefined &&
      relationshipIds.has(edit.relationshipId)
    );
  });
}

function backgroundUsesRelationship(
  background: SourceBackground | undefined,
  relationshipIds: ReadonlySet<string>,
): boolean {
  return (
    background?.kind === "fill" &&
    background.fill.kind === "image" &&
    background.fill.blipRelationshipId !== undefined &&
    relationshipIds.has(background.fill.blipRelationshipId)
  );
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
    ...(isDeletableShape(shape, slideShapes) ? { editableDelete: true } : {}),
    ...("textBody" in shape && shape.textBody !== undefined
      ? {
          textRuns: collectTextRuns(
            shape.textBody.paragraphs.flatMap((paragraph) => paragraph.runs),
          ),
          textBody: textBodyView(shape.textBody),
        }
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
    (shape.kind !== "shape" &&
      shape.kind !== "connector" &&
      shape.kind !== "image" &&
      shape.kind !== "table" &&
      shape.kind !== "chart" &&
      shape.kind !== "group") ||
    shape.handle?.nodeId === undefined
  ) {
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
  const deletedIds = new Set(
    flattenShapeTree([shape]).flatMap((candidate) =>
      candidate.nodeId === undefined ? [] : [String(candidate.nodeId)],
    ),
  );
  return flattenShapeTree(slideShapes).some(
    (candidate) =>
      candidate.kind === "connector" &&
      !deletedIds.has(String(candidate.nodeId)) &&
      (deletedIds.has(String(candidate.connection?.start?.shapeId)) ||
        deletedIds.has(String(candidate.connection?.end?.shapeId))),
  );
}

function flattenShapeTree(shapes: readonly SourceShapeNode[]): SourceShapeNode[] {
  return shapes.flatMap((shape) => [
    shape,
    ...(shape.kind === "group" ? flattenShapeTree(shape.children) : []),
  ]);
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
