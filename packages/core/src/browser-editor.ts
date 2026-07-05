import {
  findShapeNodeBySourceHandle,
  type MediaPart,
  type PartPath,
  type PptxSourceModel,
  readPptx,
  type Relationship,
  type RelationshipId,
  type SourceHandle,
  type SourceImage,
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
import { unsafeBrandAssertion } from "./unsafe-type-assertion.js";

const EMU_PER_INCH = 914400;
const DEFAULT_DPI = 96;
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const IMAGE_ACCEPT_BY_CONTENT_TYPE: Readonly<Record<string, string>> = {
  "image/png": "image/png,.png",
  "image/jpeg": "image/jpeg,.jpg,.jpeg",
  "image/gif": "image/gif,.gif",
  "image/bmp": "image/bmp,.bmp",
  "image/tiff": "image/tiff,.tif,.tiff",
  "image/webp": "image/webp,.webp",
};

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

export interface BrowserEditorImageReplacementInfo {
  readonly contentType: string;
  readonly accept: string;
  readonly mediaPartPath: string;
  readonly sharedReferenceCount: number;
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
  readonly textRuns?: readonly BrowserEditorTextRunInfo[];
  readonly editableTextBody?: BrowserEditorTextBodyInfo;
  readonly editableImageReplacement?: BrowserEditorImageReplacementInfo;
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
    return slide.shapes.flatMap((shape, index) => shapeInfo(this.#session.document, shape, index));
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
}

export function createBrowserPptxEditorSession(
  input: Uint8Array,
  renderOptions?: BrowserEditorRenderOptions,
): Promise<BrowserPptxEditorSession> {
  return BrowserPptxEditorSession.create(input, renderOptions);
}

function shapeInfo(
  source: PptxSourceModel,
  shape: SourceShapeNode,
  index: number,
  editableTransform = true,
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
    ...(shape.kind === "image" ? editableImageReplacement(source, shape) : {}),
  };

  if (shape.kind !== "group") return [base];
  return [
    base,
    ...shape.children.flatMap((child, childIndex) => shapeInfo(source, child, childIndex, false)),
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

function editableImageReplacement(
  source: PptxSourceModel,
  image: SourceImage,
): Partial<Pick<BrowserEditorShapeInfo, "editableImageReplacement">> {
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
