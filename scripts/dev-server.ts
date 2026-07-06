import { execFile } from "child_process";
import { watch } from "fs";
import { lstat, mkdtemp, readFile, realpath, rm, writeFile } from "fs/promises";
import { createServer, type IncomingMessage, type ServerResponse } from "http";
import { tmpdir } from "os";
import { basename, dirname, extname, isAbsolute, join, relative, resolve } from "path";
import { pathToFileURL } from "url";
import { promisify } from "util";
import { type WebSocket, WebSocketServer } from "ws";

import {
  asEmu,
  asPartPath,
  asPt,
  asRelationshipId,
  asSourceNodeId,
  type ConnectorPresetGeometry,
  type EditableTextRunProperties,
  type EditableTextRunProperty,
  findShapeNodeBySourceHandle,
  type MediaPart,
  type PartPath,
  readPptx,
  type Relationship,
  type RelationshipId,
  type SourceArrowEndpoint,
  type SourceHandle,
  type SourceImage,
  type SourceShapeNode,
  type SourceTextBody,
  type SourceTextRun,
  writePptx,
} from "../packages/document/src/index.js";
import {
  createEditorSession,
  type EditorCommand,
  type EditorCommandWarning,
  type PptxTextBodyProseMirrorDocJson,
  proseMirrorDocJsonToEditorCommands,
  textBodyToProseMirrorDocJson,
} from "../packages/editor-core/src/index.js";
import { generateDevEditorHtml } from "./dev-server-editor/template.js";
import { unsafeScriptInputAssertion } from "./unsafe-type-assertion.js";

const DEFAULT_PORT = 3000;
const DEBOUNCE_MS = 300;
// The in-process editor session imports document/editor-core once; restart the server for those changes.
const WATCH_DIRS = [resolve("packages/core/src"), resolve("packages/renderer/src")];
const RENDER_TIMEOUT_MS = 30_000;
const MAX_BUFFER = 50 * 1024 * 1024;
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
const MAX_JSON_BODY_BYTES = 20 * 1024 * 1024;
const MAX_IMAGE_REPLACEMENT_BYTES = 5 * 1024 * 1024;

const IMAGE_ACCEPT_BY_CONTENT_TYPE: Readonly<Record<string, string>> = {
  "image/png": "image/png,.png",
  "image/jpeg": "image/jpeg,.jpg,.jpeg",
  "image/gif": "image/gif,.gif",
  "image/bmp": "image/bmp,.bmp",
  "image/tiff": "image/tiff,.tif,.tiff",
  "image/webp": "image/webp,.webp",
};

const execFileAsync = promisify(execFile);

interface SlideSvg {
  slideNumber: number;
  svg: string;
  handle?: SourceHandle;
}

interface EditorHistoryState {
  canUndo: boolean;
  canRedo: boolean;
  undoDepth: number;
  redoDepth: number;
}

interface ShapeBoundsPx {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface EditorTextRunInfo {
  text: string;
  handle: SourceHandle;
}

interface EditorTextBodyInfo {
  docJson: PptxTextBodyProseMirrorDocJson;
}

interface EditorImageReplacementInfo {
  contentType: string;
  accept: string;
  mediaPartPath: string;
  sharedReferenceCount: number;
}

interface EditorShapeInfo {
  id: string;
  kind: SourceShapeNode["kind"];
  name?: string;
  handle?: SourceHandle;
  bounds?: ShapeBoundsPx;
  editableTransform?: boolean;
  editableDelete?: boolean;
  textRuns?: EditorTextRunInfo[];
  editableTextBody?: EditorTextBodyInfo;
  editableImageReplacement?: EditorImageReplacementInfo;
}

interface EditorSlidesResponse {
  slides: SlideSvg[];
  history: EditorHistoryState;
  selection?: EditorSelectionInfo;
  warnings?: readonly EditorCommandWarning[];
}

interface EditorSaveResponse {
  ok: true;
  path: string;
  history: EditorHistoryState;
}

interface EditorSelectionInfo {
  shapeHandle: SourceHandle;
}

type RenderPptxBytes = (input: Uint8Array) => Promise<SlideSvg[]>;

// --- Rendering via child process ---

async function renderSlides(pptxPath: string): Promise<SlideSvg[]> {
  const workerPath = resolve("scripts/dev-server-render.ts");
  const { stdout } = await execFileAsync("npx", ["tsx", workerPath, pptxPath], {
    maxBuffer: MAX_BUFFER,
    timeout: RENDER_TIMEOUT_MS,
  });
  return unsafeScriptInputAssertion<SlideSvg[]>(JSON.parse(stdout));
}

async function renderPptxBytes(input: Uint8Array): Promise<SlideSvg[]> {
  const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-dev-"));
  const pptxPath = join(dir, "document.pptx");
  try {
    await writeFile(pptxPath, input);
    return await renderSlides(pptxPath);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

export class DevEditorBackend {
  #session: ReturnType<typeof createEditorSession>;
  #slides: SlideSvg[] = [];
  #dirty = false;
  #mutationQueue: Promise<void> = Promise.resolve();
  readonly #render: RenderPptxBytes;

  private constructor(
    readonly sourcePath: string,
    source: ReturnType<typeof readPptx>,
    render: RenderPptxBytes,
  ) {
    this.#session = createEditorSession(source);
    this.#render = render;
  }

  static async load(
    pptxPath: string,
    render: RenderPptxBytes = renderPptxBytes,
  ): Promise<DevEditorBackend> {
    const input = await readFile(pptxPath);
    const backend = new DevEditorBackend(pptxPath, readPptx(input), render);
    await backend.renderSlidesFromBytes(input);
    return backend;
  }

  get slides(): readonly SlideSvg[] {
    return this.#slides;
  }

  get history(): EditorHistoryState {
    return {
      canUndo: this.#session.canUndo,
      canRedo: this.#session.canRedo,
      undoDepth: this.#session.undoDepth,
      redoDepth: this.#session.redoDepth,
    };
  }

  get selection(): EditorSelectionInfo | undefined {
    return this.#session.selection;
  }

  response(warnings?: readonly EditorCommandWarning[]): EditorSlidesResponse {
    return {
      slides: this.#slides,
      history: this.history,
      ...(this.selection !== undefined ? { selection: this.selection } : {}),
      ...(warnings !== undefined && warnings.length > 0 ? { warnings } : {}),
    };
  }

  shapes(slideNumber: number): EditorShapeInfo[] {
    const slide = this.#session.document.slides[slideNumber - 1];
    if (slide === undefined) return [];
    return slide.shapes.flatMap((shape, index) =>
      shapeInfo(this.#session.document, shape, index, true, slide.shapes),
    );
  }

  async renderCurrentSlides(): Promise<readonly SlideSvg[]> {
    if (this.#dirty) {
      await this.renderSlidesFromBytes(writePptx(this.#session.document));
      return this.#slides;
    }
    await this.renderSlidesFromBytes(await readFile(this.sourcePath));
    return this.#slides;
  }

  async renderSlidesFromBytes(input: Uint8Array): Promise<readonly SlideSvg[]> {
    this.#slides = addSlideHandles(await this.#render(input), this.#session.document);
    return this.#slides;
  }

  async apply(command: EditorCommand): Promise<EditorSlidesResponse> {
    return this.#enqueueMutation(async () => {
      const result = this.#session.apply(command);
      if (!result.ok) {
        throw new Error(result.message);
      }
      this.#dirty = true;
      await this.renderCurrentSlides();
      return this.response(result.warnings);
    });
  }

  async addTextBox(slideNumber: number): Promise<EditorSlidesResponse> {
    return this.#enqueueMutation(async () => {
      const slide = this.#session.document.slides[slideNumber - 1];
      if (slide?.handle === undefined) {
        throw new Error("addTextBox: slide handle was not found in PptxSourceModel source");
      }
      const existingShapeKeys = new Set(slide.shapes.map(shapeSourceKey));
      const result = this.#session.apply({
        kind: "addTextBox",
        slideHandle: slide.handle,
        offsetX: pxToEmu(DEFAULT_TEXT_BOX_BOUNDS_PX.x),
        offsetY: pxToEmu(DEFAULT_TEXT_BOX_BOUNDS_PX.y),
        width: pxToEmu(DEFAULT_TEXT_BOX_BOUNDS_PX.width),
        height: pxToEmu(DEFAULT_TEXT_BOX_BOUNDS_PX.height),
        text: DEFAULT_TEXT_BOX_TEXT,
      });
      if (!result.ok) {
        throw new Error(result.message);
      }
      this.#selectNewShape(slideNumber, existingShapeKeys);
      this.#dirty = true;
      await this.renderCurrentSlides();
      return this.response();
    });
  }

  async addConnector(slideNumber: number): Promise<EditorSlidesResponse> {
    return this.#enqueueMutation(async () => {
      const slide = this.#session.document.slides[slideNumber - 1];
      if (slide?.handle === undefined) {
        throw new Error("addConnector: slide handle was not found in PptxSourceModel source");
      }
      const existingShapeKeys = new Set(slide.shapes.map(shapeSourceKey));
      const result = this.#session.apply({
        kind: "addConnector",
        slideHandle: slide.handle,
        preset: "straightConnector1",
        offsetX: pxToEmu(DEFAULT_CONNECTOR_BOUNDS_PX.x),
        offsetY: pxToEmu(DEFAULT_CONNECTOR_BOUNDS_PX.y),
        width: pxToEmu(DEFAULT_CONNECTOR_BOUNDS_PX.width),
        height: pxToEmu(DEFAULT_CONNECTOR_BOUNDS_PX.height),
        outline: {
          tailEnd: { type: "triangle", width: "med", length: "med" },
        },
      });
      if (!result.ok) {
        throw new Error(result.message);
      }
      this.#selectNewShape(slideNumber, existingShapeKeys);
      this.#dirty = true;
      await this.renderCurrentSlides();
      return this.response();
    });
  }

  async applyTextBodyDocJson(
    handle: SourceHandle,
    docJson: unknown,
  ): Promise<EditorSlidesResponse> {
    return this.#enqueueMutation(async () => {
      const textBody = this.#requireEditableShapeTextBody(handle);
      const commands = proseMirrorDocJsonToEditorCommands(textBody, docJson);
      if (commands.length === 0) return this.response();

      for (const command of commands) {
        const result = this.#session.apply(command);
        if (!result.ok) {
          throw new Error(result.message);
        }
      }
      this.#dirty = true;
      await this.renderCurrentSlides();
      return this.response();
    });
  }

  async selectShape(handle: SourceHandle): Promise<EditorSlidesResponse> {
    return this.#enqueueMutation(() => {
      const result = this.#session.selectShape(handle);
      if (!result.ok) {
        throw new Error(result.message);
      }
      return Promise.resolve(this.response());
    });
  }

  async undo(): Promise<EditorSlidesResponse> {
    return this.#enqueueMutation(async () => {
      const result = this.#session.undo();
      if (!result.ok) {
        throw new Error(result.reason);
      }
      this.#dirty = this.#session.undoDepth > 0;
      await this.renderCurrentSlides();
      return this.response();
    });
  }

  async redo(): Promise<EditorSlidesResponse> {
    return this.#enqueueMutation(async () => {
      const result = this.#session.redo();
      if (!result.ok) {
        throw new Error(result.reason);
      }
      this.#dirty = this.#session.undoDepth > 0;
      await this.renderCurrentSlides();
      return this.response();
    });
  }

  async save(outputPath?: string): Promise<EditorSaveResponse> {
    return this.#enqueueMutation(async () => {
      const path = await resolveSavePath(this.sourcePath, outputPath);
      const output = writePptx(this.#session.document);
      readPptx(output);
      await writeFile(path, output);
      return { ok: true, path, history: this.history };
    });
  }

  #enqueueMutation<T>(operation: () => Promise<T>): Promise<T> {
    const queued = this.#mutationQueue.then(operation, operation);
    this.#mutationQueue = queued.then(
      () => undefined,
      () => undefined,
    );
    return queued;
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

function defaultEditedPath(sourcePath: string): string {
  const extension = extname(sourcePath);
  const base = basename(sourcePath, extension);
  return join(dirname(sourcePath), `${base}.edited${extension || ".pptx"}`);
}

async function resolveSavePath(sourcePath: string, outputPath?: string): Promise<string> {
  const sourceDir = dirname(sourcePath);
  const path =
    outputPath === undefined
      ? defaultEditedPath(sourcePath)
      : resolve(isAbsolute(outputPath) ? outputPath : join(sourceDir, outputPath));
  if (extname(path).toLowerCase() !== ".pptx") {
    throw new Error("save path must use the .pptx extension");
  }

  const existingTarget = await lstat(path).catch((error: unknown) => {
    if (isNodeError(error) && error.code === "ENOENT") return null;
    throw error;
  });
  if (existingTarget?.isSymbolicLink()) {
    throw new Error("save path must not be a symbolic link");
  }

  const sourceDirRealPath = await realpath(sourceDir);
  const outputDirRealPath = await realpath(dirname(path));
  const relativePath = relative(sourceDirRealPath, outputDirRealPath);
  if (relativePath.startsWith("..") || isAbsolute(relativePath)) {
    throw new Error("save path must be inside the source PPTX directory");
  }
  return path;
}

function isNodeError(error: unknown): error is NodeJS.ErrnoException {
  return error instanceof Error && "code" in error;
}

function shapeInfo(
  source: ReturnType<typeof readPptx>,
  shape: SourceShapeNode,
  index: number,
  editableTransform = true,
  slideShapes: readonly SourceShapeNode[] = [],
): EditorShapeInfo[] {
  const base: EditorShapeInfo = {
    id: String(shape.nodeId ?? shape.handle?.nodeId ?? `${shape.kind}:${String(index)}`),
    kind: shape.kind,
    ...(shapeName(shape) !== undefined ? { name: shapeName(shape) } : {}),
    ...(shape.handle !== undefined ? { handle: shape.handle } : {}),
    ...(shape.kind !== "raw" && shape.transform !== undefined
      ? {
          bounds: transformBoundsPx(shape.transform),
          editableTransform: editableTransform && isEditableTransformShape(shape),
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
    ...(shape.kind === "shape" && shape.textBody !== undefined && shape.transform !== undefined
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

function collectTextRuns(runs: readonly SourceTextRun[]): EditorTextRunInfo[] {
  return runs.flatMap((run) => {
    if (run.handle === undefined) return [];
    return [{ text: run.text, handle: run.handle }];
  });
}

function editableTextBody(
  textBody: SourceTextBody,
): Partial<Pick<EditorShapeInfo, "editableTextBody">> {
  try {
    return { editableTextBody: { docJson: textBodyToProseMirrorDocJson(textBody) } };
  } catch {
    return {};
  }
}

function editableImageReplacement(
  source: ReturnType<typeof readPptx>,
  image: SourceImage,
): Partial<Pick<EditorShapeInfo, "editableImageReplacement">> {
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

function imageMediaPart(
  source: ReturnType<typeof readPptx>,
  image: SourceImage,
): MediaPart | undefined {
  if (image.blipRelationshipId === undefined || image.handle?.partPath === undefined) {
    return undefined;
  }
  const mediaPartPath = imageMediaPartPath(source, image.handle.partPath, image.blipRelationshipId);
  if (mediaPartPath === undefined) return undefined;
  return source.packageGraph.media.find((part) => part.partPath === mediaPartPath);
}

function imageMediaPartPath(
  source: ReturnType<typeof readPptx>,
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

function countImageReferencesToMedia(
  source: ReturnType<typeof readPptx>,
  mediaPartPath: PartPath,
): number {
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
  source: ReturnType<typeof readPptx>,
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
  return unsafeScriptInputAssertion<PartPath>(
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
}): ShapeBoundsPx {
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

function addSlideHandles(
  slides: readonly SlideSvg[],
  source: ReturnType<typeof readPptx>,
): SlideSvg[] {
  return slides.map((slide) => ({
    ...slide,
    ...(source.slides[slide.slideNumber - 1]?.handle !== undefined
      ? { handle: source.slides[slide.slideNumber - 1]?.handle }
      : {}),
  }));
}

// --- WebSocket ---

function broadcast(wss: WebSocketServer, data: unknown): void {
  const msg = JSON.stringify(data);
  for (const client of wss.clients) {
    if (client.readyState === 1) {
      client.send(msg);
    }
  }
}

// --- File watcher ---

function watchSourceFiles(onChange: () => void): void {
  let timeout: ReturnType<typeof setTimeout> | null = null;

  const handler = (_event: string, filename: string | null) => {
    if (!filename || !filename.endsWith(".ts")) return;
    if (filename.endsWith(".test.ts")) return;

    console.log(`Change detected: ${filename}`);

    if (timeout) clearTimeout(timeout);
    timeout = setTimeout(() => {
      timeout = null;
      onChange();
    }, DEBOUNCE_MS);
  };
  for (const dir of WATCH_DIRS) {
    watch(dir, { recursive: true }, handler);
  }
}

// --- HTML template ---

function generateHtml(slides: SlideSvg[], pptxName: string): string {
  return generateDevEditorHtml({
    slides,
    pptxName,
    emuPerPixel: EMU_PER_PIXEL,
    maxImageReplacementBytes: MAX_IMAGE_REPLACEMENT_BYTES,
  });
}

// --- HTTP API ---

export function createDevServerRequestHandler(
  backend: DevEditorBackend,
  pptxName: string,
): (req: IncomingMessage, res: ServerResponse) => void {
  return (req, res) => {
    void handleRequest(req, res, backend, pptxName).catch((error: unknown) => {
      sendError(res, error);
    });
  };
}

async function handleRequest(
  req: IncomingMessage,
  res: ServerResponse,
  backend: DevEditorBackend,
  pptxName: string,
): Promise<void> {
  const url = new URL(req.url ?? "/", "http://localhost");

  if (req.method === "GET" && (url.pathname === "/" || url.pathname === "/index.html")) {
    res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
    res.end(generateHtml([...backend.slides], pptxName));
    return;
  }

  if (url.pathname === "/api/editor/slides" && req.method === "GET") {
    sendJson(res, 200, backend.response());
    return;
  }

  if (url.pathname === "/api/editor/shapes" && req.method === "GET") {
    const slide = Number(url.searchParams.get("slide") ?? "1");
    sendJson(res, 200, { slideNumber: slide, shapes: backend.shapes(slide) });
    return;
  }

  if (url.pathname === "/api/editor/command" && req.method === "POST") {
    const body = await readJsonBody(req);
    const command = normalizeCommand(isRecord(body) ? body.command : undefined);
    sendJson(res, 200, await backend.apply(command));
    return;
  }

  if (url.pathname === "/api/editor/add-text-box" && req.method === "POST") {
    const body = await readJsonBody(req);
    const slideNumber = isRecord(body) ? body.slide : undefined;
    sendJson(res, 200, await backend.addTextBox(normalizeSlideNumber(slideNumber)));
    return;
  }

  if (url.pathname === "/api/editor/add-connector" && req.method === "POST") {
    const body = await readJsonBody(req);
    const slideNumber = isRecord(body) ? body.slide : undefined;
    sendJson(res, 200, await backend.addConnector(normalizeSlideNumber(slideNumber)));
    return;
  }

  if (url.pathname === "/api/editor/text-body" && req.method === "POST") {
    const body = await readJsonBody(req);
    const handle = normalizeHandle(isRecord(body) ? body.handle : undefined);
    const docJson = isRecord(body) ? body.docJson : undefined;
    sendJson(res, 200, await backend.applyTextBodyDocJson(handle, docJson));
    return;
  }

  if (url.pathname === "/api/editor/select" && req.method === "POST") {
    const body = await readJsonBody(req);
    const handle = normalizeHandle(isRecord(body) ? body.handle : undefined);
    sendJson(res, 200, await backend.selectShape(handle));
    return;
  }

  if (url.pathname === "/api/editor/undo" && req.method === "POST") {
    sendJson(res, 200, await backend.undo());
    return;
  }

  if (url.pathname === "/api/editor/redo" && req.method === "POST") {
    sendJson(res, 200, await backend.redo());
    return;
  }

  if (url.pathname === "/api/editor/save" && req.method === "POST") {
    const body = await readJsonBody(req);
    const path = isRecord(body) ? body.path : undefined;
    sendJson(res, 200, await backend.save(typeof path === "string" ? path : undefined));
    return;
  }

  res.writeHead(404);
  res.end("Not Found");
}

function sendJson(res: ServerResponse, status: number, data: unknown): void {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(data));
}

function sendError(res: ServerResponse, error: unknown): void {
  const message = error instanceof Error ? error.message : String(error);
  sendJson(res, 400, { ok: false, error: message });
}

function readJsonBody(req: IncomingMessage): Promise<unknown> {
  return new Promise((resolveBody, reject) => {
    let raw = "";
    req.setEncoding("utf8");
    req.on("data", (chunk: string) => {
      raw += chunk;
      if (raw.length > MAX_JSON_BODY_BYTES) {
        req.destroy(new Error("Request body is too large"));
      }
    });
    req.on("end", () => {
      if (raw.trim() === "") {
        resolveBody({});
        return;
      }
      try {
        resolveBody(JSON.parse(raw));
      } catch (error) {
        reject(error instanceof Error ? error : new Error(String(error)));
      }
    });
    req.on("error", reject);
  });
}

function normalizeCommand(command: unknown): EditorCommand {
  if (!isRecord(command)) throw new Error("command must be an object");
  if (command.kind === "replaceTextRunPlainText") {
    if (typeof command.text !== "string") {
      throw new Error("replaceTextRunPlainText.text must be a string");
    }
    return {
      kind: "replaceTextRunPlainText",
      handle: normalizeHandle(command.handle),
      text: command.text,
    };
  }
  if (command.kind === "setTextRunProperties") {
    return {
      kind: "setTextRunProperties",
      handle: normalizeHandle(command.handle),
      properties: normalizeTextRunProperties(command.properties),
    };
  }
  if (command.kind === "clearTextRunProperties") {
    return {
      kind: "clearTextRunProperties",
      handle: normalizeHandle(command.handle),
      properties: normalizeTextRunPropertyNames(command.properties),
    };
  }
  if (command.kind === "moveShape") {
    return {
      kind: "moveShape",
      handle: normalizeHandle(command.handle),
      offsetX: asEmu(requireFiniteNumber(command.offsetX, "moveShape.offsetX")),
      offsetY: asEmu(requireFiniteNumber(command.offsetY, "moveShape.offsetY")),
    };
  }
  if (command.kind === "resizeShape") {
    return {
      kind: "resizeShape",
      handle: normalizeHandle(command.handle),
      width: asEmu(requireFiniteNumber(command.width, "resizeShape.width")),
      height: asEmu(requireFiniteNumber(command.height, "resizeShape.height")),
    };
  }
  if (command.kind === "setShapeTransform") {
    return {
      kind: "setShapeTransform",
      handle: normalizeHandle(command.handle),
      offsetX: asEmu(requireFiniteNumber(command.offsetX, "setShapeTransform.offsetX")),
      offsetY: asEmu(requireFiniteNumber(command.offsetY, "setShapeTransform.offsetY")),
      width: asEmu(requireFiniteNumber(command.width, "setShapeTransform.width")),
      height: asEmu(requireFiniteNumber(command.height, "setShapeTransform.height")),
    };
  }
  if (command.kind === "addTextBox") {
    const text = command.text ?? DEFAULT_TEXT_BOX_TEXT;
    if (typeof text !== "string") {
      throw new Error("addTextBox.text must be a string");
    }
    if (command.name !== undefined && typeof command.name !== "string") {
      throw new Error("addTextBox.name must be a string");
    }
    return {
      kind: "addTextBox",
      slideHandle: normalizeHandle(command.slideHandle),
      offsetX: asEmu(requireFiniteNumber(command.offsetX, "addTextBox.offsetX")),
      offsetY: asEmu(requireFiniteNumber(command.offsetY, "addTextBox.offsetY")),
      width: asEmu(requireFinitePositiveNumber(command.width, "addTextBox.width")),
      height: asEmu(requireFinitePositiveNumber(command.height, "addTextBox.height")),
      text,
      ...(typeof command.name === "string" ? { name: command.name } : {}),
    };
  }
  if (command.kind === "addConnector") {
    if (command.name !== undefined && typeof command.name !== "string") {
      throw new Error("addConnector.name must be a string");
    }
    const outline = normalizeConnectorOutline(command.outline);
    return {
      kind: "addConnector",
      slideHandle: normalizeHandle(command.slideHandle),
      preset: normalizeConnectorPreset(command.preset),
      offsetX: asEmu(requireFiniteNumber(command.offsetX, "addConnector.offsetX")),
      offsetY: asEmu(requireFiniteNumber(command.offsetY, "addConnector.offsetY")),
      width: asEmu(requireFinitePositiveNumber(command.width, "addConnector.width")),
      height: asEmu(requireFinitePositiveNumber(command.height, "addConnector.height")),
      ...(command.start !== undefined
        ? { start: normalizeConnectorEndpoint(command.start, "start") }
        : {}),
      ...(command.end !== undefined ? { end: normalizeConnectorEndpoint(command.end, "end") } : {}),
      ...(outline !== undefined ? { outline } : {}),
      ...(typeof command.name === "string" ? { name: command.name } : {}),
    };
  }
  if (command.kind === "deleteShape") {
    return {
      kind: "deleteShape",
      handle: normalizeHandle(command.handle),
    };
  }
  if (command.kind === "duplicateSlide") {
    return {
      kind: "duplicateSlide",
      handle: normalizeHandle(command.handle),
    };
  }
  if (command.kind === "deleteSlide") {
    return {
      kind: "deleteSlide",
      handle: normalizeHandle(command.handle),
    };
  }
  if (command.kind === "replaceImage") {
    return {
      kind: "replaceImage",
      handle: normalizeHandle(command.handle),
      bytes: normalizeByteArray(command.bytes, "replaceImage.bytes"),
    };
  }
  throw new Error("unsupported command kind");
}

function normalizeSlideNumber(value: unknown): number {
  const slideNumber = value === undefined ? 1 : value;
  if (!Number.isInteger(slideNumber) || slideNumber < 1) {
    throw new Error("slide must be a positive integer");
  }
  return slideNumber;
}

function normalizeByteArray(value: unknown, field: string): Uint8Array {
  if (!Array.isArray(value)) throw new Error(`${field} must be an array of bytes`);
  if (value.length > MAX_IMAGE_REPLACEMENT_BYTES) {
    throw new Error(`${field} must not exceed ${String(MAX_IMAGE_REPLACEMENT_BYTES)} bytes`);
  }
  return new Uint8Array(
    value.map((byte, index) => {
      if (!Number.isInteger(byte) || byte < 0 || byte > 255) {
        throw new Error(`${field}[${String(index)}] must be an integer byte`);
      }
      return byte;
    }),
  );
}

function normalizeTextRunProperties(value: unknown): EditableTextRunProperties {
  if (!isRecord(value)) throw new Error("setTextRunProperties.properties must be an object");
  const properties: EditableTextRunProperties = {};
  if ("bold" in value) {
    if (typeof value.bold !== "boolean")
      throw new Error("setTextRunProperties.bold must be boolean");
    properties.bold = value.bold;
  }
  if ("italic" in value) {
    if (typeof value.italic !== "boolean") {
      throw new Error("setTextRunProperties.italic must be boolean");
    }
    properties.italic = value.italic;
  }
  if ("underline" in value) {
    if (typeof value.underline !== "boolean") {
      throw new Error("setTextRunProperties.underline must be boolean");
    }
    properties.underline = value.underline;
  }
  if ("fontSize" in value) {
    properties.fontSize = asPt(requireFinitePositiveNumber(value.fontSize, "fontSize"));
  }
  if ("color" in value) {
    properties.color = normalizeSrgbColor(value.color);
  }
  if ("typeface" in value) {
    if (typeof value.typeface !== "string" || value.typeface.trim().length === 0) {
      throw new Error("setTextRunProperties.typeface must be a non-empty string");
    }
    properties.typeface = value.typeface;
  }
  return properties;
}

function normalizeSrgbColor(value: unknown): { kind: "srgb"; hex: string } {
  if (!isRecord(value) || value.kind !== "srgb" || typeof value.hex !== "string") {
    throw new Error("setTextRunProperties.color must be an srgb color object");
  }
  if (!/^[0-9a-fA-F]{6}$/.test(value.hex)) {
    throw new Error("setTextRunProperties.color.hex must be six hex digits");
  }
  return { kind: "srgb", hex: value.hex.toUpperCase() };
}

function normalizeConnectorPreset(value: unknown): ConnectorPresetGeometry {
  const preset = value ?? "straightConnector1";
  switch (preset) {
    case "straightConnector1":
    case "bentConnector3":
    case "curvedConnector3":
      return preset;
  }
  throw new Error(
    "addConnector.preset must be straightConnector1, bentConnector3, or curvedConnector3",
  );
}

function normalizeConnectorEndpoint(
  value: unknown,
  field: "start" | "end",
): {
  shapeHandle: SourceHandle;
  connectionSiteIndex: number;
} {
  if (!isRecord(value)) throw new Error(`addConnector.${field} must be an object`);
  const connectionSiteIndex = value.connectionSiteIndex;
  if (!Number.isInteger(connectionSiteIndex) || connectionSiteIndex < 0) {
    throw new Error(`addConnector.${field}.connectionSiteIndex must be a non-negative integer`);
  }
  return {
    shapeHandle: normalizeHandle(value.shapeHandle),
    connectionSiteIndex,
  };
}

function normalizeConnectorOutline(
  value: unknown,
): { headEnd?: SourceArrowEndpoint; tailEnd?: SourceArrowEndpoint } | undefined {
  if (value === undefined) return undefined;
  if (!isRecord(value)) throw new Error("addConnector.outline must be an object");
  const headEnd = normalizeArrowEndpoint(value.headEnd, "headEnd");
  const tailEnd = normalizeArrowEndpoint(value.tailEnd, "tailEnd");
  return {
    ...(headEnd !== undefined ? { headEnd } : {}),
    ...(tailEnd !== undefined ? { tailEnd } : {}),
  };
}

function normalizeArrowEndpoint(
  value: unknown,
  field: "headEnd" | "tailEnd",
): SourceArrowEndpoint | undefined {
  if (value === undefined) return undefined;
  if (!isRecord(value)) throw new Error(`addConnector.outline.${field} must be an object`);
  return {
    type: normalizeArrowType(value.type, field),
    width: normalizeArrowSize(value.width, field, "width"),
    length: normalizeArrowSize(value.length, field, "length"),
  };
}

function normalizeArrowType(
  value: unknown,
  field: "headEnd" | "tailEnd",
): SourceArrowEndpoint["type"] {
  switch (value) {
    case "triangle":
    case "stealth":
    case "diamond":
    case "oval":
    case "arrow":
      return value;
  }
  throw new Error(`addConnector.outline.${field}.type is not supported`);
}

function normalizeArrowSize(
  value: unknown,
  field: "headEnd" | "tailEnd",
  sizeField: "width" | "length",
): SourceArrowEndpoint["width"] {
  switch (value) {
    case "sm":
    case "med":
    case "lg":
      return value;
  }
  throw new Error(`addConnector.outline.${field}.${sizeField} is not supported`);
}

function normalizeTextRunPropertyNames(value: unknown): EditableTextRunProperty[] {
  if (!Array.isArray(value)) throw new Error("clearTextRunProperties.properties must be an array");
  return value.map((property): EditableTextRunProperty => {
    if (!isEditableTextRunProperty(property)) {
      throw new Error(
        `clearTextRunProperties: unsupported text run property '${String(property)}'`,
      );
    }
    return property;
  });
}

function isEditableTextRunProperty(value: unknown): value is EditableTextRunProperty {
  return (
    value === "bold" ||
    value === "italic" ||
    value === "underline" ||
    value === "fontSize" ||
    value === "color" ||
    value === "typeface"
  );
}

function normalizeHandle(value: unknown): SourceHandle {
  if (!isRecord(value)) throw new Error("command.handle must be an object");
  if (typeof value.partPath !== "string") {
    throw new Error("command.handle.partPath must be a string");
  }
  return {
    partPath: asPartPath(value.partPath),
    ...(typeof value.nodeId === "string" ? { nodeId: asSourceNodeId(value.nodeId) } : {}),
    ...(typeof value.relationshipId === "string"
      ? { relationshipId: asRelationshipId(value.relationshipId) }
      : {}),
    ...(typeof value.orderingSlot === "number" ? { orderingSlot: value.orderingSlot } : {}),
  };
}

function requireFiniteNumber(value: unknown, field: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${field} must be a finite number`);
  }
  return value;
}

function requireFinitePositiveNumber(value: unknown, field: string): number {
  const number = requireFiniteNumber(value, field);
  if (number <= 0) throw new Error(`${field} must be positive`);
  return number;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

// --- Main ---

function parseCliArgs(argv: readonly string[]): { pptxPath?: string; port: number } {
  const args = argv[0] === "--" ? argv.slice(1) : [...argv];
  const portArgIdx = args.indexOf("--port");
  const port = portArgIdx !== -1 ? Number(args[portArgIdx + 1]) : DEFAULT_PORT;
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error("--port must be an integer between 1 and 65535");
  }
  const pptxPath = args.find((arg, index) => {
    if (arg === "--port") return false;
    if (index > 0 && args[index - 1] === "--port") return false;
    return true;
  });
  return { pptxPath, port };
}

async function main(): Promise<void> {
  const { pptxPath, port } = parseCliArgs(process.argv.slice(2));
  if (!pptxPath) {
    console.error("Usage: npm run dev -- <pptx-file> [--port <port>]");
    process.exit(1);
  }

  const resolvedPath = resolve(pptxPath);
  const pptxName = basename(resolvedPath);

  console.log(`Loading: ${resolvedPath}`);

  const backend = await DevEditorBackend.load(resolvedPath);
  console.log(`Rendered ${String(backend.slides.length)} slide(s)`);

  const server = createServer(createDevServerRequestHandler(backend, pptxName));

  const wss = new WebSocketServer({ server });
  wss.on("connection", (_ws: WebSocket) => {
    console.log("Browser connected");
  });

  server.listen(port, () => {
    console.log(`Dev server running at http://localhost:${String(port)}`);
    console.log(`Watching: ${WATCH_DIRS.join(", ")}`);
  });

  let rendering = false;

  watchSourceFiles(() => {
    if (rendering) return;
    rendering = true;

    console.log("Re-rendering...");
    broadcast(wss, { type: "rendering" });

    backend
      .renderCurrentSlides()
      .then((result) => {
        console.log(`Re-rendered ${String(result.length)} slide(s)`);
        broadcast(wss, { type: "reload" });
      })
      .catch((err: unknown) => {
        const message = err instanceof Error ? err.message : String(err);
        console.error(`Render error: ${message}`);
        broadcast(wss, { type: "error", message });
      })
      .finally(() => {
        rendering = false;
      });
  });
}

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}
