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
  type EditableTextRunProperties,
  type EditableTextRunProperty,
  findShapeNodeBySourceHandle,
  readPptx,
  type SourceHandle,
  type SourceShapeNode,
  type SourceTextBody,
  type SourceTextRun,
  writePptx,
} from "../packages/document/src/index.js";
import {
  createEditorSession,
  type EditorCommand,
  type PptxTextBodyProseMirrorDocJson,
  proseMirrorDocJsonToEditorCommands,
  textBodyToProseMirrorDocJson,
} from "../packages/editor-core/src/index.js";
import { unsafeScriptInputAssertion } from "./unsafe-type-assertion.js";

const DEFAULT_PORT = 3000;
const DEBOUNCE_MS = 300;
// The in-process editor session imports document/editor-core once; restart the server for those changes.
const WATCH_DIRS = [resolve("packages/core/src"), resolve("packages/renderer/src")];
const RENDER_TIMEOUT_MS = 30_000;
const MAX_BUFFER = 50 * 1024 * 1024;
const EMU_PER_INCH = 914400;
const DEFAULT_DPI = 96;

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

interface EditorShapeInfo {
  id: string;
  kind: SourceShapeNode["kind"];
  name?: string;
  handle?: SourceHandle;
  bounds?: ShapeBoundsPx;
  editableTransform?: boolean;
  textRuns?: EditorTextRunInfo[];
  editableTextBody?: EditorTextBodyInfo;
}

interface EditorSlidesResponse {
  slides: SlideSvg[];
  history: EditorHistoryState;
  selection?: EditorSelectionInfo;
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

  response(): EditorSlidesResponse {
    return {
      slides: this.#slides,
      history: this.history,
      ...(this.selection !== undefined ? { selection: this.selection } : {}),
    };
  }

  shapes(slideNumber: number): EditorShapeInfo[] {
    const slide = this.#session.document.slides[slideNumber - 1];
    if (slide === undefined) return [];
    return slide.shapes.flatMap((shape, index) => shapeInfo(shape, index));
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
  shape: SourceShapeNode,
  index: number,
  editableTransform = true,
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
  };

  if (shape.kind !== "group") return [base];
  return [
    base,
    ...shape.children.flatMap((child, childIndex) => shapeInfo(child, childIndex, false)),
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
  const thumbnailsHtml = generateThumbnailsHtml(slides);

  const firstSvg = slides.length > 0 ? slides[0].svg : "<p>No slides</p>";

  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>pptx-glimpse dev - ${escapeHtml(pptxName)}</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: #1a1a2e;
      color: #e0e0e0;
    }
    #header {
      padding: 12px 20px;
      background: #16213e;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }
    #header h1 { font-size: 14px; font-weight: 600; color: #a0a0c0; }
    #status { font-size: 12px; color: #4caf50; }
    #status.rendering { color: #ff9800; }
    #status.error { color: #f44336; }
    #main { display: flex; height: calc(100vh - 48px); }
    #sidebar {
      width: 180px;
      overflow-y: auto;
      background: #16213e;
      padding: 8px;
      border-right: 1px solid #2a2a4a;
    }
    #editor-panel {
      width: 260px;
      background: #111827;
      border-left: 1px solid #2a2a4a;
      padding: 12px;
      display: flex;
      flex-direction: column;
      gap: 10px;
    }
    #editor-panel label {
      display: flex;
      flex-direction: column;
      gap: 4px;
      color: #a0a0c0;
      font-size: 11px;
      font-weight: 600;
    }
    #editor-panel select,
    #editor-panel input {
      width: 100%;
      min-height: 32px;
      border: 1px solid #334155;
      border-radius: 4px;
      background: #0f172a;
      color: #e5e7eb;
      padding: 6px 8px;
      font: inherit;
    }
    #editor-panel button {
      min-height: 32px;
      border: 1px solid #475569;
      border-radius: 4px;
      background: #1f2937;
      color: #f8fafc;
      cursor: pointer;
      font: inherit;
      font-weight: 600;
    }
    #editor-panel button:hover:not(:disabled) { background: #334155; }
    #editor-panel button:disabled {
      opacity: 0.45;
      cursor: not-allowed;
    }
    #editor-actions {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 8px;
    }
    #save-button { grid-column: 1 / -1; }
    #editor-message {
      min-height: 18px;
      color: #94a3b8;
      font-size: 11px;
      overflow-wrap: anywhere;
    }
    .thumbnail {
      margin-bottom: 8px;
      padding: 4px;
      border: 2px solid transparent;
      border-radius: 4px;
      cursor: pointer;
      background: #fff;
    }
    .thumbnail.active { border-color: #4472c4; }
    .thumbnail:hover { border-color: #6090d0; }
    .thumb-label {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 4px;
      font-size: 10px;
      color: #888;
      padding: 2px 0;
      background: #16213e;
    }
    .thumb-title { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .thumb-actions { display: flex; gap: 3px; }
    .thumb-action {
      min-width: 22px;
      min-height: 20px;
      border: 1px solid #334155;
      border-radius: 3px;
      background: #0f172a;
      color: #e5e7eb;
      cursor: pointer;
      font-size: 10px;
      font-weight: 700;
      line-height: 1;
    }
    .thumb-action:hover:not(:disabled) { background: #1d4ed8; }
    .thumb-action:disabled { opacity: 0.4; cursor: not-allowed; }
    .thumb-svg svg { width: 100%; height: auto; display: block; }
    #viewer {
      flex: 1;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
      overflow: auto;
    }
    #slide-container {
      background: #fff;
      border-radius: 4px;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
      max-width: 100%;
      max-height: 100%;
      position: relative;
    }
    #slide-container svg { display: block; width: 100%; height: auto; }
    #selection-overlay {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      overflow: visible;
      touch-action: none;
      z-index: 2;
    }
    #text-editor-overlay {
      position: absolute;
      min-width: 40px;
      min-height: 28px;
      padding: 4px;
      border: 1px solid #2563eb;
      background: rgba(255, 255, 255, 0.96);
      color: #111827;
      z-index: 3;
      overflow: hidden;
      box-shadow: 0 0 0 1px rgba(37, 99, 235, 0.25);
    }
    .text-editor-paragraph {
      min-height: 20px;
      line-height: 1.25;
      white-space: pre-wrap;
      overflow-wrap: anywhere;
    }
    .text-editor-run {
      outline: none;
      white-space: pre-wrap;
    }
    .text-run-format-toolbar {
      display: grid;
      grid-template-columns: repeat(3, 24px) 44px 20px 28px 20px minmax(48px, 1fr) 20px;
      gap: 3px;
      align-items: center;
      margin-bottom: 4px;
    }
    .text-run-format-toolbar button,
    .text-run-format-toolbar input {
      min-height: 22px;
      border: 1px solid #94a3b8;
      border-radius: 4px;
      background: #f8fafc;
      color: #0f172a;
      font-size: 10px;
      font-weight: 650;
    }
    .text-run-format-toolbar button {
      padding: 0;
      min-width: 0;
    }
    .text-run-format-toolbar button[aria-pressed="true"] {
      background: #dbeafe;
      border-color: #2563eb;
    }
    .text-run-format-toolbar input {
      min-width: 0;
      padding: 2px 4px;
      font-weight: 500;
    }
    .text-run-format-toolbar input[type="color"] {
      padding: 1px;
    }
    .text-run-format-toolbar :disabled {
      opacity: 0.45;
      cursor: not-allowed;
    }
    .text-editor-actions {
      display: flex;
      justify-content: flex-end;
      gap: 4px;
      margin-top: 4px;
    }
    .text-editor-actions button {
      min-height: 24px;
      padding: 2px 8px;
      border: 1px solid #64748b;
      border-radius: 4px;
      background: #f8fafc;
      color: #0f172a;
      font-size: 11px;
      font-weight: 600;
    }
    .shape-hit-area {
      cursor: move;
      fill: transparent;
      pointer-events: all;
    }
    .selection-box {
      fill: none;
      stroke: #0f766e;
      stroke-width: 1.5;
      vector-effect: non-scaling-stroke;
      pointer-events: none;
    }
    .selection-handle {
      fill: #f8fafc;
      stroke: #0f766e;
      stroke-width: 1.5;
      cursor: nwse-resize;
      vector-effect: non-scaling-stroke;
      pointer-events: all;
    }
    .selection-handle[data-handle="ne"],
    .selection-handle[data-handle="sw"] {
      cursor: nesw-resize;
    }
    #info {
      padding: 4px 20px;
      background: #16213e;
      font-size: 11px;
      color: #666;
      text-align: center;
    }
  </style>
</head>
<body>
  <div id="header">
    <h1>pptx-glimpse dev &mdash; ${escapeHtml(pptxName)}</h1>
    <span id="status">Connected</span>
  </div>
  <div id="main">
    <div id="sidebar">${thumbnailsHtml}</div>
    <div id="viewer">
      <div id="slide-container">${firstSvg}</div>
    </div>
    <div id="editor-panel">
      <label>Text run<select id="text-run-select"></select></label>
      <label>Text<input id="text-run-input" type="text"></label>
      <button id="apply-text-button" type="button">Apply</button>
      <div id="editor-actions">
        <button id="undo-button" type="button">Undo</button>
        <button id="redo-button" type="button">Redo</button>
        <button id="save-button" type="button">Save</button>
      </div>
      <div id="editor-message"></div>
    </div>
  </div>
  <div id="info">Slide 1 / ${String(slides.length)}</div>
  <script>
    var currentIndex = 0;
    var slideCount = ${String(slides.length)};
    var slides = ${jsonForScript(slides)};
    var editorHistory = { canUndo: false, canRedo: false, undoDepth: 0, redoDepth: 0 };
    var textRunOptions = [];
    var shapeOptions = [];
    var selectedShapeKey = null;
    var selectedShape = null;
    var dragState = null;
    var activeTextEditor = null;
    var shapeRequestId = 0;
    var EMU_PER_PIXEL = ${String(EMU_PER_INCH / DEFAULT_DPI)};

    function selectSlide(index) {
      if (activeTextEditor) {
        commitTextEditor()
          .then(function () {
            selectSlide(index);
          })
          .catch(function () {
            // Keep the editor open; commitTextEditor already reported the failure.
          });
        return;
      }
      var slideChanged = currentIndex !== index;
      currentIndex = index;
      shapeOptions = [];
      textRunOptions = [];
      selectedShape = null;
      if (slideChanged) selectedShapeKey = null;
      var thumbs = document.querySelectorAll(".thumbnail");
      for (var i = 0; i < thumbs.length; i++) {
        if (i === index) {
          thumbs[i].classList.add("active");
        } else {
          thumbs[i].classList.remove("active");
        }
      }
      document.getElementById("slide-container").innerHTML =
        slides[index] ? slides[index].svg : "<p>No slides</p>";
      var svg = document.querySelector("#slide-container svg");
      if (svg) {
        svg.removeAttribute("width");
        svg.removeAttribute("height");
        svg.style.width = "100%";
        svg.style.height = "auto";
      }
      document.getElementById("info").textContent =
        "Slide " + (index + 1) + " / " + slideCount;
      renderSelectionOverlay();
      loadShapeOptions(index + 1);
    }

    function renderThumbnails() {
      var sidebar = document.getElementById("sidebar");
      sidebar.innerHTML = slides
        .map(function (s, i) {
          return '<div class="thumbnail' + (i === currentIndex ? ' active' : '') +
            '" data-index="' + i + '">' +
            '<div class="thumb-label">' +
            '<span class="thumb-title">Slide ' + s.slideNumber + '</span>' +
            '<span class="thumb-actions">' +
            '<button class="thumb-action" data-testid="duplicate-slide-' + i + '" data-action="duplicate" data-index="' + i + '" type="button" title="Duplicate slide">D</button>' +
            '<button class="thumb-action" data-testid="delete-slide-' + i + '" data-action="delete" data-index="' + i + '" type="button" title="Delete slide"' +
            (slideCount <= 1 ? ' disabled' : '') + '>X</button>' +
            '</span>' +
            '</div>' +
            '<div class="thumb-svg">' + s.svg + '</div>' +
            '</div>';
        })
        .join("");

      var thumbs = document.querySelectorAll(".thumbnail");
      for (var i = 0; i < thumbs.length; i++) {
        (function (idx) {
          thumbs[idx].addEventListener("click", function () {
            selectSlide(idx);
          });
        })(i);
      }

      Array.prototype.forEach.call(sidebar.querySelectorAll("[data-action]"), function (button) {
        button.addEventListener("mousedown", function (event) {
          event.preventDefault();
        });
        button.addEventListener("click", function (event) {
          event.preventDefault();
          event.stopPropagation();
          var index = Number(button.getAttribute("data-index") || "-1");
          var action = button.getAttribute("data-action");
          if (action === "duplicate") duplicateSlide(index);
          if (action === "delete") deleteSlide(index);
        });
      });
    }

    function updateHistory(history) {
      editorHistory = history || editorHistory;
      document.getElementById("undo-button").disabled = !editorHistory.canUndo;
      document.getElementById("redo-button").disabled = !editorHistory.canRedo;
    }

    function shapeKey(shape) {
      return shape && shape.handle ? handleKey(shape.handle) : "";
    }

    function handleKey(handle) {
      return [
        handle.partPath || "",
        handle.nodeId || "",
        handle.relationshipId || "",
        handle.orderingSlot == null ? "" : String(handle.orderingSlot)
      ].join("\\u0000");
    }

    function syncSelection(selection) {
      selectedShapeKey = selection && selection.shapeHandle ? handleKey(selection.shapeHandle) : selectedShapeKey;
      selectedShape = null;
      if (selectedShapeKey) {
        for (var i = 0; i < shapeOptions.length; i++) {
          if (shapeKey(shapeOptions[i]) === selectedShapeKey && shapeOptions[i].bounds) {
            selectedShape = cloneShape(shapeOptions[i]);
            break;
          }
        }
      }
      if (!selectedShape) selectedShapeKey = null;
      renderSelectionOverlay();
    }

    function cloneShape(shape) {
      return {
        id: shape.id,
        name: shape.name,
        handle: shape.handle,
        editableTextBody: shape.editableTextBody,
        bounds: {
          x: shape.bounds.x,
          y: shape.bounds.y,
          width: shape.bounds.width,
          height: shape.bounds.height
        }
      };
    }

    function setEditorMessage(message, isError) {
      var element = document.getElementById("editor-message");
      element.textContent = message;
      element.style.color = isError ? "#fca5a5" : "#94a3b8";
    }

    function loadShapeOptions(slideNumber) {
      var requestId = ++shapeRequestId;
      fetch("/api/editor/shapes?slide=" + slideNumber)
        .then(function (res) {
          if (!res.ok) throw new Error("shape request failed");
          return res.json();
        })
        .then(function (data) {
          if (requestId !== shapeRequestId || slideNumber !== currentIndex + 1) return;
          shapeOptions = (data.shapes || []).filter(function (shape) {
            return shape.handle && shape.bounds && shape.editableTransform;
          });
          textRunOptions = [];
          data.shapes.forEach(function (shape) {
            (shape.textRuns || []).forEach(function (run, index) {
              textRunOptions.push({
                label: (shape.name || shape.id) + " / " + (index + 1),
                text: run.text,
                handle: run.handle
              });
            });
          });
          var select = document.getElementById("text-run-select");
          select.innerHTML = textRunOptions
            .map(function (option, index) {
              return '<option value="' + index + '">' + escapeHtmlClient(option.label) + '</option>';
            })
            .join("");
          select.disabled = textRunOptions.length === 0;
          document.getElementById("apply-text-button").disabled = textRunOptions.length === 0;
          syncTextInput();
          syncSelection();
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function syncTextInput() {
      var select = document.getElementById("text-run-select");
      var input = document.getElementById("text-run-input");
      var option = textRunOptions[Number(select.value || "0")];
      input.value = option ? option.text : "";
      input.disabled = !option;
    }

    function applyEditorResponse(data, preferredIndex) {
      slides = data.slides || slides;
      slideCount = slides.length;
      updateHistory(data.history);
      if (data.selection && data.selection.shapeHandle) {
        selectedShapeKey = handleKey(data.selection.shapeHandle);
      }
      renderThumbnails();
      var nextIndex = preferredIndex == null ? currentIndex : preferredIndex;
      selectSlide(Math.min(Math.max(nextIndex, 0), Math.max(slides.length - 1, 0)));
    }

    function renderSelectionOverlay() {
      var container = document.getElementById("slide-container");
      var renderedSvg = container.querySelector("svg:not(#selection-overlay)");
      var existing = document.getElementById("selection-overlay");
      if (existing) existing.remove();
      if (!renderedSvg) return;

      var viewBox = renderedSvg.getAttribute("viewBox") || "0 0 960 540";
      var overlay = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      overlay.setAttribute("id", "selection-overlay");
      overlay.setAttribute("viewBox", viewBox);
      overlay.setAttribute("data-testid", "selection-overlay");

      shapeOptions.forEach(function (shape, index) {
        if (!shape.bounds) return;
        var hit = document.createElementNS("http://www.w3.org/2000/svg", "rect");
        hit.setAttribute("class", "shape-hit-area");
        hit.setAttribute("data-testid", "shape-hit-area");
        hit.setAttribute("data-shape-index", String(index));
        setRectAttributes(hit, shape.bounds);
        hit.addEventListener("pointerdown", function (event) {
          selectShape(shape, event);
        });
        hit.addEventListener("mousedown", function (event) {
          if (activeTextEditor) return;
          if (event.detail >= 2 && shape.editableTextBody) {
            event.preventDefault();
            openTextEditor(shape);
          }
        });
        hit.addEventListener("dblclick", function (event) {
          event.preventDefault();
          event.stopPropagation();
          openTextEditor(shape);
        });
        overlay.appendChild(hit);
      });

      if (selectedShape && selectedShape.bounds) {
        var box = document.createElementNS("http://www.w3.org/2000/svg", "rect");
        box.setAttribute("class", "selection-box");
        box.setAttribute("data-testid", "selection-box");
        setRectAttributes(box, selectedShape.bounds);
        overlay.appendChild(box);

        ["nw", "ne", "sw", "se"].forEach(function (handle) {
          var point = handlePoint(selectedShape.bounds, handle);
          var rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
          rect.setAttribute("class", "selection-handle");
          rect.setAttribute("data-testid", "selection-handle-" + handle);
          rect.setAttribute("data-handle", handle);
          rect.setAttribute("x", String(point.x - 4));
          rect.setAttribute("y", String(point.y - 4));
          rect.setAttribute("width", "8");
          rect.setAttribute("height", "8");
          rect.addEventListener("pointerdown", function (event) {
            event.preventDefault();
            event.stopPropagation();
            beginDrag("resize", handle, event);
          });
          overlay.appendChild(rect);
        });
      }

      container.appendChild(overlay);
      positionActiveTextEditor();
    }

    function setRectAttributes(rect, bounds) {
      rect.setAttribute("x", String(bounds.x));
      rect.setAttribute("y", String(bounds.y));
      rect.setAttribute("width", String(bounds.width));
      rect.setAttribute("height", String(bounds.height));
    }

    function handlePoint(bounds, handle) {
      var right = bounds.x + bounds.width;
      var bottom = bounds.y + bounds.height;
      if (handle === "nw") return { x: bounds.x, y: bounds.y };
      if (handle === "ne") return { x: right, y: bounds.y };
      if (handle === "sw") return { x: bounds.x, y: bottom };
      return { x: right, y: bottom };
    }

    function selectShape(shape, event) {
      if (activeTextEditor) {
        commitTextEditor()
          .then(function () {
            selectShapeAfterCommit(shape);
          })
          .catch(function () {});
        return;
      }
      selectedShapeKey = shapeKey(shape);
      selectedShape = cloneShape(shape);
      beginDrag("move", null, event);
      window.setTimeout(renderSelectionOverlay, 0);
      postJson("/api/editor/select", { handle: shape.handle })
        .then(function (data) {
          updateHistory(data.history);
          if (data.selection && data.selection.shapeHandle) {
            selectedShapeKey = handleKey(data.selection.shapeHandle);
          }
          renderSelectionOverlay();
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function selectShapeAfterCommit(shape) {
      selectedShapeKey = shapeKey(shape);
      selectedShape = cloneShape(shape);
      renderSelectionOverlay();
      postJson("/api/editor/select", { handle: shape.handle })
        .then(function (data) {
          updateHistory(data.history);
          if (data.selection && data.selection.shapeHandle) {
            selectedShapeKey = handleKey(data.selection.shapeHandle);
          }
          renderSelectionOverlay();
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function openTextEditor(shape) {
      if (!shape || !shape.handle || !shape.bounds || !shape.editableTextBody) return;
      if (activeTextEditor) return;
      if (dragState) {
        dragState = null;
        detachDragListeners();
      }
      closeTextEditor();
      selectedShapeKey = shapeKey(shape);
      selectedShape = cloneShape(shape);
      renderSelectionOverlay();

      var container = document.getElementById("slide-container");
      var overlay = document.createElement("div");
      overlay.id = "text-editor-overlay";
      overlay.setAttribute("data-testid", "text-editor-overlay");
      overlay.dataset.shapeKey = selectedShapeKey || "";
      overlay.appendChild(createTextRunFormatToolbar());
      overlay.appendChild(createTextEditorContent(shape.editableTextBody.docJson));

      var actions = document.createElement("div");
      actions.className = "text-editor-actions";
      var done = document.createElement("button");
      done.type = "button";
      done.textContent = "Done";
      done.setAttribute("data-testid", "text-editor-done");
      done.addEventListener("click", function () {
        commitTextEditor().catch(function () {});
      });
      actions.appendChild(done);
      overlay.appendChild(actions);

      overlay.addEventListener("keydown", function (event) {
        if (event.isComposing || event.keyCode === 229) return;
        if (event.key === "Enter" && !event.shiftKey) {
          event.preventDefault();
          commitTextEditor().catch(function () {});
        }
        if (event.key === "Escape") {
          event.preventDefault();
          closeTextEditor();
        }
      });
      overlay.addEventListener("focusout", function (event) {
        window.setTimeout(function () {
          if (!activeTextEditor) return;
          if (event.relatedTarget && activeTextEditor.element.contains(event.relatedTarget)) return;
          commitTextEditor().catch(function () {});
        }, 0);
      });

      container.appendChild(overlay);
      activeTextEditor = {
        element: overlay,
        shape: cloneShape(shape),
        originalDocJson: shape.editableTextBody.docJson,
        selectedRunElement: null,
        committing: false,
        commitPromise: null
      };
      positionActiveTextEditor();
      var firstRun = overlay.querySelector(".text-editor-run");
      if (firstRun) {
        setActiveTextRunElement(firstRun);
        firstRun.focus();
        selectElementContents(firstRun);
      }
    }

    function createTextRunFormatToolbar() {
      var toolbar = document.createElement("div");
      toolbar.className = "text-run-format-toolbar";
      toolbar.setAttribute("data-testid", "text-run-format-toolbar");

      [["bold", "B"], ["italic", "I"], ["underline", "U"]].forEach(function (item) {
        var button = document.createElement("button");
        button.type = "button";
        button.textContent = item[1];
        button.dataset.property = item[0];
        button.setAttribute("data-testid", "text-run-format-" + item[0]);
        button.setAttribute("aria-pressed", "false");
        button.addEventListener("click", function () {
          toggleActiveTextRunBooleanProperty(item[0]);
        });
        toolbar.appendChild(button);
      });

      var size = document.createElement("input");
      size.type = "number";
      size.min = "1";
      size.step = "0.5";
      size.placeholder = "pt";
      size.setAttribute("data-testid", "text-run-format-font-size");
      size.addEventListener("change", function () {
        var value = Number(size.value);
        if (size.value.trim() === "") {
          applyActiveTextRunPropertyClear(["fontSize"]);
        } else if (Number.isFinite(value) && value > 0) {
          applyActiveTextRunPropertySet({ fontSize: value });
        } else {
          setEditorMessage("font size must be positive", true);
        }
      });
      toolbar.appendChild(size);

      var clearSize = clearButton("font-size", function () {
        applyActiveTextRunPropertyClear(["fontSize"]);
      });
      toolbar.appendChild(clearSize);

      var color = document.createElement("input");
      color.type = "color";
      color.setAttribute("data-testid", "text-run-format-color");
      color.addEventListener("change", function () {
        applyActiveTextRunPropertySet({
          color: { kind: "srgb", hex: color.value.replace(/^#/, "").toUpperCase() }
        });
      });
      toolbar.appendChild(color);

      var clearColor = clearButton("color", function () {
        applyActiveTextRunPropertyClear(["color"]);
      });
      toolbar.appendChild(clearColor);

      var typeface = document.createElement("input");
      typeface.type = "text";
      typeface.placeholder = "Font";
      typeface.setAttribute("data-testid", "text-run-format-typeface");
      typeface.addEventListener("keydown", function (event) {
        if (event.key === "Enter") {
          event.preventDefault();
          applyTypefaceInput(typeface);
        }
      });
      typeface.addEventListener("change", function () {
        applyTypefaceInput(typeface);
      });
      toolbar.appendChild(typeface);

      var clearTypeface = clearButton("typeface", function () {
        applyActiveTextRunPropertyClear(["typeface"]);
      });
      toolbar.appendChild(clearTypeface);

      window.setTimeout(refreshTextRunFormatToolbar, 0);
      return toolbar;
    }

    function clearButton(name, onClick) {
      var button = document.createElement("button");
      button.type = "button";
      button.textContent = "x";
      button.setAttribute("data-testid", "text-run-format-clear-" + name);
      button.addEventListener("click", onClick);
      return button;
    }

    function applyTypefaceInput(input) {
      var value = input.value.trim();
      if (value.length === 0) {
        applyActiveTextRunPropertyClear(["typeface"]);
        return;
      }
      applyActiveTextRunPropertySet({ typeface: value });
    }

    function setActiveTextRunElement(element) {
      if (!activeTextEditor) return;
      activeTextEditor.selectedRunElement = element;
      refreshTextRunFormatToolbar();
    }

    function activeTextRunInfo() {
      if (!activeTextEditor || !activeTextEditor.selectedRunElement) return null;
      var element = activeTextEditor.selectedRunElement;
      var paragraphIndex = Number(element.dataset.paragraphIndex || "-1");
      var runIndex = Number(element.dataset.runIndex || "-1");
      var paragraph = (activeTextEditor.originalDocJson.content || [])[paragraphIndex];
      var textNode = paragraph && (paragraph.content || [])[runIndex];
      var mark = textNode && (textNode.marks || []).find(function (candidate) {
        return candidate.type === "pptxRun";
      });
      var attrs = mark && mark.attrs ? mark.attrs : {};
      return {
        handle: attrs.handle || null,
        properties: attrs.properties || {}
      };
    }

    function refreshTextRunFormatToolbar() {
      if (!activeTextEditor) return;
      var toolbar = activeTextEditor.element.querySelector('[data-testid="text-run-format-toolbar"]');
      if (!toolbar) return;
      var info = activeTextRunInfo();
      var properties = info && info.properties ? info.properties : {};
      var disabled = !info || !info.handle;
      setToolbarControl(toolbar, "bold", disabled, Boolean(properties.bold));
      setToolbarControl(toolbar, "italic", disabled, Boolean(properties.italic));
      setToolbarControl(toolbar, "underline", disabled, Boolean(properties.underline));
      var size = toolbar.querySelector('[data-testid="text-run-format-font-size"]');
      if (size) {
        size.disabled = disabled;
        size.value = properties.fontSize == null ? "" : String(properties.fontSize);
      }
      var color = toolbar.querySelector('[data-testid="text-run-format-color"]');
      if (color) {
        color.disabled = disabled;
        color.value =
          properties.color && properties.color.kind === "srgb" && typeof properties.color.hex === "string"
            ? "#" + properties.color.hex
            : "#000000";
      }
      var typeface = toolbar.querySelector('[data-testid="text-run-format-typeface"]');
      if (typeface) {
        typeface.disabled = disabled;
        typeface.value = typeof properties.typeface === "string" ? properties.typeface : "";
      }
      Array.prototype.forEach.call(toolbar.querySelectorAll('[data-testid^="text-run-format-clear-"]'), function (control) {
        control.disabled = disabled;
      });
    }

    function setToolbarControl(toolbar, property, disabled, pressed) {
      var control = toolbar.querySelector('[data-testid="text-run-format-' + property + '"]');
      if (!control) return;
      control.disabled = disabled;
      control.setAttribute("aria-pressed", pressed ? "true" : "false");
    }

    function toggleActiveTextRunBooleanProperty(property) {
      var info = activeTextRunInfo();
      if (!info || !info.handle) {
        setEditorMessage("No editable text run selected", true);
        return;
      }
      if ((info.properties || {})[property] === true) {
        applyActiveTextRunPropertyClear([property]);
        return;
      }
      var properties = {};
      properties[property] = true;
      applyActiveTextRunPropertySet(properties);
    }

    function applyActiveTextRunPropertySet(properties) {
      var info = activeTextRunInfo();
      if (!info || !info.handle) {
        setEditorMessage("No editable text run selected", true);
        return;
      }
      applyActiveTextRunPropertyCommand({
        kind: "setTextRunProperties",
        handle: info.handle,
        properties: properties
      });
    }

    function applyActiveTextRunPropertyClear(properties) {
      var info = activeTextRunInfo();
      if (!info || !info.handle) {
        setEditorMessage("No editable text run selected", true);
        return;
      }
      applyActiveTextRunPropertyCommand({
        kind: "clearTextRunProperties",
        handle: info.handle,
        properties: properties
      });
    }

    function applyActiveTextRunPropertyCommand(command) {
      postJson("/api/editor/command", { command: command })
        .then(function (data) {
          applyEditorResponseBehindTextEditor(data);
          patchActiveTextRunProperties(command);
          setEditorMessage("Applied", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function applyEditorResponseBehindTextEditor(data) {
      slides = data.slides || slides;
      slideCount = slides.length;
      updateHistory(data.history);
      if (data.selection && data.selection.shapeHandle) {
        selectedShapeKey = handleKey(data.selection.shapeHandle);
      }
      renderThumbnails();
      replaceCurrentRenderedSvg();
      renderSelectionOverlay();
      loadShapeOptions(currentIndex + 1);
    }

    function replaceCurrentRenderedSvg() {
      var container = document.getElementById("slide-container");
      var previous = container.querySelector("svg:not(#selection-overlay)");
      var slide = slides[currentIndex];
      if (!slide) return;
      var wrapper = document.createElement("div");
      wrapper.innerHTML = slide.svg;
      var next = wrapper.querySelector("svg");
      if (!next) return;
      next.removeAttribute("width");
      next.removeAttribute("height");
      next.style.width = "100%";
      next.style.height = "auto";
      if (previous) {
        previous.replaceWith(next);
      } else {
        container.insertBefore(next, container.firstChild);
      }
    }

    function patchActiveTextRunProperties(command) {
      var textNode = activeTextRunTextNode();
      if (!textNode) return;
      var mark = (textNode.marks || []).find(function (candidate) {
        return candidate.type === "pptxRun";
      });
      if (!mark) return;
      var attrs = mark.attrs || {};
      var properties = { ...(attrs.properties || {}) };
      if (command.kind === "setTextRunProperties") {
        Object.keys(command.properties || {}).forEach(function (property) {
          properties[property] = command.properties[property];
        });
      }
      if (command.kind === "clearTextRunProperties") {
        (command.properties || []).forEach(function (property) {
          delete properties[property];
        });
      }
      attrs.properties = Object.keys(properties).length > 0 ? properties : null;
      mark.attrs = attrs;
      refreshTextRunFormatToolbar();
      syncActiveTextRunStyle(properties);
    }

    function activeTextRunTextNode() {
      if (!activeTextEditor || !activeTextEditor.selectedRunElement) return null;
      var element = activeTextEditor.selectedRunElement;
      var paragraphIndex = Number(element.dataset.paragraphIndex || "-1");
      var runIndex = Number(element.dataset.runIndex || "-1");
      var paragraph = (activeTextEditor.originalDocJson.content || [])[paragraphIndex];
      return paragraph && (paragraph.content || [])[runIndex] ? paragraph.content[runIndex] : null;
    }

    function syncActiveTextRunStyle(properties) {
      if (!activeTextEditor || !activeTextEditor.selectedRunElement) return;
      var run = activeTextEditor.selectedRunElement;
      run.style.fontWeight = properties.bold === true ? "700" : "";
      run.style.fontStyle = properties.italic === true ? "italic" : "";
      run.style.textDecoration = properties.underline === true ? "underline" : "";
      run.style.fontSize = properties.fontSize != null ? String(properties.fontSize) + "pt" : "";
      run.style.fontFamily = properties.typeface
        ? '"' + String(properties.typeface).replace(/"/g, "") + '"'
        : "";
      run.style.color =
        properties.color && properties.color.kind === "srgb" && typeof properties.color.hex === "string"
          ? "#" + properties.color.hex
          : "";
    }

    function createTextEditorContent(docJson) {
      var body = document.createElement("div");
      body.setAttribute("data-testid", "text-editor-content");
      (docJson.content || []).forEach(function (paragraph, paragraphIndex) {
        var paragraphElement = document.createElement("div");
        paragraphElement.className = "text-editor-paragraph";
        paragraphElement.dataset.paragraphIndex = String(paragraphIndex);
        (paragraph.content || []).forEach(function (textNode, runIndex) {
          var run = document.createElement("span");
          run.className = "text-editor-run";
          run.setAttribute("data-testid", "text-editor-run");
          run.contentEditable = "true";
          run.dataset.paragraphIndex = String(paragraphIndex);
          run.dataset.runIndex = String(runIndex);
          run.textContent = textNode.text || "";
          applyTextRunStyle(run, textNode);
          run.addEventListener("focus", function () {
            setActiveTextRunElement(run);
          });
          run.addEventListener("pointerdown", function () {
            setActiveTextRunElement(run);
          });
          run.addEventListener("beforeinput", function (event) {
            if (event.inputType === "insertParagraph" || event.inputType === "insertLineBreak") {
              event.preventDefault();
              commitTextEditor().catch(function () {});
            }
          });
          run.addEventListener("paste", function (event) {
            event.preventDefault();
            var text = (event.clipboardData ? event.clipboardData.getData("text/plain") : "")
              .replace(/\\r?\\n/g, " ");
            document.execCommand("insertText", false, text);
          });
          paragraphElement.appendChild(run);
        });
        body.appendChild(paragraphElement);
      });
      return body;
    }

    function applyTextRunStyle(run, textNode) {
      var mark = (textNode.marks || []).find(function (candidate) {
        return candidate.type === "pptxRun";
      });
      var properties = mark && mark.attrs && mark.attrs.properties ? mark.attrs.properties : {};
      if (properties.bold === true) run.style.fontWeight = "700";
      if (properties.italic === true) run.style.fontStyle = "italic";
      if (properties.underline === true) run.style.textDecoration = "underline";
      if (properties.fontSize != null) run.style.fontSize = String(properties.fontSize) + "pt";
      if (properties.typeface) run.style.fontFamily = '"' + String(properties.typeface).replace(/"/g, "") + '"';
      if (properties.color && properties.color.kind === "srgb" && typeof properties.color.hex === "string") {
        run.style.color = "#" + properties.color.hex;
      }
    }

    function selectElementContents(element) {
      var range = document.createRange();
      range.selectNodeContents(element);
      var selection = window.getSelection();
      if (!selection) return;
      selection.removeAllRanges();
      selection.addRange(range);
    }

    function positionActiveTextEditor() {
      if (!activeTextEditor || !activeTextEditor.shape || !activeTextEditor.shape.bounds) return;
      var container = document.getElementById("slide-container");
      var renderedSvg = container.querySelector("svg:not(#selection-overlay)");
      if (!renderedSvg) return;
      var svgRect = renderedSvg.getBoundingClientRect();
      var containerRect = container.getBoundingClientRect();
      var viewBox = parseViewBox(renderedSvg.getAttribute("viewBox") || "0 0 960 540");
      var bounds = activeTextEditor.shape.bounds;
      var left = svgRect.left - containerRect.left + ((bounds.x - viewBox.x) / viewBox.width) * svgRect.width;
      var top = svgRect.top - containerRect.top + ((bounds.y - viewBox.y) / viewBox.height) * svgRect.height;
      var width = (bounds.width / viewBox.width) * svgRect.width;
      var height = (bounds.height / viewBox.height) * svgRect.height;
      activeTextEditor.element.style.left = left + "px";
      activeTextEditor.element.style.top = top + "px";
      activeTextEditor.element.style.width = width + "px";
      activeTextEditor.element.style.height = height + "px";
    }

    function parseViewBox(value) {
      var parts = String(value).trim().split(/\\s+/).map(Number);
      return {
        x: Number.isFinite(parts[0]) ? parts[0] : 0,
        y: Number.isFinite(parts[1]) ? parts[1] : 0,
        width: Number.isFinite(parts[2]) && parts[2] > 0 ? parts[2] : 960,
        height: Number.isFinite(parts[3]) && parts[3] > 0 ? parts[3] : 540
      };
    }

    function textEditorDocJson() {
      if (!activeTextEditor) return null;
      var original = activeTextEditor.originalDocJson;
      var paragraphs = [];
      (original.content || []).forEach(function (paragraph, paragraphIndex) {
        var paragraphElement = activeTextEditor.element.querySelector(
          '.text-editor-paragraph[data-paragraph-index="' + paragraphIndex + '"]'
        );
        var content = [];
        var emptyRunCount = 0;
        (paragraph.content || []).forEach(function (textNode, runIndex) {
          var runElement = paragraphElement
            ? paragraphElement.querySelector('.text-editor-run[data-run-index="' + runIndex + '"]')
            : null;
          var text = runElement ? runElement.textContent || "" : textNode.text || "";
          if (text.length === 0) {
            emptyRunCount += 1;
            return;
          }
          content.push({
            type: "text",
            text: text,
            marks: textNode.marks || []
          });
        });
        if (emptyRunCount > 0 && content.length > 0) {
          throw new Error("Clearing individual runs in multi-run text is unsupported.");
        }
        paragraphs.push({
          type: "paragraph",
          attrs: paragraph.attrs || {},
          content: content
        });
      });
      return { type: "doc", content: paragraphs };
    }

    function commitTextEditor() {
      if (!activeTextEditor) return Promise.resolve();
      var editor = activeTextEditor;
      if (editor.committing) return editor.commitPromise || Promise.resolve();
      var docJson = textEditorDocJson();
      if (!docJson) return Promise.resolve();
      editor.committing = true;
      editor.commitPromise = postJson("/api/editor/text-body", {
        handle: editor.shape.handle,
        docJson: docJson
      })
        .then(function (data) {
          closeTextEditor(editor);
          applyEditorResponse(data);
          setEditorMessage("Applied", false);
        })
        .catch(function (err) {
          editor.committing = false;
          editor.commitPromise = null;
          setEditorMessage(err.message, true);
          throw err;
        });
      return editor.commitPromise;
    }

    function closeTextEditor(editor) {
      var target = editor || activeTextEditor;
      if (!target) return;
      if (editor && activeTextEditor !== editor) return;
      if (!editor && target.committing) return;
      target.element.remove();
      if (activeTextEditor === target) activeTextEditor = null;
    }

    function beginDrag(kind, handle, event) {
      if (!selectedShape || !selectedShape.bounds) return;
      var overlay = document.getElementById("selection-overlay");
      if (!overlay) return;
      var point = eventPoint(overlay, event);
      if (!point) return;
      dragState = {
        kind: kind,
        handle: handle,
        pointerId: event.pointerId,
        startPoint: point,
        startBounds: {
          x: selectedShape.bounds.x,
          y: selectedShape.bounds.y,
          width: selectedShape.bounds.width,
          height: selectedShape.bounds.height
        }
      };
      try {
        overlay.setPointerCapture(event.pointerId);
      } catch (_error) {
        // Pointer capture can be unavailable after a fast re-render; window listeners still cover drag.
      }
      window.addEventListener("pointermove", updateDrag);
      window.addEventListener("pointerup", finishDrag);
      window.addEventListener("pointercancel", cancelDrag);
    }

    function updateDrag(event) {
      if (!dragState || event.pointerId !== dragState.pointerId || !selectedShape) return;
      var overlay = document.getElementById("selection-overlay");
      if (!overlay) return;
      var point = eventPoint(overlay, event);
      if (!point) return;
      var dx = point.x - dragState.startPoint.x;
      var dy = point.y - dragState.startPoint.y;
      selectedShape.bounds =
        dragState.kind === "move"
          ? movedBounds(dragState.startBounds, dx, dy)
          : resizedBounds(dragState.startBounds, dragState.handle, dx, dy);
      renderSelectionOverlay();
    }

    function finishDrag(event) {
      if (!dragState || event.pointerId !== dragState.pointerId || !selectedShape) return;
      var nextBounds = selectedShape.bounds;
      var previousDrag = dragState;
      dragState = null;
      detachDragListeners();
      applyShapeBoundsEdit(previousDrag.startBounds, nextBounds).catch(function (err) {
        setEditorMessage(err.message, true);
        selectedShape.bounds = previousDrag.startBounds;
        renderSelectionOverlay();
      });
    }

    function cancelDrag(event) {
      if (!dragState || event.pointerId !== dragState.pointerId || !selectedShape) return;
      selectedShape.bounds = dragState.startBounds;
      dragState = null;
      detachDragListeners();
      renderSelectionOverlay();
    }

    function detachDragListeners() {
      window.removeEventListener("pointermove", updateDrag);
      window.removeEventListener("pointerup", finishDrag);
      window.removeEventListener("pointercancel", cancelDrag);
    }

    function eventPoint(svg, event) {
      var matrix = svg.getScreenCTM();
      if (!matrix) return null;
      var point = svg.createSVGPoint();
      point.x = event.clientX;
      point.y = event.clientY;
      return point.matrixTransform(matrix.inverse());
    }

    function movedBounds(bounds, dx, dy) {
      return {
        x: bounds.x + dx,
        y: bounds.y + dy,
        width: bounds.width,
        height: bounds.height
      };
    }

    function resizedBounds(bounds, handle, dx, dy) {
      var minSize = 8;
      var right = bounds.x + bounds.width;
      var bottom = bounds.y + bounds.height;
      var next = {
        x: bounds.x,
        y: bounds.y,
        width: bounds.width,
        height: bounds.height
      };
      if (handle === "nw" || handle === "sw") {
        next.x = Math.min(bounds.x + dx, right - minSize);
        next.width = right - next.x;
      }
      if (handle === "ne" || handle === "se") {
        next.width = Math.max(minSize, bounds.width + dx);
      }
      if (handle === "nw" || handle === "ne") {
        next.y = Math.min(bounds.y + dy, bottom - minSize);
        next.height = bottom - next.y;
      }
      if (handle === "sw" || handle === "se") {
        next.height = Math.max(minSize, bounds.height + dy);
      }
      return next;
    }

    function applyShapeBoundsEdit(startBounds, nextBounds) {
      if (!selectedShape || !selectedShape.handle) return Promise.resolve();
      var handle = selectedShape.handle;
      var changed =
        Math.round(startBounds.x) !== Math.round(nextBounds.x) ||
        Math.round(startBounds.y) !== Math.round(nextBounds.y) ||
        Math.round(startBounds.width) !== Math.round(nextBounds.width) ||
        Math.round(startBounds.height) !== Math.round(nextBounds.height);
      if (!changed) return Promise.resolve();
      return postJson("/api/editor/command", {
        command: {
          kind: "setShapeTransform",
          handle: handle,
          offsetX: Math.round(nextBounds.x * EMU_PER_PIXEL),
          offsetY: Math.round(nextBounds.y * EMU_PER_PIXEL),
          width: Math.round(nextBounds.width * EMU_PER_PIXEL),
          height: Math.round(nextBounds.height * EMU_PER_PIXEL)
        }
      }).then(function (data) {
        if (data) {
          applyEditorResponse(data);
          setEditorMessage("Applied", false);
        }
      });
    }

    function duplicateSlide(index) {
      if (activeTextEditor) {
        setEditorMessage("Finish text editing before slide operations", true);
        return;
      }
      var slide = slides[index];
      if (!slide || !slide.handle) {
        setEditorMessage("Slide handle is unavailable", true);
        return;
      }
      postJson("/api/editor/command", {
        command: {
          kind: "duplicateSlide",
          handle: slide.handle
        }
      })
        .then(function (data) {
          applyEditorResponse(data, index + 1);
          setEditorMessage("Duplicated slide", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function deleteSlide(index) {
      if (activeTextEditor) {
        setEditorMessage("Finish text editing before slide operations", true);
        return;
      }
      if (slideCount <= 1) {
        setEditorMessage("Cannot delete the last slide", true);
        return;
      }
      var slide = slides[index];
      if (!slide || !slide.handle) {
        setEditorMessage("Slide handle is unavailable", true);
        return;
      }
      postJson("/api/editor/command", {
        command: {
          kind: "deleteSlide",
          handle: slide.handle
        }
      })
        .then(function (data) {
          applyEditorResponse(data, Math.min(index, (data.slides || slides).length - 1));
          setEditorMessage("Deleted slide", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function postJson(url, payload) {
      return fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload || {})
      }).then(function (res) {
        return res.json().then(function (data) {
          if (!res.ok) throw new Error(data.error || "request failed");
          return data;
        });
      });
    }

    function escapeHtmlClient(value) {
      return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
    }

    renderThumbnails();
    updateHistory(editorHistory);
    loadShapeOptions(1);

    document.getElementById("text-run-select").addEventListener("change", syncTextInput);
    document.getElementById("apply-text-button").addEventListener("click", function () {
      var option = textRunOptions[Number(document.getElementById("text-run-select").value || "0")];
      if (!option) return;
      postJson("/api/editor/command", {
        command: {
          kind: "replaceTextRunPlainText",
          handle: option.handle,
          text: document.getElementById("text-run-input").value
        }
      })
        .then(function (data) {
          applyEditorResponse(data);
          setEditorMessage("Applied", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    });
    document.getElementById("undo-button").addEventListener("click", function () {
      postJson("/api/editor/undo")
        .then(function (data) {
          applyEditorResponse(data);
          setEditorMessage("Undone", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    });
    document.getElementById("redo-button").addEventListener("click", function () {
      postJson("/api/editor/redo")
        .then(function (data) {
          applyEditorResponse(data);
          setEditorMessage("Redone", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    });
    document.getElementById("save-button").addEventListener("click", function () {
      postJson("/api/editor/save")
        .then(function (data) {
          updateHistory(data.history);
          setEditorMessage("Saved: " + data.path, false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    });

    // WebSocket for live reload
    function connect() {
      var ws = new WebSocket("ws://" + location.host);
      var status = document.getElementById("status");

      ws.onopen = function () {
        status.textContent = "Connected";
        status.className = "";
      };
      ws.onclose = function () {
        status.textContent = "Disconnected - reconnecting...";
        status.className = "error";
        setTimeout(connect, 2000);
      };
      ws.onmessage = function (event) {
        var data = JSON.parse(event.data);
        if (data.type === "rendering") {
          status.textContent = "Re-rendering...";
          status.className = "rendering";
        } else if (data.type === "reload") {
          status.textContent = "Updating...";
          status.className = "rendering";
          location.reload();
        } else if (data.type === "error") {
          status.textContent = "Error: " + data.message;
          status.className = "error";
        }
      };
    }
    connect();

    // Initial: resize the main SVG
    (function () {
      var svg = document.querySelector("#slide-container svg");
      if (svg) {
        svg.removeAttribute("width");
        svg.removeAttribute("height");
        svg.style.width = "100%";
        svg.style.height = "auto";
      }
    })();

    // Keyboard navigation
    document.addEventListener("keydown", function (e) {
      if (e.key === "ArrowLeft" && currentIndex > 0) selectSlide(currentIndex - 1);
      if (e.key === "ArrowRight" && currentIndex < slideCount - 1)
        selectSlide(currentIndex + 1);
    });
  </script>
</body>
</html>`;
}

function generateThumbnailsHtml(slides: SlideSvg[]): string {
  return slides
    .map(
      (s, i) =>
        `<div class="thumbnail${i === 0 ? " active" : ""}" data-index="${i}">` +
        `<div class="thumb-label"><span class="thumb-title">Slide ${String(s.slideNumber)}</span>` +
        `<span class="thumb-actions">` +
        `<button class="thumb-action" data-testid="duplicate-slide-${String(i)}" data-action="duplicate" data-index="${String(i)}" type="button" title="Duplicate slide">D</button>` +
        `<button class="thumb-action" data-testid="delete-slide-${String(i)}" data-action="delete" data-index="${String(i)}" type="button" title="Delete slide"${slides.length <= 1 ? " disabled" : ""}>X</button>` +
        `</span></div>` +
        `<div class="thumb-svg">${s.svg}</div>` +
        `</div>`,
    )
    .join("");
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function jsonForScript(value: unknown): string {
  return JSON.stringify(value).replace(/</g, "\\u003c");
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
      if (raw.length > 1024 * 1024) {
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
  throw new Error("unsupported command kind");
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
