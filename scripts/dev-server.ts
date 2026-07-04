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
  asRelationshipId,
  asSourceNodeId,
  readPptx,
  type SourceHandle,
  type SourceShapeNode,
  type SourceTextRun,
  writePptx,
} from "../packages/document/src/index.js";
import { createEditorSession, type EditorCommand } from "../packages/editor-core/src/index.js";
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

interface EditorShapeInfo {
  id: string;
  kind: SourceShapeNode["kind"];
  name?: string;
  handle?: SourceHandle;
  bounds?: ShapeBoundsPx;
  editableTransform?: boolean;
  textRuns?: EditorTextRunInfo[];
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
    this.#slides = await this.#render(input);
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
      font-size: 10px;
      color: #888;
      text-align: center;
      padding: 2px 0;
      background: #16213e;
    }
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
    var shapeRequestId = 0;
    var EMU_PER_PIXEL = ${String(EMU_PER_INCH / DEFAULT_DPI)};

    function selectSlide(index) {
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
            '<div class="thumb-label">Slide ' + s.slideNumber + '</div>' +
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

    function applyEditorResponse(data) {
      slides = data.slides || slides;
      slideCount = slides.length;
      updateHistory(data.history);
      if (data.selection && data.selection.shapeHandle) {
        selectedShapeKey = handleKey(data.selection.shapeHandle);
      }
      renderThumbnails();
      selectSlide(Math.min(currentIndex, Math.max(slides.length - 1, 0)));
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
          event.preventDefault();
          selectShape(shape, event);
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
      selectedShapeKey = shapeKey(shape);
      selectedShape = cloneShape(shape);
      renderSelectionOverlay();
      beginDrag("move", null, event);
      postJson("/api/editor/select", { handle: shape.handle })
        .then(function (data) {
          updateHistory(data.history);
          if (data.selection && data.selection.shapeHandle) {
            selectedShapeKey = handleKey(data.selection.shapeHandle);
          }
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
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
        `<div class="thumb-label">Slide ${String(s.slideNumber)}</div>` +
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
  throw new Error("unsupported command kind");
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
