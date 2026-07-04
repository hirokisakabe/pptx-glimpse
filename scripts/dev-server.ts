import { execFile } from "child_process";
import { watch } from "fs";
import { mkdtemp, readFile, rm, writeFile } from "fs/promises";
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
  textRuns?: EditorTextRunInfo[];
}

interface EditorSlidesResponse {
  slides: SlideSvg[];
  history: EditorHistoryState;
}

interface EditorSaveResponse {
  ok: true;
  path: string;
  history: EditorHistoryState;
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

  response(): EditorSlidesResponse {
    return { slides: this.#slides, history: this.history };
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
    const result = this.#session.apply(command);
    if (!result.ok) {
      throw new Error(result.message);
    }
    this.#dirty = true;
    await this.renderCurrentSlides();
    return this.response();
  }

  async undo(): Promise<EditorSlidesResponse> {
    const result = this.#session.undo();
    if (!result.ok) {
      throw new Error(result.reason);
    }
    this.#dirty = this.#session.undoDepth > 0;
    await this.renderCurrentSlides();
    return this.response();
  }

  async redo(): Promise<EditorSlidesResponse> {
    const result = this.#session.redo();
    if (!result.ok) {
      throw new Error(result.reason);
    }
    this.#dirty = this.#session.undoDepth > 0;
    await this.renderCurrentSlides();
    return this.response();
  }

  async save(outputPath?: string): Promise<EditorSaveResponse> {
    const path = resolveSavePath(this.sourcePath, outputPath);
    const output = writePptx(this.#session.document);
    readPptx(output);
    await writeFile(path, output);
    return { ok: true, path, history: this.history };
  }
}

function defaultEditedPath(sourcePath: string): string {
  const extension = extname(sourcePath);
  const base = basename(sourcePath, extension);
  return join(dirname(sourcePath), `${base}.edited${extension || ".pptx"}`);
}

function resolveSavePath(sourcePath: string, outputPath?: string): string {
  const sourceDir = dirname(sourcePath);
  const path =
    outputPath === undefined
      ? defaultEditedPath(sourcePath)
      : resolve(isAbsolute(outputPath) ? outputPath : join(sourceDir, outputPath));
  const relativePath = relative(sourceDir, path);
  if (relativePath.startsWith("..") || isAbsolute(relativePath)) {
    throw new Error("save path must be inside the source PPTX directory");
  }
  if (extname(path).toLowerCase() !== ".pptx") {
    throw new Error("save path must use the .pptx extension");
  }
  return path;
}

function shapeInfo(shape: SourceShapeNode, index: number): EditorShapeInfo[] {
  const base: EditorShapeInfo = {
    id: String(shape.nodeId ?? shape.handle?.nodeId ?? `${shape.kind}:${String(index)}`),
    kind: shape.kind,
    ...(shapeName(shape) !== undefined ? { name: shapeName(shape) } : {}),
    ...(shape.handle !== undefined ? { handle: shape.handle } : {}),
    ...(shape.kind !== "raw" && shape.transform !== undefined
      ? { bounds: transformBoundsPx(shape.transform) }
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
  return [base, ...shape.children.flatMap((child, childIndex) => shapeInfo(child, childIndex))];
}

function shapeName(shape: SourceShapeNode): string | undefined {
  return "name" in shape ? shape.name : undefined;
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
    }
    #slide-container svg { display: block; width: 100%; height: auto; }
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
    var shapeRequestId = 0;

    function selectSlide(index) {
      currentIndex = index;
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
      renderThumbnails();
      selectSlide(Math.min(currentIndex, Math.max(slides.length - 1, 0)));
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
