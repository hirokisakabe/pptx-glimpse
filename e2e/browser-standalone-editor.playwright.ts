import { Buffer } from "node:buffer";
import { execFile } from "node:child_process";
import { mkdir, mkdtemp, readFile, rm, symlink, writeFile } from "node:fs/promises";
import { createServer, type Server } from "node:http";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { expect, type Page, test } from "@playwright/test";
import { build } from "esbuild";
import JSZip from "jszip";

import {
  type PptxSourceModel,
  readPptx,
  type SourceShape,
} from "../packages/document/src/index.js";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const corePackageRoot = resolve(repoRoot, "packages/core");
const documentPackageRoot = resolve(repoRoot, "packages/document");
const editorCorePackageRoot = resolve(repoRoot, "packages/editor-core");
const rendererPackageRoot = resolve(repoRoot, "packages/renderer");
const execFileAsync = promisify(execFile);
const encoder = new TextEncoder();
const EMU_PER_PIXEL = 9525;
const RED_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEUlEQVR4nGP8z4AATEhsPBwAM9EBBzDn4UwAAAAASUVORK5CYII=";
const BLUE_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAE0lEQVR4nGNkYPjPAANMcBZeDgAx0wEH1s7nlgAAAABJRU5ErkJggg==";
const RED_PNG = pngBytes(RED_PNG_BASE64);
const BLUE_PNG = pngBytes(BLUE_PNG_BASE64);
let coreDistBuildPromise: Promise<void> | null = null;

test("runs a browser-only editor move, resize, text, undo, redo, download, and reopen flow", async ({
  page,
}) => {
  const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-browser-editor-test-"));
  const editor = await startStandaloneEditor();
  try {
    const sourcePath = join(dir, "fixture.pptx");
    const savedPath = join(dir, "fixture.edited.pptx");
    await writeFile(sourcePath, await buildShapeFixture());

    await page.goto(editor.url);
    await page.getByTestId("pptx-input").setInputFiles(sourcePath);
    await expect(page.getByTestId("status")).toContainText("1 slide ready");
    await expect(page.getByTestId("shape-hit-area").first()).toBeVisible();

    await clickSvgPoint(page, 240, 240);
    await expect(page.getByTestId("selection-box")).toHaveAttribute("x", "96");
    await expect(page.getByTestId("selection-handle-se")).toBeVisible();
    await expect(page.getByTestId("image-replacement-input")).toBeDisabled();

    await dragSvgPoint(page, { x: 240, y: 240 }, { x: 264, y: 256 });
    await expect(page.getByTestId("selection-box")).toHaveAttribute("x", "120");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("y", "208");

    await dragSvgPoint(page, { x: 408, y: 304 }, { x: 456, y: 328 });
    await expect(page.getByTestId("selection-box")).toHaveAttribute("width", "336");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("height", "120");

    await page.getByRole("button", { name: "Undo" }).click();
    await expect(page.getByTestId("selection-box")).toHaveAttribute("width", "288");
    await page.getByRole("button", { name: "Redo" }).click();
    await expect(page.getByTestId("selection-box")).toHaveAttribute("width", "336");

    expect(await page.evaluate(() => window.__pptxGlimpseEditorSmoke?.hasEditableTextBody)).toBe(
      true,
    );
    await dblclickSvgPoint(page, 240, 240);
    const overlay = page.getByTestId("text-editor-overlay");
    await expect(overlay).toBeVisible();
    await expect(overlay).toContainText("Original");
    await expectOverlayNearSvgBounds(page, { x: 120, y: 208, width: 336, height: 120 });

    await page.getByTestId("text-editor-run").first().fill("Browser edited");
    await page.getByTestId("text-editor-done").click();
    await expect(page.locator("#slide-container")).toContainText("Browser edited");

    const download = page.waitForEvent("download");
    await page.getByRole("button", { name: "Download" }).click();
    await (await download).saveAs(savedPath);

    const saved = readPptx(await readFile(savedPath));
    expect(firstShape(saved).transform).toMatchObject({
      offsetX: 120 * EMU_PER_PIXEL,
      offsetY: 208 * EMU_PER_PIXEL,
      width: 336 * EMU_PER_PIXEL,
      height: 120 * EMU_PER_PIXEL,
    });
    expect(firstText(saved)).toBe("Browser edited");

    await page.getByTestId("pptx-input").setInputFiles(savedPath);
    await expect(page.getByTestId("status")).toContainText("1 slide ready");
    await expect(page.locator("#slide-container")).toContainText("Browser edited");
  } finally {
    await editor.close();
    await rm(dir, { recursive: true, force: true });
  }
});

test("replaces a selected image in the browser-only editor with warning, reject, undo, and redo", async ({
  page,
}) => {
  const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-browser-editor-image-test-"));
  const editor = await startStandaloneEditor();
  try {
    const sourcePath = join(dir, "fixture.pptx");
    const savedPath = join(dir, "fixture.edited.pptx");
    const replacementPath = join(dir, "replacement.png");
    const wrongFormatPath = join(dir, "wrong-format.jpg");
    await writeFile(sourcePath, await buildImageFixture());
    await writeFile(replacementPath, BLUE_PNG);
    await writeFile(wrongFormatPath, new Uint8Array([0xff, 0xd8, 0xff, 0xd9]));

    await page.goto(editor.url);
    await page.getByTestId("pptx-input").setInputFiles(sourcePath);
    await expect(page.getByTestId("status")).toContainText("1 slide ready");
    await expect(page.getByTestId("image-replacement-input")).toBeDisabled();
    await expectImageHref(page, RED_PNG_BASE64);

    await clickSvgPoint(page, 120, 120);
    await expect(page.getByTestId("image-replacement-input")).toBeEnabled();
    await expect(page.getByTestId("image-replacement-input")).toHaveAttribute(
      "accept",
      /image\/png/,
    );

    await page.getByTestId("image-replacement-input").setInputFiles(wrongFormatPath);
    await expect(page.getByTestId("message")).toContainText("does not match existing media");
    await expectImageHref(page, RED_PNG_BASE64);

    await page.getByTestId("image-replacement-input").setInputFiles(replacementPath);
    await expect(page.getByTestId("message")).toContainText("shared media part affects 2 pictures");
    await expectImageHref(page, BLUE_PNG_BASE64);

    await page.getByRole("button", { name: "Undo" }).click();
    await expectImageHref(page, RED_PNG_BASE64);
    await page.getByRole("button", { name: "Redo" }).click();
    await expectImageHref(page, BLUE_PNG_BASE64);

    const download = page.waitForEvent("download");
    await page.getByRole("button", { name: "Download" }).click();
    await (await download).saveAs(savedPath);
    expect(mediaBytes(readPptx(await readFile(savedPath)), "ppt/media/image1.png")).toEqual(
      BLUE_PNG,
    );
  } finally {
    await editor.close();
    await rm(dir, { recursive: true, force: true });
  }
});

async function clickSvgPoint(page: Page, x: number, y: number) {
  const point = await svgPointToClient(page, x, y);
  await page.mouse.click(point.x, point.y);
}

async function dblclickSvgPoint(page: Page, x: number, y: number) {
  const point = await svgPointToClient(page, x, y);
  await page.mouse.dblclick(point.x, point.y);
}

async function dragSvgPoint(
  page: Page,
  from: { x: number; y: number },
  to: { x: number; y: number },
) {
  const start = await svgPointToClient(page, from.x, from.y);
  const end = await svgPointToClient(page, to.x, to.y);
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(end.x, end.y, { steps: 6 });
  await page.mouse.up();
}

async function svgPointToClient(
  page: Page,
  x: number,
  y: number,
): Promise<{ x: number; y: number }> {
  return page.evaluate(
    ({ svgX, svgY }) => {
      const overlay = document.getElementById("selection-overlay");
      if (!(overlay instanceof SVGSVGElement)) throw new Error("selection overlay not found");
      const point = overlay.createSVGPoint();
      point.x = svgX;
      point.y = svgY;
      const matrix = overlay.getScreenCTM();
      if (matrix === null) throw new Error("selection overlay matrix not found");
      const screenPoint = point.matrixTransform(matrix);
      return { x: screenPoint.x, y: screenPoint.y };
    },
    { svgX: x, svgY: y },
  );
}

async function expectOverlayNearSvgBounds(
  page: Page,
  bounds: { x: number; y: number; width: number; height: number },
) {
  const rects = await page.evaluate((expected) => {
    const overlay = document.querySelector('[data-testid="text-editor-overlay"]');
    const selectionOverlay = document.getElementById("selection-overlay");
    if (!(overlay instanceof HTMLElement)) throw new Error("text editor overlay not found");
    if (!(selectionOverlay instanceof SVGSVGElement))
      throw new Error("selection overlay not found");
    const point = selectionOverlay.createSVGPoint();
    const matrix = selectionOverlay.getScreenCTM();
    if (matrix === null) throw new Error("selection overlay matrix not found");
    point.x = expected.x;
    point.y = expected.y;
    const topLeft = point.matrixTransform(matrix);
    point.x = expected.x + expected.width;
    point.y = expected.y + expected.height;
    const bottomRight = point.matrixTransform(matrix);
    const overlayRect = overlay.getBoundingClientRect();
    return {
      expected: {
        left: topLeft.x,
        top: topLeft.y,
        width: bottomRight.x - topLeft.x,
        height: bottomRight.y - topLeft.y,
      },
      actual: {
        left: overlayRect.left,
        top: overlayRect.top,
        width: overlayRect.width,
        height: overlayRect.height,
      },
    };
  }, bounds);

  expect(Math.abs(rects.actual.left - rects.expected.left)).toBeLessThan(3);
  expect(Math.abs(rects.actual.top - rects.expected.top)).toBeLessThan(3);
  expect(Math.abs(rects.actual.width - rects.expected.width)).toBeLessThan(3);
  expect(Math.abs(rects.actual.height - rects.expected.height)).toBeLessThan(3);
}

async function expectImageHref(page: Page, base64: string) {
  await expect
    .poll(async () => page.locator("#slide-container image").first().getAttribute("href"))
    .toContain(base64);
}

interface EditorServer {
  readonly url: string;
  close(): Promise<void>;
}

async function startStandaloneEditor(): Promise<EditorServer> {
  const appBundle = await buildStandaloneEditorBundle();
  const server = createServer((request, response) => {
    const url = new URL(request.url ?? "/", "http://127.0.0.1");
    if (url.pathname === "/" || url.pathname === "/index.html") {
      response.setHeader("Content-Type", "text/html; charset=utf-8");
      response.end(editorHtml);
      return;
    }
    if (url.pathname === "/app.js") {
      response.setHeader("Content-Type", "text/javascript; charset=utf-8");
      response.end(appBundle);
      return;
    }
    response.statusCode = 404;
    response.end("Not found");
  });
  const url = await listen(server);
  return { url, close: () => closeServer(server) };
}

async function buildStandaloneEditorBundle(): Promise<string> {
  await ensureCoreDist();
  const packageResolveDir = await createPackageResolveDir();
  try {
    const result = await build({
      stdin: {
        contents: editorAppSource,
        resolveDir: packageResolveDir,
        sourcefile: "browser-standalone-editor.ts",
        loader: "ts",
      },
      bundle: true,
      write: false,
      format: "esm",
      platform: "browser",
      conditions: ["browser", "import"],
      logLevel: "silent",
      absWorkingDir: repoRoot,
    });
    const bundled = result.outputFiles[0].text;
    expect(bundled).not.toMatch(
      /(?:node:fs|node:path|node:os|node:buffer|fs\/promises|from "fs"|from "path"|from "os"|from "module")/,
    );
    return bundled;
  } finally {
    await rm(packageResolveDir, { recursive: true, force: true });
  }
}

async function ensureCoreDist(): Promise<void> {
  coreDistBuildPromise ??= (async () => {
    await runPackageBuild(documentPackageRoot);
    await runPackageBuild(editorCorePackageRoot);
    await runPackageBuild(rendererPackageRoot);
    await runPackageBuild(corePackageRoot);
  })();
  await coreDistBuildPromise;
}

async function runPackageBuild(packageRoot: string): Promise<void> {
  await execFileAsync(resolve(repoRoot, "node_modules/.bin/tsup"), ["--config", "tsup.config.ts"], {
    cwd: packageRoot,
    maxBuffer: 10 * 1024 * 1024,
  });
}

async function createPackageResolveDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-browser-editor-resolve-"));
  const nodeModules = join(dir, "node_modules");
  await mkdir(nodeModules);
  await symlink(corePackageRoot, join(nodeModules, "pptx-glimpse"), "dir");
  return dir;
}

const editorHtml = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>pptx-glimpse browser editor smoke</title>
    <style>
      * { box-sizing: border-box; }
      body { margin: 0; min-height: 100vh; background: #20242a; color: #e5e7eb; font-family: ui-sans-serif, system-ui, sans-serif; }
      main { display: grid; grid-template-columns: 1fr 220px; gap: 16px; height: 100vh; padding: 16px; }
      #workspace { display: grid; place-items: center; min-width: 0; overflow: auto; background: #2d333b; }
      #slide-container { position: relative; width: min(960px, 100%); background: #fff; box-shadow: 0 10px 24px rgba(0, 0, 0, 0.28); }
      #slide-container svg:not(#selection-overlay) { display: block; width: 100%; height: auto; }
      #selection-overlay { position: absolute; inset: 0; width: 100%; height: 100%; overflow: visible; touch-action: none; z-index: 2; }
      #panel { display: flex; flex-direction: column; gap: 10px; padding: 12px; background: #111827; border-left: 1px solid #384252; }
      button, input { min-height: 32px; border: 1px solid #4b5563; border-radius: 4px; background: #1f2937; color: #f9fafb; font: inherit; }
      button { cursor: pointer; font-weight: 650; }
      button:disabled { opacity: 0.45; cursor: not-allowed; }
      #status, #message { min-height: 18px; font-size: 12px; color: #a7f3d0; overflow-wrap: anywhere; }
      .shape-hit-area { cursor: move; fill: transparent; pointer-events: all; }
      .selection-box { fill: none; stroke: #0f766e; stroke-width: 1.5; vector-effect: non-scaling-stroke; pointer-events: none; }
      .selection-handle { fill: #f8fafc; stroke: #0f766e; stroke-width: 1.5; cursor: nwse-resize; vector-effect: non-scaling-stroke; pointer-events: all; }
      #text-editor-overlay { position: absolute; min-width: 40px; min-height: 28px; padding: 4px; border: 1px solid #2563eb; background: rgba(255, 255, 255, 0.96); color: #111827; z-index: 3; overflow: hidden; }
      .text-editor-run { outline: none; white-space: pre-wrap; }
      .text-editor-actions { position: absolute; right: 4px; bottom: 4px; }
      .text-editor-actions button { min-height: 24px; color: #111827; background: #f8fafc; }
    </style>
  </head>
  <body>
    <main>
      <section id="workspace"><div id="slide-container"></div></section>
      <aside id="panel">
        <input data-testid="pptx-input" id="pptx-input" type="file" accept=".pptx">
        <button id="undo-button" type="button" disabled>Undo</button>
        <button id="redo-button" type="button" disabled>Redo</button>
        <input data-testid="image-replacement-input" id="image-replacement-input" type="file" disabled>
        <button id="download-button" type="button" disabled>Download</button>
        <div data-testid="status" id="status">Ready</div>
        <div data-testid="message" id="message"></div>
      </aside>
    </main>
    <script type="module" src="/app.js"></script>
  </body>
</html>`;

const editorAppSource = `
import { createBrowserPptxEditorSession } from "pptx-glimpse";

const pptxInput = document.getElementById("pptx-input");
const status = document.getElementById("status");
const message = document.getElementById("message");
const container = document.getElementById("slide-container");
const undoButton = document.getElementById("undo-button");
const redoButton = document.getElementById("redo-button");
const downloadButton = document.getElementById("download-button");
const imageReplacementInput = document.getElementById("image-replacement-input");
const EMU_PER_PIXEL = 9525;

let editor = null;
let fileName = "edited.pptx";
let shapes = [];
let selectedShape = null;
let selectedShapeKey = null;
let dragState = null;
let activeTextEditor = null;

pptxInput.addEventListener("change", async () => {
  const file = pptxInput.files?.[0];
  if (!file) return;
  fileName = file.name.replace(/\\.pptx$/i, ".edited.pptx");
  status.textContent = "Opening";
  editor = await createBrowserPptxEditorSession(new Uint8Array(await file.arrayBuffer()), {
    skipSystemFonts: true,
    textOutput: "text",
  });
  render();
  status.textContent = editor.slides.length + " slide" + (editor.slides.length === 1 ? "" : "s") + " ready";
  downloadButton.disabled = false;
});

undoButton.addEventListener("click", async () => {
  if (!editor) return;
  await editor.undo();
  render();
  message.textContent = "Undone";
});

redoButton.addEventListener("click", async () => {
  if (!editor) return;
  await editor.redo();
  render();
  message.textContent = "Redone";
});

imageReplacementInput.addEventListener("change", async () => {
  const file = imageReplacementInput.files?.[0];
  if (!file) return;
  try {
    await replaceSelectedImage(file);
  } catch (error) {
    message.textContent = error instanceof Error ? error.message : String(error);
  } finally {
    syncImageReplacementInput();
  }
});

downloadButton.addEventListener("click", () => {
  if (!editor) return;
  const saved = editor.save();
  updateHistory(saved.history);
  const href = URL.createObjectURL(new Blob([saved.pptx], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  }));
  const link = document.createElement("a");
  link.href = href;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(href);
  message.textContent = "Downloaded";
});

function render() {
  if (!editor) return;
  const slide = editor.slides[0];
  container.innerHTML = slide ? slide.svg : "";
  const svg = container.querySelector("svg");
  if (svg) {
    svg.removeAttribute("width");
    svg.removeAttribute("height");
    svg.style.width = "100%";
    svg.style.height = "auto";
  }
  shapes = editor.shapes(1).filter(
    (shape) => shape.handle && shape.bounds && (shape.editableTransform || shape.editableImageReplacement),
  );
  window.__pptxGlimpseEditorSmoke = {
    hasEditableTextBody: shapes.some((shape) => Boolean(shape.editableTextBody)),
  };
  selectedShape = selectedShapeKey
    ? shapes.find((shape) => shapeKey(shape) === selectedShapeKey) || null
    : null;
  updateHistory(editor.history);
  syncImageReplacementInput();
  renderSelectionOverlay();
}

function updateHistory(history) {
  undoButton.disabled = !history.canUndo;
  redoButton.disabled = !history.canRedo;
}

function renderSelectionOverlay() {
  const renderedSvg = container.querySelector("svg:not(#selection-overlay)");
  const existing = document.getElementById("selection-overlay");
  if (existing) existing.remove();
  if (!renderedSvg) return;

  const overlay = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  overlay.id = "selection-overlay";
  overlay.setAttribute("data-testid", "selection-overlay");
  overlay.setAttribute("viewBox", renderedSvg.getAttribute("viewBox") || "0 0 960 540");

  shapes.forEach((shape, index) => {
    const hit = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    hit.setAttribute("class", "shape-hit-area");
    hit.setAttribute("data-testid", "shape-hit-area");
    hit.setAttribute("data-shape-index", String(index));
    setRectAttributes(hit, shape.bounds);
    hit.addEventListener("pointerdown", (event) => selectShape(shape, event));
    hit.addEventListener("mousedown", (event) => {
      if (event.detail >= 2 && shape.editableTextBody) {
        event.preventDefault();
        openTextEditor(shape);
      }
    });
    hit.addEventListener("dblclick", (event) => {
      event.preventDefault();
      event.stopPropagation();
      openTextEditor(shape);
    });
    overlay.appendChild(hit);
  });

  if (selectedShape?.bounds) {
    const box = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    box.setAttribute("class", "selection-box");
    box.setAttribute("data-testid", "selection-box");
    setRectAttributes(box, selectedShape.bounds);
    overlay.appendChild(box);

    ["nw", "ne", "sw", "se"].forEach((handle) => {
      const point = handlePoint(selectedShape.bounds, handle);
      const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      rect.setAttribute("class", "selection-handle");
      rect.setAttribute("data-testid", "selection-handle-" + handle);
      rect.setAttribute("x", String(point.x - 4));
      rect.setAttribute("y", String(point.y - 4));
      rect.setAttribute("width", "8");
      rect.setAttribute("height", "8");
      rect.addEventListener("pointerdown", (event) => {
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

function selectShape(shape, event) {
  selectedShape = cloneShape(shape);
  selectedShapeKey = shapeKey(shape);
  editor.selectShape(shape.handle);
  syncImageReplacementInput();
  beginDrag("move", null, event);
  window.setTimeout(renderSelectionOverlay, 0);
}

function beginDrag(kind, handle, event) {
  if (!selectedShape?.bounds) return;
  const overlay = document.getElementById("selection-overlay");
  const point = eventPoint(overlay, event);
  if (!point) return;
  dragState = {
    kind,
    handle,
    pointerId: event.pointerId,
    startPoint: point,
    startBounds: { ...selectedShape.bounds },
  };
  window.addEventListener("pointermove", updateDrag);
  window.addEventListener("pointerup", finishDrag);
}

function updateDrag(event) {
  if (!dragState || event.pointerId !== dragState.pointerId || !selectedShape) return;
  const point = eventPoint(document.getElementById("selection-overlay"), event);
  if (!point) return;
  const dx = point.x - dragState.startPoint.x;
  const dy = point.y - dragState.startPoint.y;
  selectedShape.bounds =
    dragState.kind === "move"
      ? movedBounds(dragState.startBounds, dx, dy)
      : resizedBounds(dragState.startBounds, dragState.handle, dx, dy);
  renderSelectionOverlay();
}

async function finishDrag(event) {
  if (!dragState || event.pointerId !== dragState.pointerId || !selectedShape) return;
  const startBounds = dragState.startBounds;
  const nextBounds = selectedShape.bounds;
  dragState = null;
  window.removeEventListener("pointermove", updateDrag);
  window.removeEventListener("pointerup", finishDrag);
  if (
    Math.round(startBounds.x) === Math.round(nextBounds.x) &&
    Math.round(startBounds.y) === Math.round(nextBounds.y) &&
    Math.round(startBounds.width) === Math.round(nextBounds.width) &&
    Math.round(startBounds.height) === Math.round(nextBounds.height)
  ) {
    return;
  }
  await editor.apply({
    kind: "setShapeTransform",
    handle: selectedShape.handle,
    offsetX: Math.round(nextBounds.x * EMU_PER_PIXEL),
    offsetY: Math.round(nextBounds.y * EMU_PER_PIXEL),
    width: Math.round(nextBounds.width * EMU_PER_PIXEL),
    height: Math.round(nextBounds.height * EMU_PER_PIXEL),
  });
  render();
  message.textContent = "Applied";
}

function openTextEditor(shape) {
  if (!shape.editableTextBody || !shape.bounds) return;
  closeTextEditor();
  selectedShape = cloneShape(shape);
  selectedShapeKey = shapeKey(shape);
  renderSelectionOverlay();

  const overlay = document.createElement("div");
  overlay.id = "text-editor-overlay";
  overlay.setAttribute("data-testid", "text-editor-overlay");
  overlay.appendChild(createTextEditorContent(shape.editableTextBody.docJson));
  const actions = document.createElement("div");
  actions.className = "text-editor-actions";
  const done = document.createElement("button");
  done.type = "button";
  done.textContent = "Done";
  done.setAttribute("data-testid", "text-editor-done");
  done.addEventListener("click", () => commitTextEditor());
  actions.appendChild(done);
  overlay.appendChild(actions);
  container.appendChild(overlay);
  activeTextEditor = { element: overlay, shape: cloneShape(shape), originalDocJson: shape.editableTextBody.docJson };
  positionActiveTextEditor();
  const firstRun = overlay.querySelector(".text-editor-run");
  if (firstRun) firstRun.focus();
}

function createTextEditorContent(docJson) {
  const body = document.createElement("div");
  (docJson.content || []).forEach((paragraph, paragraphIndex) => {
    const paragraphElement = document.createElement("div");
    paragraphElement.dataset.paragraphIndex = String(paragraphIndex);
    (paragraph.content || []).forEach((textNode, runIndex) => {
      const run = document.createElement("span");
      run.className = "text-editor-run";
      run.setAttribute("data-testid", "text-editor-run");
      run.contentEditable = "true";
      run.dataset.paragraphIndex = String(paragraphIndex);
      run.dataset.runIndex = String(runIndex);
      run.textContent = textNode.text || "";
      paragraphElement.appendChild(run);
    });
    body.appendChild(paragraphElement);
  });
  return body;
}

async function commitTextEditor() {
  if (!activeTextEditor) return;
  const docJson = textEditorDocJson();
  const handle = activeTextEditor.shape.handle;
  closeTextEditor();
  await editor.applyTextBodyDocJson(handle, docJson);
  render();
  message.textContent = "Applied";
}

async function replaceSelectedImage(file) {
  if (!selectedShape?.handle || !selectedShape.editableImageReplacement) {
    throw new Error("Select an image shape before replacing media.");
  }
  const result = await editor.apply({
    kind: "replaceImage",
    handle: selectedShape.handle,
    bytes: new Uint8Array(await file.arrayBuffer()),
  });
  render();
  message.textContent = imageReplacementMessage(result);
}

function imageReplacementMessage(result) {
  const shared = result.warnings?.find((warning) => warning.code === "shared-media-part");
  if (!shared) return "Image replaced";
  return (
    "Image replaced; shared media part affects " +
    shared.referenceCount +
    " pictures: " +
    shared.mediaPartPath
  );
}

function syncImageReplacementInput() {
  const replacement = selectedShape?.editableImageReplacement;
  imageReplacementInput.disabled = !replacement;
  imageReplacementInput.value = "";
  if (replacement) {
    imageReplacementInput.setAttribute("accept", replacement.accept);
  } else {
    imageReplacementInput.removeAttribute("accept");
  }
}

function textEditorDocJson() {
  const original = activeTextEditor.originalDocJson;
  return {
    type: "doc",
    content: (original.content || []).map((paragraph, paragraphIndex) => {
      const paragraphElement = activeTextEditor.element.querySelector(
        '.text-editor-run[data-paragraph-index="' + paragraphIndex + '"]'
      )?.parentElement;
      return {
        type: "paragraph",
        attrs: paragraph.attrs || {},
        content: (paragraph.content || []).flatMap((textNode, runIndex) => {
          const runElement = paragraphElement?.querySelector(
            '.text-editor-run[data-run-index="' + runIndex + '"]'
          );
          const text = runElement ? runElement.textContent || "" : textNode.text || "";
          return text.length === 0 ? [] : [{ type: "text", text, marks: textNode.marks || [] }];
        }),
      };
    }),
  };
}

function closeTextEditor() {
  if (!activeTextEditor) return;
  activeTextEditor.element.remove();
  activeTextEditor = null;
}

function positionActiveTextEditor() {
  if (!activeTextEditor?.shape?.bounds) return;
  const renderedSvg = container.querySelector("svg:not(#selection-overlay)");
  if (!renderedSvg) return;
  const svgRect = renderedSvg.getBoundingClientRect();
  const containerRect = container.getBoundingClientRect();
  const viewBox = parseViewBox(renderedSvg.getAttribute("viewBox") || "0 0 960 540");
  const bounds = activeTextEditor.shape.bounds;
  activeTextEditor.element.style.left =
    svgRect.left - containerRect.left + ((bounds.x - viewBox.x) / viewBox.width) * svgRect.width + "px";
  activeTextEditor.element.style.top =
    svgRect.top - containerRect.top + ((bounds.y - viewBox.y) / viewBox.height) * svgRect.height + "px";
  activeTextEditor.element.style.width = (bounds.width / viewBox.width) * svgRect.width + "px";
  activeTextEditor.element.style.height = (bounds.height / viewBox.height) * svgRect.height + "px";
}

function shapeKey(shape) {
  return [
    shape.handle?.partPath || "",
    shape.handle?.nodeId || "",
    shape.handle?.relationshipId || "",
    shape.handle?.orderingSlot == null ? "" : String(shape.handle.orderingSlot),
  ].join("\\u0000");
}

function cloneShape(shape) {
  return { ...shape, bounds: { ...shape.bounds } };
}

function setRectAttributes(rect, bounds) {
  rect.setAttribute("x", String(bounds.x));
  rect.setAttribute("y", String(bounds.y));
  rect.setAttribute("width", String(bounds.width));
  rect.setAttribute("height", String(bounds.height));
}

function handlePoint(bounds, handle) {
  const right = bounds.x + bounds.width;
  const bottom = bounds.y + bounds.height;
  if (handle === "nw") return { x: bounds.x, y: bounds.y };
  if (handle === "ne") return { x: right, y: bounds.y };
  if (handle === "sw") return { x: bounds.x, y: bottom };
  return { x: right, y: bottom };
}

function eventPoint(svg, event) {
  const matrix = svg?.getScreenCTM();
  if (!matrix) return null;
  const point = svg.createSVGPoint();
  point.x = event.clientX;
  point.y = event.clientY;
  return point.matrixTransform(matrix.inverse());
}

function movedBounds(bounds, dx, dy) {
  return { x: bounds.x + dx, y: bounds.y + dy, width: bounds.width, height: bounds.height };
}

function resizedBounds(bounds, handle, dx, dy) {
  const minSize = 8;
  const right = bounds.x + bounds.width;
  const bottom = bounds.y + bounds.height;
  const next = { ...bounds };
  if (handle === "nw" || handle === "sw") {
    next.x = Math.min(bounds.x + dx, right - minSize);
    next.width = right - next.x;
  }
  if (handle === "ne" || handle === "se") next.width = Math.max(minSize, bounds.width + dx);
  if (handle === "nw" || handle === "ne") {
    next.y = Math.min(bounds.y + dy, bottom - minSize);
    next.height = bottom - next.y;
  }
  if (handle === "sw" || handle === "se") next.height = Math.max(minSize, bounds.height + dy);
  return next;
}

function parseViewBox(value) {
  const parts = String(value).trim().split(/\\s+/).map(Number);
  return {
    x: Number.isFinite(parts[0]) ? parts[0] : 0,
    y: Number.isFinite(parts[1]) ? parts[1] : 0,
    width: Number.isFinite(parts[2]) && parts[2] > 0 ? parts[2] : 960,
    height: Number.isFinite(parts[3]) && parts[3] > 0 ? parts[3] : 540,
  };
}
`;

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

function pngBytes(base64: string): Uint8Array {
  return new Uint8Array(Buffer.from(base64, "base64"));
}

async function buildShapeFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<p:cSld><p:spTree>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/>` +
        `<a:p><a:r><a:rPr sz="2400"><a:latin typeface="Aptos"/></a:rPr><a:t>Original</a:t></a:r></a:p>` +
        `</p:txBody>` +
        `</p:sp>` +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );

  return zip.generateAsync({ type: "uint8array" });
}

async function buildImageFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `</Types>`,
    ),
  );
  zip.file(
    "_rels/.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/presentation.xml",
    xml(
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    xml(
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<p:cSld><p:spTree>` +
        `<p:pic><p:nvPicPr><p:cNvPr id="20" name="Shared Picture A"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>` +
        `<p:pic><p:nvPicPr><p:cNvPr id="21" name="Shared Picture B"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>` +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
        `</Relationships>`,
    ),
  );
  zip.file("ppt/media/image1.png", RED_PNG);

  return zip.generateAsync({ type: "uint8array" });
}

function firstShape(source: PptxSourceModel): SourceShape {
  const shape = source.slides[0]?.shapes.find((node): node is SourceShape => node.kind === "shape");
  if (shape === undefined) throw new Error("fixture shape not found");
  return shape;
}

function firstText(source: PptxSourceModel): string {
  const run = firstShape(source).textBody?.paragraphs[0]?.runs[0];
  if (run === undefined) throw new Error("fixture text run not found");
  return run.text;
}

function mediaBytes(source: PptxSourceModel, partPath: string): Uint8Array {
  const media = source.packageGraph.media.find((part) => part.partPath === partPath);
  if (media === undefined) throw new Error(`media not found: ${partPath}`);
  return media.bytes;
}

async function listen(server: Server): Promise<string> {
  await new Promise<void>((resolveListener) => {
    server.listen(0, "127.0.0.1", resolveListener);
  });
  const address = server.address();
  if (address === null || typeof address === "string") {
    throw new Error("server did not expose a TCP port");
  }
  return `http://127.0.0.1:${address.port.toString()}`;
}

async function closeServer(server: Server): Promise<void> {
  if (!server.listening) return;
  await new Promise<void>((resolveClose, rejectClose) => {
    server.close((error) => {
      if (error) rejectClose(error);
      else resolveClose();
    });
  });
}

declare global {
  interface Window {
    __pptxGlimpseEditorSmoke?: {
      hasEditableTextBody: boolean;
    };
  }
}
