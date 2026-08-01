import { type ChildProcessWithoutNullStreams, execFile, spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { createServer, type Server } from "node:http";
import { builtinModules } from "node:module";
import { tmpdir } from "node:os";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { expect, type Locator, type Page, test } from "@playwright/test";

import {
  type PptxSourceModel,
  readPptx,
  type SourceShape,
} from "../packages/document/src/index.js";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const demoRoot = resolve(repoRoot, "demo");
const execFileAsync = promisify(execFile);
const EMU_PER_PIXEL = 9525;
const BLUE_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAE0lEQVR4nGNkYPjPAANMcBZeDgAx0wEH1s7nlgAAAABJRU5ErkJggg==";
const BLUE_PNG = new Uint8Array(Buffer.from(BLUE_PNG_BASE64, "base64"));

let demoBuildPromise: Promise<void> | null = null;
let demoServer: DemoServer | null = null;

test.beforeAll(async () => {
  test.setTimeout(240_000);
  demoServer = await startDemoServer();
});

test.afterAll(async () => {
  await demoServer?.close();
  demoServer = null;
});

test("navigates the documentation guides and public package references", async ({ page }) => {
  if (demoServer === null) throw new Error("demo server was not started");

  await page.goto(demoServer.url);
  await expect(page.getByRole("link", { name: "GitHub" })).toHaveAttribute("target", "_blank");
  await expect(page.getByRole("link", { name: "GitHub" })).toHaveAttribute("rel", "noreferrer");
  await page.getByRole("link", { name: "Documentation" }).click();

  await expect(page).toHaveURL(`${demoServer.url}/docs`);
  await expect(page).toHaveTitle("Overview | pptx-glimpse");
  await expect(page.getByRole("heading", { name: "Overview", exact: true })).toBeVisible();
  const docsSidebar = page.locator(".nextra-sidebar");
  await expect(docsSidebar).toBeVisible();
  await expect(docsSidebar.getByRole("link", { name: "Overview", exact: true })).toBeVisible();

  await docsSidebar.getByRole("link", { name: "Choosing a package", exact: true }).click();
  await expect(page).toHaveURL(`${demoServer.url}/docs/packages`);
  await expect(page).toHaveTitle("Choosing a package | pptx-glimpse");
  await expect(page.getByRole("heading", { name: "pptx-glimpse", exact: true })).toBeVisible();
  await expect(
    page.getByRole("heading", { name: "@pptx-glimpse/document", exact: true }),
  ).toBeVisible();
  await expect(
    page.getByRole("heading", { name: "@pptx-glimpse/editor", exact: true }),
  ).toBeVisible();
  for (const href of [
    "https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/core/README.md",
    "https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/README.md",
    "https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/editor/README.md",
  ]) {
    await expect(page.locator(`a[href="${href}"]`)).toBeVisible();
  }

  await page.goto(`${demoServer.url}/docs/api`);
  await expect(page.getByRole("heading", { name: "API Reference", exact: true })).toBeVisible();
  await expect(docsSidebar.getByText("Conversion", { exact: true })).toBeVisible();

  await page.goto(`${demoServer.url}/docs/api/node/functions/convertPptxToSvg`);
  await expect(page.getByRole("heading", { name: "Function: convertPptxToSvg()" })).toBeVisible();
  await expect(page.locator("#convertoptions")).toBeVisible();

  await page.goto(`${demoServer.url}/docs/api/browser/functions/initResvgWasm`);
  await expect(page.getByRole("heading", { name: "Function: initResvgWasm()" })).toBeVisible();

  await page.setViewportSize({ width: 390, height: 720 });
  await expect(docsSidebar).toBeHidden();
  await expect(page.getByRole("button", { name: "Menu" })).toBeVisible();

  const sitemapResponse = await page.request.get(`${demoServer.url}/sitemap.xml`);
  expect(sitemapResponse.ok()).toBe(true);
  const sitemap = await sitemapResponse.text();
  for (const route of [
    "/docs",
    "/docs/getting-started",
    "/docs/why",
    "/docs/rendering",
    "/docs/editing",
    "/docs/fonts",
    "/docs/browser",
    "/docs/nodejs",
    "/docs/api",
    "/docs/feature-support",
    "/docs/packages",
  ]) {
    expect(sitemap).toContain(`https://glimpse.pptx.app${route}`);
  }
});

test("opens the sample editor first and replaces it with an uploaded PPTX", async ({ page }) => {
  test.setTimeout(120_000);
  if (demoServer === null) throw new Error("demo server was not started");

  const sampleResponse = await page.request.get(`${demoServer.url}/samples/editor-demo.pptx`);
  expect(sampleResponse.ok()).toBe(true);
  expect(await sampleResponse.body()).toEqual(
    await readFile(resolve(demoRoot, "assets/editor-demo.pptx")),
  );

  await page.goto(demoServer.url);
  await expect(page.getByTestId("editor-workspace")).toBeVisible();
  const fileNameInput = page.getByRole("textbox", { name: "Presentation file name" });
  await expect(fileNameInput).toHaveValue("editor-demo");
  await expect(page.getByRole("button", { name: "Open PPTX" })).toBeVisible();
  await expect(page.getByRole("button", { name: "Download PPTX" })).toBeVisible();

  await page
    .getByTestId("pptx-input")
    .setInputFiles(resolve(repoRoot, "shared-fixtures/real-product-page.pptx"));
  await expect(page.getByTestId("editor-workspace")).toBeVisible();
  await expect(fileNameInput).toHaveValue("real-product-page");

  await page.getByRole("button", { name: "Add text box" }).click();
  page.once("dialog", async (dialog) => {
    expect(dialog.message()).toContain("Discard your unsaved changes");
    await dialog.dismiss();
  });
  await page.getByRole("button", { name: "Open sample" }).click();
  await expect(fileNameInput).toHaveValue("real-product-page");

  page.once("dialog", (dialog) => dialog.dismiss());
  await page.getByRole("link", { name: "Documentation" }).click();
  await expect(page).toHaveURL(demoServer.url);

  await page.getByTestId("pptx-input").setInputFiles({
    name: "invalid.pptx",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    buffer: Buffer.from("not a PPTX"),
  });
  await expect(page.locator(".replacement-error")).toBeVisible();
  await expect(fileNameInput).toHaveValue("real-product-page");

  page.once("dialog", (dialog) => dialog.accept());
  await page.getByRole("button", { name: "Open sample" }).click();
  await expect(page.getByTestId("editor-workspace")).toBeVisible();
  await expect(fileNameInput).toHaveValue("editor-demo");
  await expect(
    page.getByRole("heading", { name: "pptx-glimpse browser editor demo" }),
  ).toBeAttached();
  await page.getByRole("link", { name: "Documentation" }).click();
  await expect(page).toHaveURL(`${demoServer.url}/docs`);
});

test("runs the public demo browser editor flow entirely client-side", async ({ page }) => {
  test.setTimeout(120_000);
  if (demoServer === null) throw new Error("demo server was not started");
  const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-demo-editor-test-"));
  try {
    const savedPath = join(dir, "demo.edited.pptx");
    const focusoutSavedPath = join(dir, "focusout.edited.pptx");
    const replacementImagePath = join(dir, "replacement.png");
    await writeFile(replacementImagePath, BLUE_PNG);

    await page.goto(demoServer.url);
    await expect(page.getByTestId("editor-workspace")).toBeVisible();
    const fileNameInput = page.getByRole("textbox", { name: "Presentation file name" });
    await expect(fileNameInput).toHaveValue("editor-demo");
    await fileNameInput.fill("browser-workshop");

    const slideFrame = page.getByTestId("editor-slide-frame");
    const firstEditableTextShape = page
      .locator('[data-testid="shape-hit-area"][data-editable-text="true"]')
      .nth(1);
    await firstEditableTextShape.dblclick();
    const directTextEditor = page.getByTestId("direct-text-editor");
    await expect(directTextEditor).toBeVisible();
    const focusedHitAreaBounds = await firstEditableTextShape.boundingBox();
    if (focusedHitAreaBounds === null) {
      throw new Error("focused editable text shape bounds were not available");
    }
    const editorBounds = await directTextEditor.boundingBox();
    if (editorBounds === null) throw new Error("direct text editor bounds were not available");
    expect(editorBounds.x).toBeCloseTo(focusedHitAreaBounds.x, 0);
    expect(editorBounds.y).toBeCloseTo(focusedHitAreaBounds.y, 0);
    expect(editorBounds.width).toBeCloseTo(focusedHitAreaBounds.width, 0);
    expect(editorBounds.height).toBeCloseTo(focusedHitAreaBounds.height, 0);

    const directTextRun = page.getByTestId("direct-text-editor-run").first();
    await expect(directTextRun).toBeFocused();
    await directTextRun.fill("Canceled direct edit");
    await directTextRun.press("Escape");
    await expect(directTextEditor).toHaveCount(0);
    await expect(slideFrame).not.toContainText("Canceled direct edit");
    await expect(page.getByRole("button", { name: "Undo" })).toBeDisabled();

    await page.getByRole("button", { name: "Duplicate" }).click();
    await expect(page.getByTestId("editor-status")).toContainText("Slide duplicated");
    await expect(page.getByTestId("editor-thumbnail")).toHaveCount(4);
    await page.getByRole("button", { name: "Delete", exact: true }).click();
    await expect(page.getByTestId("editor-thumbnail")).toHaveCount(3);

    const thumbnailsBeforeSort = page.getByTestId("editor-thumbnail");
    await expect(thumbnailsBeforeSort.nth(1)).toContainText("A SMALL, EDITABLE SYSTEM");
    await expect(thumbnailsBeforeSort.nth(1)).toBeEnabled();
    await expect(thumbnailsBeforeSort.first()).toHaveCSS("touch-action", "pan-y");
    await clickElementWithMovement(page, thumbnailsBeforeSort.first(), 2);
    await expect(thumbnailsBeforeSort.first()).toHaveClass(/active/);
    await clickElementWithMovement(page, thumbnailsBeforeSort.nth(1), 6);
    await expect(thumbnailsBeforeSort.nth(1)).toHaveClass(/active/);
    await dragElementAfter(page, thumbnailsBeforeSort.nth(1), thumbnailsBeforeSort.nth(2));
    await expect(page.getByTestId("editor-status")).toContainText("moved to position 3 of 3");
    await expect(page.getByTestId("editor-thumbnail").nth(2)).toContainText(
      "A SMALL, EDITABLE SYSTEM",
    );
    await expect(page.getByTestId("editor-thumbnail").nth(2)).toHaveClass(/active/);
    await page.getByRole("button", { name: "Undo" }).click();
    await expect(page.getByTestId("editor-thumbnail").nth(1)).toContainText(
      "A SMALL, EDITABLE SYSTEM",
    );
    await expect(page.getByTestId("editor-thumbnail").nth(1)).toHaveClass(/active/);
    await page.getByRole("button", { name: "Redo" }).click();
    await expect(page.getByTestId("editor-thumbnail").nth(2)).toContainText(
      "A SMALL, EDITABLE SYSTEM",
    );
    await expect(page.getByTestId("editor-thumbnail").nth(2)).toHaveClass(/active/);
    await page.getByTestId("editor-thumbnail").nth(2).press("Alt+ArrowUp");
    await expect(page.getByTestId("editor-status")).toContainText("moved to position 2 of 3");
    await expect(page.getByTestId("editor-thumbnail").nth(1)).toContainText(
      "A SMALL, EDITABLE SYSTEM",
    );
    await page.getByTestId("editor-thumbnail").nth(1).press("Alt+ArrowDown");
    await expect(page.getByTestId("editor-status")).toContainText("moved to position 3 of 3");
    await page.getByTestId("editor-thumbnail").first().click();
    await expect(page.getByTestId("editor-thumbnail").first()).toHaveClass(/active/);

    await firstEditableTextShape.click();
    await expect(slideFrame).toBeFocused();
    await page.keyboard.press("Enter");
    await expect(directTextEditor).toBeVisible();
    await expect(directTextRun).toBeFocused();
    await directTextRun.fill("Direct edit done");
    await page.getByTestId("direct-text-editor-done").click();
    await expect(slideFrame).toContainText("Direct edit done");

    await page.getByRole("button", { name: "Undo" }).click();
    await expect(slideFrame).toContainText("Make the");
    await expect(slideFrame).toContainText("presentation");
    await expect(slideFrame).not.toContainText("Direct edit done");
    await page.getByRole("button", { name: "Redo" }).click();
    await expect(slideFrame).toContainText("Direct edit done");

    await firstEditableTextShape.dblclick();
    await directTextRun.fill("Direct edit saved");
    const focusoutDownload = page.waitForEvent("download");
    await page.getByRole("button", { name: "Download PPTX" }).click();
    const focusoutDownloadedFile = await focusoutDownload;
    expect(focusoutDownloadedFile.suggestedFilename()).toBe("browser-workshop.pptx");
    await focusoutDownloadedFile.saveAs(focusoutSavedPath);
    await expect(directTextEditor).toHaveCount(0);
    await expect(slideFrame).toContainText("Direct edit saved");
    expect(
      shapeByText(readPptx(await readFile(focusoutSavedPath)), "Direct edit saved"),
    ).toBeDefined();
    await page.getByRole("button", { name: "Undo" }).click();
    await expect(slideFrame).toContainText("Direct edit done");
    await page.getByRole("button", { name: "Redo" }).click();
    await expect(slideFrame).toContainText("Direct edit saved");

    await page.getByRole("button", { name: "Add text box" }).click();
    await expect(page.getByTestId("selection-box")).toHaveAttribute("x", "96");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("y", "96");

    await dragSvgPoint(page, { x: 240, y: 132 }, { x: 264, y: 148 });
    await expect(page.getByTestId("selection-box")).toHaveAttribute("x", "120");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("y", "112");

    await dragSvgPoint(page, { x: 408, y: 184 }, { x: 456, y: 208 });
    await expect(page.getByTestId("selection-box")).toHaveAttribute("width", "336");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("height", "96");

    await page.getByRole("button", { name: "Undo" }).click();
    await expect(page.getByTestId("selection-box")).toHaveAttribute("width", "288");
    await page.getByRole("button", { name: "Redo" }).click();
    await expect(page.getByTestId("selection-box")).toHaveAttribute("width", "336");

    await slideFrame.focus();
    await page.keyboard.press("Enter");
    await expect(directTextEditor).toBeVisible();
    await directTextRun.fill("Added from demo e2e");
    await page.getByTestId("direct-text-editor-done").click();
    await expect(page.getByTestId("editor-slide-frame")).toContainText("Added from demo e2e");

    await page.getByRole("button", { name: "B", exact: true }).click();
    await expect(page.getByTestId("editor-status")).toContainText("Text style updated");

    await page.getByRole("button", { name: "Delete shape" }).click();
    await expect(page.getByTestId("editor-slide-frame")).not.toContainText("Added from demo e2e");
    await page.getByRole("button", { name: "Undo" }).click();
    await expect(page.getByTestId("editor-slide-frame")).toContainText("Added from demo e2e");

    const secondThumbnail = page.getByTestId("editor-thumbnail").nth(1);
    await expect(secondThumbnail).toBeEnabled();
    await secondThumbnail.click();
    await expect(secondThumbnail).toHaveClass(/active/);
    const firstThumbnail = page.getByTestId("editor-thumbnail").first();
    await firstThumbnail.click();
    await expect(firstThumbnail).toHaveClass(/active/);
    await selectFirstReplaceableImage(page);
    await page.getByTestId("image-replacement-input").setInputFiles(replacementImagePath);
    await expect(page.getByTestId("editor-status")).toContainText("Image replaced");

    const download = page.waitForEvent("download");
    await page.getByRole("button", { name: "Download PPTX" }).click();
    const downloadedFile = await download;
    expect(downloadedFile.suggestedFilename()).toBe("browser-workshop.pptx");
    await downloadedFile.saveAs(savedPath);

    const saved = readPptx(await readFile(savedPath));
    expect(shapeByText(saved, "Direct edit saved")).toBeDefined();
    const addedShape = shapeByText(saved, "Added from demo e2e");
    expect(addedShape.transform).toMatchObject({
      offsetX: 120 * EMU_PER_PIXEL,
      offsetY: 112 * EMU_PER_PIXEL,
      width: 336 * EMU_PER_PIXEL,
      height: 96 * EMU_PER_PIXEL,
    });
    expect(addedShape.textBody?.paragraphs[0]?.runs[0]?.properties?.bold).toBe(true);
    expect(mediaBytes(saved, "ppt/media/image.png")).toEqual(BLUE_PNG);
    expect(slideContainsText(saved, 2, "A SMALL, EDITABLE SYSTEM")).toBe(true);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

async function dragSvgPoint(
  page: Page,
  from: { readonly x: number; readonly y: number },
  to: { readonly x: number; readonly y: number },
) {
  const start = await svgPointToClient(page, from.x, from.y);
  const end = await svgPointToClient(page, to.x, to.y);
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(end.x, end.y, { steps: 6 });
  await page.mouse.up();
}

async function dragElementAfter(page: Page, source: Locator, target: Locator): Promise<void> {
  const sourceBounds = await source.boundingBox();
  const targetBounds = await target.boundingBox();
  if (sourceBounds === null || targetBounds === null) {
    throw new Error("drag source or target bounds were not available");
  }
  await page.mouse.move(
    sourceBounds.x + sourceBounds.width / 2,
    sourceBounds.y + sourceBounds.height / 2,
  );
  await page.mouse.down();
  await page.mouse.move(
    targetBounds.x + targetBounds.width / 2,
    targetBounds.y + targetBounds.height * 0.75,
    { steps: 12 },
  );
  await page.mouse.up();
}

async function clickElementWithMovement(
  page: Page,
  element: Locator,
  movementY: number,
): Promise<void> {
  const bounds = await element.boundingBox();
  if (bounds === null) throw new Error("click target bounds were not available");
  const x = bounds.x + bounds.width / 2;
  const y = bounds.y + bounds.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x, y + movementY, { steps: 2 });
  await page.mouse.up();
}

async function selectFirstReplaceableImage(page: Page): Promise<void> {
  const hitAreas = page.locator(
    '[data-testid="shape-hit-area"][data-editable-image-replacement="true"]',
  );
  const imageButton = page.getByTestId("replace-image-button");
  await expect(hitAreas.first()).toBeVisible();
  const count = await hitAreas.count();
  if (count === 0) throw new Error("replaceable image shape was not found");
  const bounds = await hitAreas.first().evaluate((element) => {
    if (!(element instanceof SVGRectElement)) throw new Error("shape hit area is not a rect");
    return {
      x: element.x.baseVal.value,
      y: element.y.baseVal.value,
      width: element.width.baseVal.value,
      height: element.height.baseVal.value,
    };
  });
  const point = await svgPointToClient(
    page,
    bounds.x + bounds.width / 2,
    bounds.y + bounds.height / 2,
  );
  await page.mouse.click(point.x, point.y);
  await expect(imageButton).toBeEnabled();
}

async function svgPointToClient(
  page: Page,
  x: number,
  y: number,
): Promise<{ readonly x: number; readonly y: number }> {
  return page.evaluate(
    ({ svgX, svgY }) => {
      const overlay = document.querySelector('[data-testid="selection-overlay"]');
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

interface DemoServer {
  readonly url: string;
  close(): Promise<void>;
}

async function startDemoServer(): Promise<DemoServer> {
  await ensureDemoBuild();
  const port = await findFreePort();
  const child = spawn(
    "npm",
    ["run", "start", "--", "--hostname", "127.0.0.1", "--port", String(port)],
    { cwd: demoRoot },
  );
  try {
    await waitForServerReady(child, port);
  } catch (error) {
    await stopChild(child);
    throw error;
  }
  return {
    url: `http://127.0.0.1:${String(port)}`,
    close: async () => {
      await stopChild(child);
    },
  };
}

async function ensureDemoBuild(): Promise<void> {
  demoBuildPromise ??= (async () => {
    if (!existsSync(resolve(demoRoot, "node_modules/.bin/next"))) {
      throw new Error("demo dependencies are not installed; run `cd demo && npm ci` first.");
    }
    await execFileAsync("npm", ["run", "build"], { cwd: demoRoot, maxBuffer: 20 * 1024 * 1024 });
    await assertDemoBrowserBundle();
  })();
  await demoBuildPromise;
}

async function assertDemoBrowserBundle(): Promise<void> {
  const chunksDirectory = resolve(demoRoot, ".next/static/chunks");
  const chunkPaths = await collectJavaScriptFiles(chunksDirectory);
  expect(chunkPaths.length, "the demo build should emit browser chunks").toBeGreaterThan(0);
  const builtinSpecifierPattern = builtinModules
    .map((specifier) => specifier.replace(/^node:/, ""))
    .map((specifier) => specifier.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
    .join("|");
  const nodeOnlyPattern = new RegExp(
    `(?:node:(?:${builtinSpecifierPattern})|\\[externals\\]/(?:node:)?(?:${builtinSpecifierPattern})|(?:from|import|import\\s*\\(|require\\s*\\()\\s*["'](?:${builtinSpecifierPattern})["']|node-font-loader|packages/core/dist/index)`,
  );

  for (const chunkPath of chunkPaths) {
    const source = await readFile(chunkPath, "utf8");
    const chunkName = relative(chunksDirectory, chunkPath);
    expect(source, `${chunkName} includes a Node-only dependency`).not.toMatch(nodeOnlyPattern);
  }
}

async function collectJavaScriptFiles(directory: string): Promise<string[]> {
  const entries = await readdir(directory, { withFileTypes: true });
  const nestedFiles = await Promise.all(
    entries.map(async (entry) => {
      const entryPath = resolve(directory, entry.name);
      if (entry.isDirectory()) return collectJavaScriptFiles(entryPath);
      return entry.isFile() && entry.name.endsWith(".js") ? [entryPath] : [];
    }),
  );
  return nestedFiles.flat();
}

async function findFreePort(): Promise<number> {
  const server = createServer();
  await new Promise<void>((resolveListen) => {
    server.listen(0, "127.0.0.1", resolveListen);
  });
  const address = server.address();
  if (address === null || typeof address === "string") {
    throw new Error("server did not expose a TCP port");
  }
  const port = address.port;
  await closeServer(server);
  return port;
}

async function closeServer(server: Server): Promise<void> {
  await new Promise<void>((resolveClose, rejectClose) => {
    server.close((error) => {
      if (error) rejectClose(error);
      else resolveClose();
    });
  });
}

async function stopChild(child: ChildProcessWithoutNullStreams): Promise<void> {
  if (hasExited(child)) return;
  await new Promise<void>((resolveExit) => {
    child.once("exit", () => resolveExit());
    child.kill();
  });
}

async function waitForServerReady(
  child: ChildProcessWithoutNullStreams,
  port: number,
): Promise<void> {
  const output: string[] = [];
  child.stdout?.on("data", (chunk: Buffer) => output.push(chunk.toString()));
  child.stderr?.on("data", (chunk: Buffer) => output.push(chunk.toString()));

  const deadline = Date.now() + 30_000;
  while (Date.now() < deadline) {
    if (hasExited(child)) {
      throw new Error(`demo server exited early:\n${output.join("")}`);
    }
    try {
      const response = await fetch(`http://127.0.0.1:${String(port)}`);
      if (response.ok) return;
    } catch {
      // Server is still starting.
    }
    await new Promise((resolveDelay) => setTimeout(resolveDelay, 250));
  }
  throw new Error(`demo server did not become ready:\n${output.join("")}`);
}

function hasExited(child: ChildProcessWithoutNullStreams): boolean {
  return child.exitCode !== null || child.signalCode !== null;
}

function shapeByText(source: PptxSourceModel, text: string): SourceShape {
  const shape = source.slides
    .flatMap((slide) => slide.shapes)
    .find((node): node is SourceShape => {
      if (node.kind !== "shape") return false;
      return node.textBody?.paragraphs.some((paragraph) =>
        paragraph.runs.some((run) => run.text === text),
      );
    });
  if (shape === undefined) throw new Error(`shape not found: ${text}`);
  return shape;
}

function slideContainsText(source: PptxSourceModel, index: number, text: string): boolean {
  return (
    source.slides[index]?.shapes.some(
      (node) =>
        node.kind === "shape" &&
        node.textBody?.paragraphs.some((paragraph) =>
          paragraph.runs.some((run) => run.text.includes(text)),
        ),
    ) ?? false
  );
}

function mediaBytes(source: PptxSourceModel, partPath: string): Uint8Array {
  const media = source.packageGraph.media.find((part) => part.partPath === partPath);
  if (media === undefined) throw new Error(`media not found: ${partPath}`);
  return media.bytes;
}
