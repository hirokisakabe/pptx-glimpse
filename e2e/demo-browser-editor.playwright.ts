import { type ChildProcessWithoutNullStreams, execFile, spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { createServer, type Server } from "node:http";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { expect, type Page, test } from "@playwright/test";

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

test("links the toolkit demo to all public package documentation", async ({ page }) => {
  if (demoServer === null) throw new Error("demo server was not started");

  await page.goto(demoServer.url);
  await expect(page.getByRole("link", { name: "GitHub" })).toHaveAttribute("target", "_blank");
  await expect(page.getByRole("link", { name: "GitHub" })).toHaveAttribute("rel", "noreferrer");
  await page.getByRole("link", { name: "Documentation" }).click();

  await expect(page).toHaveURL(`${demoServer.url}/docs`);
  await expect(page).toHaveTitle("Documentation | pptx-glimpse");
  await expect(page.getByRole("heading", { name: "pptx-glimpse", exact: true })).toBeVisible();
  await expect(
    page.getByRole("heading", { name: "@pptx-glimpse/document", exact: true }),
  ).toBeVisible();
  await expect(
    page.getByRole("heading", { name: "@pptx-glimpse/editor", exact: true }),
  ).toBeVisible();
  await expect(page.getByRole("link", { name: /Start with the toolkit/ })).toHaveAttribute(
    "href",
    "https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/core/README.md",
  );
  await expect(page.getByRole("link", { name: /Choose a document workflow/ })).toHaveAttribute(
    "href",
    "https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/document/README.md",
  );
  await expect(page.getByRole("link", { name: /Build a headless editor/ })).toHaveAttribute(
    "href",
    "https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/editor/README.md",
  );

  await page.setViewportSize({ width: 900, height: 720 });
  await expect(page.locator(".package-routes")).toHaveCSS(
    "grid-template-columns",
    /^[0-9.]+px [0-9.]+px$/,
  );

  const sitemapResponse = await page.request.get(`${demoServer.url}/sitemap.xml`);
  expect(sitemapResponse.ok()).toBe(true);
  expect(await sitemapResponse.text()).toContain("https://glimpse.pptx.app/docs");
});

test("opens the sample editor first and replaces it with an uploaded PPTX", async ({ page }) => {
  test.setTimeout(120_000);
  if (demoServer === null) throw new Error("demo server was not started");

  await page.goto(demoServer.url);
  await expect(page.getByTestId("editor-workspace")).toBeVisible();
  await expect(page.getByText("real-basic-theme.pptx")).toBeVisible();
  await expect(page.getByRole("button", { name: "Open PPTX" })).toBeVisible();
  await expect(page.getByRole("button", { name: "Download PPTX" })).toBeVisible();

  await page
    .getByTestId("pptx-input")
    .setInputFiles(resolve(repoRoot, "shared-fixtures/real-product-page.pptx"));
  await expect(page.getByTestId("editor-workspace")).toBeVisible();
  await expect(page.getByText("real-product-page.pptx")).toBeVisible();

  await page.getByRole("button", { name: "Add text box" }).click();
  page.once("dialog", async (dialog) => {
    expect(dialog.message()).toContain("Discard your unsaved changes");
    await dialog.dismiss();
  });
  await page.getByRole("button", { name: "Open sample" }).click();
  await expect(page.getByText("real-product-page.pptx")).toBeVisible();

  page.once("dialog", (dialog) => dialog.dismiss());
  await page.getByRole("link", { name: "Documentation" }).click();
  await expect(page).toHaveURL(demoServer.url);

  await page.getByTestId("pptx-input").setInputFiles({
    name: "invalid.pptx",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    buffer: Buffer.from("not a PPTX"),
  });
  await expect(page.getByRole("alert")).toBeVisible();
  await expect(page.getByText("real-product-page.pptx")).toBeVisible();

  page.once("dialog", (dialog) => dialog.accept());
  await page.getByRole("button", { name: "Open sample" }).click();
  await expect(page.getByTestId("editor-workspace")).toBeVisible();
  await expect(page.getByText("real-basic-theme.pptx")).toBeVisible();
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

    const slideFrame = page.getByTestId("editor-slide-frame");
    const firstEditableTextShape = page
      .locator('[data-testid="shape-hit-area"][data-editable-text="true"]')
      .first();
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
    await expect(page.getByTestId("editor-thumbnail")).toHaveCount(3);
    await page.getByRole("button", { name: "Delete", exact: true }).click();
    await expect(page.getByTestId("editor-thumbnail")).toHaveCount(2);

    await firstEditableTextShape.click();
    await expect(slideFrame).toBeFocused();
    await page.keyboard.press("Enter");
    await expect(directTextEditor).toBeVisible();
    await expect(directTextRun).toBeFocused();
    await directTextRun.fill("Direct edit done");
    await page.getByTestId("direct-text-editor-done").click();
    await expect(slideFrame).toContainText("Direct edit done");

    await page.getByRole("button", { name: "Undo" }).click();
    await expect(slideFrame).toContainText("たいとる");
    await expect(slideFrame).not.toContainText("Direct edit done");
    await page.getByRole("button", { name: "Redo" }).click();
    await expect(slideFrame).toContainText("Direct edit done");

    await firstEditableTextShape.dblclick();
    await directTextRun.fill("Direct edit saved");
    const focusoutDownload = page.waitForEvent("download");
    await page.getByRole("button", { name: "Download PPTX" }).click();
    await (await focusoutDownload).saveAs(focusoutSavedPath);
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
    await selectFirstReplaceableImage(page);
    await page.getByTestId("image-replacement-input").setInputFiles(replacementImagePath);
    await expect(page.getByTestId("editor-status")).toContainText("Image replaced");

    const download = page.waitForEvent("download");
    await page.getByRole("button", { name: "Download PPTX" }).click();
    await (await download).saveAs(savedPath);

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
    expect(mediaBytes(saved, "ppt/media/image1.png")).toEqual(BLUE_PNG);
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
  })();
  await demoBuildPromise;
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

function mediaBytes(source: PptxSourceModel, partPath: string): Uint8Array {
  const media = source.packageGraph.media.find((part) => part.partPath === partPath);
  if (media === undefined) throw new Error(`media not found: ${partPath}`);
  return media.bytes;
}
