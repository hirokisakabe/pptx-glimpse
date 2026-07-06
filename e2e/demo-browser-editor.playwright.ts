import { execFile } from "node:child_process";
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
const documentPackageRoot = resolve(repoRoot, "packages/document");
const editorCorePackageRoot = resolve(repoRoot, "packages/editor-core");
const rendererPackageRoot = resolve(repoRoot, "packages/renderer");
const corePackageRoot = resolve(repoRoot, "packages/core");
const execFileAsync = promisify(execFile);
const EMU_PER_PIXEL = 9525;
const BLUE_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAE0lEQVR4nGNkYPjPAANMcBZeDgAx0wEH1s7nlgAAAABJRU5ErkJggg==";
const BLUE_PNG = new Uint8Array(Buffer.from(BLUE_PNG_BASE64, "base64"));

let coreDistBuildPromise: Promise<void> | null = null;
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

test("runs the public demo browser editor flow entirely client-side", async ({ page }) => {
  test.setTimeout(120_000);
  if (demoServer === null) throw new Error("demo server was not started");
  const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-demo-editor-test-"));
  try {
    const savedPath = join(dir, "demo.edited.pptx");
    const replacementImagePath = join(dir, "replacement.png");
    await writeFile(replacementImagePath, BLUE_PNG);

    await page.goto(demoServer.url);

    await page.getByTestId("sample-basic-theme").click();
    await expect(page.getByTestId("viewer-status")).toContainText("slides rendered");
    await page.getByTestId("open-editor").click();
    await expect(page.getByTestId("editor-workspace")).toBeVisible();
    await expect(page.getByTestId("editor-status")).toContainText("ready");

    await page.getByRole("button", { name: "Duplicate" }).click();
    await expect(page.getByTestId("editor-thumbnail")).toHaveCount(3);
    await page.getByRole("button", { name: "Delete", exact: true }).click();
    await expect(page.getByTestId("editor-thumbnail")).toHaveCount(2);

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

    await page.getByTestId("text-run-input").fill("Added from demo e2e");
    await page.getByRole("button", { name: "Apply text" }).click();
    await expect(page.getByTestId("editor-slide-frame")).toContainText("Added from demo e2e");

    await page.getByRole("button", { name: "B", exact: true }).click();
    await expect(page.getByTestId("editor-status")).toContainText("Text style updated");

    await page.getByRole("button", { name: "Delete shape" }).click();
    await expect(page.getByTestId("editor-slide-frame")).not.toContainText("Added from demo e2e");
    await page.getByRole("button", { name: "Undo" }).click();
    await expect(page.getByTestId("editor-slide-frame")).toContainText("Added from demo e2e");

    await page.getByTestId("editor-thumbnail").nth(1).click();
    await selectFirstReplaceableImage(page);
    await page.getByTestId("image-replacement-input").setInputFiles(replacementImagePath);
    await expect(page.getByTestId("editor-status")).toContainText("Image replaced");

    const download = page.waitForEvent("download");
    await page.getByRole("button", { name: "Download PPTX" }).click();
    await (await download).saveAs(savedPath);

    const saved = readPptx(await readFile(savedPath));
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
  const child = execFile(
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
    await ensureCoreDist();
    if (!existsSync(resolve(demoRoot, "node_modules/.bin/next"))) {
      await execFileAsync("npm", ["ci"], { cwd: demoRoot, maxBuffer: 20 * 1024 * 1024 });
    }
    await execFileAsync("npm", ["run", "build"], { cwd: demoRoot, maxBuffer: 20 * 1024 * 1024 });
  })();
  await demoBuildPromise;
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

async function stopChild(child: ReturnType<typeof execFile>): Promise<void> {
  if (child.exitCode !== null) return;
  await new Promise<void>((resolveExit) => {
    child.once("exit", () => resolveExit());
    child.kill();
  });
}

async function waitForServerReady(child: ReturnType<typeof execFile>, port: number): Promise<void> {
  const output: string[] = [];
  child.stdout?.on("data", (chunk: Buffer) => output.push(chunk.toString()));
  child.stderr?.on("data", (chunk: Buffer) => output.push(chunk.toString()));

  const deadline = Date.now() + 30_000;
  while (Date.now() < deadline) {
    if (child.exitCode !== null) {
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
