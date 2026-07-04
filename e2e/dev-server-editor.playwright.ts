import { type ChildProcessWithoutNullStreams, spawn } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { createServer, type Server } from "node:http";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { expect, type Page, test } from "@playwright/test";
import JSZip from "jszip";

import {
  type PptxSourceModel,
  readPptx,
  type SourceShape,
} from "../packages/document/src/index.js";

const encoder = new TextEncoder();
const EMU_PER_PIXEL = 9525;

test("selects a shape, moves and resizes it, then saves xfrm edits", async ({ page }) => {
  const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-editor-ui-test-"));
  let serverProcess: ChildProcessWithoutNullStreams | undefined;
  try {
    const sourcePath = join(dir, "fixture.pptx");
    const savedPath = join(dir, "fixture.edited.pptx");
    await writeFile(sourcePath, await buildShapeFixture());

    const port = await findFreePort();
    serverProcess = await startDevServer(sourcePath, port);
    const baseUrl = `http://127.0.0.1:${String(port)}`;

    await page.goto(baseUrl);
    const hitArea = page.getByTestId("shape-hit-area").first();
    await expect(hitArea).toBeVisible();

    await clickSvgPoint(page, 240, 240);
    await expect(page.getByTestId("selection-box")).toHaveAttribute("x", "96");
    await expect(page.getByTestId("selection-handle-se")).toBeVisible();

    await dragSvgPoint(page, { x: 240, y: 240 }, { x: 264, y: 256 });
    await expect(page.getByTestId("selection-box")).toHaveAttribute("x", "120");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("y", "208");

    await dragSvgPoint(page, { x: 408, y: 304 }, { x: 456, y: 328 });
    await expect(page.getByTestId("selection-box")).toHaveAttribute("width", "336");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("height", "120");

    await page.getByRole("button", { name: "Save" }).click();
    await expect(page.locator("#editor-message")).toContainText(savedPath);

    const saved = readPptx(await readFile(savedPath));
    expect(firstShape(saved).transform).toMatchObject({
      offsetX: 120 * EMU_PER_PIXEL,
      offsetY: 208 * EMU_PER_PIXEL,
      width: 336 * EMU_PER_PIXEL,
      height: 120 * EMU_PER_PIXEL,
    });
  } finally {
    if (serverProcess !== undefined) {
      serverProcess.kill();
      await waitForExit(serverProcess);
    }
    await rm(dir, { recursive: true, force: true });
  }
});

async function clickSvgPoint(page: Page, x: number, y: number) {
  const point = await svgPointToClient(page, x, y);
  await page.mouse.click(point.x, point.y);
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

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
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
        `</p:sp>` +
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );

  return zip.generateAsync({ type: "uint8array" });
}

function firstShape(source: PptxSourceModel): SourceShape {
  const shape = source.slides[0]?.shapes.find((node): node is SourceShape => node.kind === "shape");
  if (shape === undefined) throw new Error("fixture shape not found");
  return shape;
}

async function findFreePort(): Promise<number> {
  const server = createServer();
  await new Promise<void>((resolve) => {
    server.listen(0, "127.0.0.1", resolve);
  });
  const address = server.address();
  if (address === null || typeof address === "string") {
    throw new Error("test server did not expose a TCP port");
  }
  const port = address.port;
  await closeServer(server);
  return port;
}

async function startDevServer(
  sourcePath: string,
  port: number,
): Promise<ChildProcessWithoutNullStreams> {
  const child = spawn(
    "pnpm",
    ["exec", "tsx", "scripts/dev-server.ts", sourcePath, "--port", String(port)],
    {
      cwd: new URL("..", import.meta.url),
      stdio: ["ignore", "pipe", "pipe"],
    },
  );

  try {
    await new Promise<void>((resolve, reject) => {
      let output = "";
      const timeout = setTimeout(() => {
        reject(new Error(`dev server did not start:\n${output}`));
      }, 30_000);

      child.stdout.on("data", (chunk: Buffer) => {
        output += chunk.toString("utf8");
        if (output.includes("Dev server running")) {
          clearTimeout(timeout);
          resolve();
        }
      });
      child.stderr.on("data", (chunk: Buffer) => {
        output += chunk.toString("utf8");
      });
      child.on("error", (error) => {
        clearTimeout(timeout);
        reject(error);
      });
      child.on("exit", (code) => {
        clearTimeout(timeout);
        reject(new Error(`dev server exited with code ${String(code)}:\n${output}`));
      });
    });
  } catch (error) {
    child.kill();
    await waitForExit(child);
    throw error;
  }

  return child;
}

async function closeServer(server: Server): Promise<void> {
  if (!server.listening) return;
  await new Promise<void>((resolve, reject) => {
    server.close((error) => {
      if (error) reject(error);
      else resolve();
    });
  });
}

async function waitForExit(child: ChildProcessWithoutNullStreams): Promise<void> {
  if (child.exitCode !== null || child.signalCode !== null) return;
  await new Promise<void>((resolve) => {
    child.once("exit", () => resolve());
  });
}
