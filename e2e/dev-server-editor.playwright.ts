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
  type SourceConnector,
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

test("edits a text shape with the overlay editor, re-renders, and saves text", async ({ page }) => {
  const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-editor-text-test-"));
  let serverProcess: ChildProcessWithoutNullStreams | undefined;
  try {
    const sourcePath = join(dir, "fixture.pptx");
    const savedPath = join(dir, "fixture.edited.pptx");
    await writeFile(sourcePath, await buildShapeFixture());

    const port = await findFreePort();
    serverProcess = await startDevServer(sourcePath, port);
    const baseUrl = `http://127.0.0.1:${String(port)}`;

    await page.goto(baseUrl);
    await expect(page.getByTestId("shape-hit-area").first()).toBeVisible();

    await dblclickSvgPoint(page, 240, 240);
    const overlay = page.getByTestId("text-editor-overlay");
    await expect(overlay).toBeVisible();
    await expect(overlay).toContainText("Original");
    await expect(page.getByRole("button", { name: "Delete shape" })).toBeDisabled();
    await expectOverlayNearSvgBounds(page, { x: 96, y: 192, width: 288, height: 96 });

    const commandResponse = page.waitForResponse(
      (response) =>
        response.url().endsWith("/api/editor/text-body") &&
        response.request().method() === "POST" &&
        response.ok(),
    );
    await page.getByTestId("text-editor-run").first().fill("Overlay edited");
    await page.getByTestId("text-editor-done").click();
    await commandResponse;

    await expect(page.locator("#slide-container")).toContainText("Overlay edited");
    await page.getByRole("button", { name: "Save" }).click();
    await expect(page.locator("#editor-message")).toContainText(savedPath);

    const saved = readPptx(await readFile(savedPath));
    expect(firstText(saved)).toBe("Overlay edited");
  } finally {
    if (serverProcess !== undefined) {
      serverProcess.kill();
      await waitForExit(serverProcess);
    }
    await rm(dir, { recursive: true, force: true });
  }
});

test("adds, edits, moves, resizes, saves, and deletes a text box from the UI", async ({ page }) => {
  const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-editor-add-delete-test-"));
  let serverProcess: ChildProcessWithoutNullStreams | undefined;
  try {
    const sourcePath = join(dir, "fixture.pptx");
    const savedPath = join(dir, "fixture.edited.pptx");
    await writeFile(sourcePath, await buildShapeFixture());

    const port = await findFreePort();
    serverProcess = await startDevServer(sourcePath, port);
    const baseUrl = `http://127.0.0.1:${String(port)}`;

    await page.goto(baseUrl);
    await expect(page.getByTestId("shape-hit-area").first()).toBeVisible();

    const addResponse = page.waitForResponse(
      (response) => response.url().endsWith("/api/editor/add-text-box") && response.ok(),
    );
    await page.getByRole("button", { name: "Add text box" }).click();
    await addResponse;
    await expect(page.getByTestId("selection-box")).toHaveAttribute("x", "96");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("y", "96");

    await dblclickSvgPoint(page, 216, 132);
    const overlay = page.getByTestId("text-editor-overlay");
    await expect(overlay).toBeVisible();
    await expect(overlay).toContainText("New text box");

    const textResponse = page.waitForResponse(
      (response) =>
        response.url().endsWith("/api/editor/text-body") &&
        response.request().method() === "POST" &&
        response.ok(),
    );
    await page.getByTestId("text-editor-run").first().fill("Added edited");
    await page.getByTestId("text-editor-done").click();
    await textResponse;
    await expect(page.locator("#slide-container")).toContainText("Added edited");

    await dragSvgPoint(page, { x: 240, y: 132 }, { x: 264, y: 148 });
    await expect(page.getByTestId("selection-box")).toHaveAttribute("x", "120");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("y", "112");

    await dragSvgPoint(page, { x: 408, y: 184 }, { x: 456, y: 208 });
    await expect(page.getByTestId("selection-box")).toHaveAttribute("width", "336");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("height", "96");

    await page.getByRole("button", { name: "Save" }).click();
    await expect(page.locator("#editor-message")).toContainText(savedPath);

    let saved = readPptx(await readFile(savedPath));
    expect(shapeByText(saved, "Added edited").transform).toMatchObject({
      offsetX: 120 * EMU_PER_PIXEL,
      offsetY: 112 * EMU_PER_PIXEL,
      width: 336 * EMU_PER_PIXEL,
      height: 96 * EMU_PER_PIXEL,
    });

    const deleteResponse = page.waitForResponse(
      (response) =>
        response.url().endsWith("/api/editor/command") &&
        response.request().method() === "POST" &&
        response.ok(),
    );
    await page.getByRole("button", { name: "Delete shape" }).click();
    await deleteResponse;
    await expect(page.getByTestId("selection-box")).toHaveCount(0);

    await page.getByRole("button", { name: "Save" }).click();
    await expect(page.locator("#editor-message")).toContainText(savedPath);
    saved = readPptx(await readFile(savedPath));
    expect(findShapeByText(saved, "Added edited")).toBeUndefined();
  } finally {
    if (serverProcess !== undefined) {
      serverProcess.kill();
      await waitForExit(serverProcess);
    }
    await rm(dir, { recursive: true, force: true });
  }
});

test("adds, moves, resizes, saves, deletes, undoes, and redoes a connector from the UI", async ({
  page,
}) => {
  const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-editor-connector-test-"));
  let serverProcess: ChildProcessWithoutNullStreams | undefined;
  try {
    const sourcePath = join(dir, "fixture.pptx");
    const savedPath = join(dir, "fixture.edited.pptx");
    await writeFile(sourcePath, await buildShapeFixture());

    const port = await findFreePort();
    serverProcess = await startDevServer(sourcePath, port);
    const baseUrl = `http://127.0.0.1:${String(port)}`;

    await page.goto(baseUrl);
    await expect(page.getByTestId("shape-hit-area").first()).toBeVisible();

    const addResponse = page.waitForResponse(
      (response) => response.url().endsWith("/api/editor/add-connector") && response.ok(),
    );
    await page.getByRole("button", { name: "Add connector" }).click();
    await addResponse;
    await expect(page.getByTestId("selection-box")).toHaveAttribute("x", "144");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("y", "144");

    await dragSvgPoint(page, { x: 240, y: 192 }, { x: 264, y: 208 });
    await expect(page.getByTestId("selection-box")).toHaveAttribute("x", "168");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("y", "160");

    await dragSvgPoint(page, { x: 456, y: 256 }, { x: 504, y: 280 });
    await expect(page.getByTestId("selection-box")).toHaveAttribute("width", "336");
    await expect(page.getByTestId("selection-box")).toHaveAttribute("height", "120");

    const undoResponse = page.waitForResponse(
      (response) => response.url().endsWith("/api/editor/undo") && response.ok(),
    );
    await page.getByRole("button", { name: "Undo" }).click();
    await undoResponse;
    await expect(page.getByTestId("selection-box")).toHaveAttribute("width", "288");

    const redoResponse = page.waitForResponse(
      (response) => response.url().endsWith("/api/editor/redo") && response.ok(),
    );
    await page.getByRole("button", { name: "Redo" }).click();
    await redoResponse;
    await expect(page.getByTestId("selection-box")).toHaveAttribute("width", "336");

    await page.getByRole("button", { name: "Save" }).click();
    await expect(page.locator("#editor-message")).toContainText(savedPath);

    const saved = readPptx(await readFile(savedPath));
    expect(connectorByName(saved, "Connector 11").transform).toMatchObject({
      offsetX: 168 * EMU_PER_PIXEL,
      offsetY: 160 * EMU_PER_PIXEL,
      width: 336 * EMU_PER_PIXEL,
      height: 120 * EMU_PER_PIXEL,
    });

    const deleteResponse = page.waitForResponse(
      (response) =>
        response.url().endsWith("/api/editor/command") &&
        response.request().method() === "POST" &&
        response.ok(),
    );
    await page.getByRole("button", { name: "Delete shape" }).click();
    await deleteResponse;
    await expect(page.getByTestId("selection-box")).toHaveCount(0);

    await page.getByRole("button", { name: "Undo" }).click();
    await expect(page.getByTestId("shape-hit-area")).toHaveCount(2);
    await page.getByRole("button", { name: "Redo" }).click();
    await expect(page.getByTestId("shape-hit-area")).toHaveCount(1);
  } finally {
    if (serverProcess !== undefined) {
      serverProcess.kill();
      await waitForExit(serverProcess);
    }
    await rm(dir, { recursive: true, force: true });
  }
});

test("applies text run decoration from the overlay toolbar with undo and redo", async ({
  page,
}) => {
  const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-editor-decoration-test-"));
  let serverProcess: ChildProcessWithoutNullStreams | undefined;
  try {
    const sourcePath = join(dir, "fixture.pptx");
    const savedPath = join(dir, "fixture.edited.pptx");
    await writeFile(sourcePath, await buildShapeFixture());

    const port = await findFreePort();
    serverProcess = await startDevServer(sourcePath, port);
    const baseUrl = `http://127.0.0.1:${String(port)}`;

    await page.goto(baseUrl);
    await expect(page.getByTestId("shape-hit-area").first()).toBeVisible();

    await dblclickSvgPoint(page, 240, 240);
    await expect(page.getByTestId("text-run-format-toolbar")).toBeVisible();
    const commandResponse = page.waitForResponse(
      (response) =>
        response.url().endsWith("/api/editor/command") &&
        response.request().method() === "POST" &&
        response.ok(),
    );
    await page.getByTestId("text-run-format-bold").click();
    await commandResponse;
    await expect(page.getByTestId("text-run-format-toolbar")).toBeVisible();

    const italicResponse = page.waitForResponse(
      (response) =>
        response.url().endsWith("/api/editor/command") &&
        response.request().method() === "POST" &&
        response.ok(),
    );
    await page.getByTestId("text-run-format-italic").click();
    await italicResponse;
    await expect(page.getByTestId("text-run-format-toolbar")).toBeVisible();
    const textBodyResponse = page.waitForResponse(
      (response) =>
        response.url().endsWith("/api/editor/text-body") &&
        response.request().method() === "POST" &&
        response.ok(),
    );
    await page.getByTestId("text-editor-done").click();
    await textBodyResponse;

    const undoResponse = page.waitForResponse(
      (response) => response.url().endsWith("/api/editor/undo") && response.ok(),
    );
    await page.getByRole("button", { name: "Undo" }).click();
    await undoResponse;

    await page.getByRole("button", { name: "Save" }).click();
    await expect(page.locator("#editor-message")).toContainText(savedPath);
    let saved = readPptx(await readFile(savedPath));
    expect(firstRunProperties(saved)).toMatchObject({ bold: true });
    expect(firstRunProperties(saved)?.italic).toBeUndefined();

    const redoResponse = page.waitForResponse(
      (response) => response.url().endsWith("/api/editor/redo") && response.ok(),
    );
    await page.getByRole("button", { name: "Redo" }).click();
    await redoResponse;

    await page.getByRole("button", { name: "Save" }).click();
    await expect(page.locator("#editor-message")).toContainText(savedPath);

    saved = readPptx(await readFile(savedPath));
    expect(firstRunProperties(saved)).toMatchObject({ bold: true, italic: true });
  } finally {
    if (serverProcess !== undefined) {
      serverProcess.kill();
      await waitForExit(serverProcess);
    }
    await rm(dir, { recursive: true, force: true });
  }
});

test("duplicates and deletes slides from thumbnails with undo and redo", async ({ page }) => {
  const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-editor-slide-test-"));
  let serverProcess: ChildProcessWithoutNullStreams | undefined;
  try {
    const sourcePath = join(dir, "fixture.pptx");
    const savedPath = join(dir, "fixture.edited.pptx");
    await writeFile(sourcePath, await buildTwoSlideFixture());

    const port = await findFreePort();
    serverProcess = await startDevServer(sourcePath, port);
    const baseUrl = `http://127.0.0.1:${String(port)}`;

    await page.goto(baseUrl);
    await expect(page.locator("#info")).toContainText("Slide 1 / 2");
    await expect(page.locator(".thumbnail")).toHaveCount(2);
    await expect(page.locator("#slide-container")).toContainText("First");

    await dblclickSvgPoint(page, 150, 130);
    await expect(page.getByTestId("text-editor-overlay")).toBeVisible();
    await page.getByTestId("duplicate-slide-0").click();
    await expect(page.locator("#editor-message")).toContainText(
      "Finish text editing before slide operations",
    );
    await expect(page.locator(".thumbnail")).toHaveCount(2);
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("text-editor-overlay")).toBeHidden();

    let commandResponse = page.waitForResponse(
      (response) =>
        response.url().endsWith("/api/editor/command") &&
        response.request().method() === "POST" &&
        response.ok(),
    );
    await page.getByTestId("duplicate-slide-0").click();
    await commandResponse;
    await expect(page.locator(".thumbnail")).toHaveCount(3);
    await expect(page.locator("#info")).toContainText("Slide 2 / 3");
    await expect(page.locator("#slide-container")).toContainText("First");

    commandResponse = page.waitForResponse(
      (response) =>
        response.url().endsWith("/api/editor/command") &&
        response.request().method() === "POST" &&
        response.ok(),
    );
    await page.getByTestId("delete-slide-1").click();
    await commandResponse;
    await expect(page.locator(".thumbnail")).toHaveCount(2);
    await expect(page.locator("#info")).toContainText("Slide 2 / 2");
    await expect(page.locator("#slide-container")).toContainText("Second");

    const undoResponse = page.waitForResponse(
      (response) => response.url().endsWith("/api/editor/undo") && response.ok(),
    );
    await page.getByRole("button", { name: "Undo" }).click();
    await undoResponse;
    await expect(page.locator(".thumbnail")).toHaveCount(3);
    await expect(page.locator("#info")).toContainText("Slide 2 / 3");
    await expect(page.locator("#slide-container")).toContainText("First");

    const redoResponse = page.waitForResponse(
      (response) => response.url().endsWith("/api/editor/redo") && response.ok(),
    );
    await page.getByRole("button", { name: "Redo" }).click();
    await redoResponse;
    await expect(page.locator(".thumbnail")).toHaveCount(2);
    await expect(page.locator("#info")).toContainText("Slide 2 / 2");
    await expect(page.locator("#slide-container")).toContainText("Second");

    commandResponse = page.waitForResponse(
      (response) =>
        response.url().endsWith("/api/editor/command") &&
        response.request().method() === "POST" &&
        response.ok(),
    );
    await page.getByTestId("delete-slide-0").click();
    await commandResponse;
    await expect(page.locator(".thumbnail")).toHaveCount(1);
    await expect(page.locator("#info")).toContainText("Slide 1 / 1");
    await expect(page.locator("#slide-container")).toContainText("Second");
    await expect(page.getByTestId("delete-slide-0")).toBeDisabled();

    await page.getByRole("button", { name: "Save" }).click();
    await expect(page.locator("#editor-message")).toContainText(savedPath);
    const saved = readPptx(await readFile(savedPath));
    expect(saved.presentation.slidePartPaths).toEqual(["ppt/slides/slide2.xml"]);
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
  const commandResponse = page.waitForResponse(
    (response) =>
      response.url().endsWith("/api/editor/command") &&
      response.request().method() === "POST" &&
      response.ok(),
  );
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(end.x, end.y, { steps: 6 });
  await page.mouse.up();
  await commandResponse;
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

async function buildTwoSlideFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
        `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
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
        `<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/><p:sldId id="257" r:id="rIdSlide2"/></p:sldIdLst>` +
        `<p:sldSz cx="9144000" cy="5143500"/>` +
        `</p:presentation>`,
    ),
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
        `</Relationships>`,
    ),
  );
  zip.file("ppt/slides/slide1.xml", textSlideXml(10, "First"));
  zip.file("ppt/slides/slide2.xml", textSlideXml(20, "Second"));

  return zip.generateAsync({ type: "uint8array" });
}

function textSlideXml(shapeId: number, text: string): Uint8Array {
  return xml(
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<p:cSld><p:spTree>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="${String(shapeId)}" name="${text}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>` +
      `</p:sp>` +
      `</p:spTree></p:cSld>` +
      `</p:sld>`,
  );
}

function firstShape(source: PptxSourceModel): SourceShape {
  const shape = source.slides[0]?.shapes.find((node): node is SourceShape => node.kind === "shape");
  if (shape === undefined) throw new Error("fixture shape not found");
  return shape;
}

function connectorByName(source: PptxSourceModel, name: string): SourceConnector {
  const connector = source.slides[0]?.shapes.find(
    (node): node is SourceConnector => node.kind === "connector" && node.name === name,
  );
  if (connector === undefined) throw new Error(`connector not found: ${name}`);
  return connector;
}

function firstText(source: PptxSourceModel): string {
  const run = firstShape(source).textBody?.paragraphs[0]?.runs[0];
  if (run === undefined) throw new Error("fixture text run not found");
  return run.text;
}

function shapeByText(source: PptxSourceModel, text: string): SourceShape {
  const shape = findShapeByText(source, text);
  if (shape === undefined) throw new Error(`shape text not found: ${text}`);
  return shape;
}

function findShapeByText(source: PptxSourceModel, text: string): SourceShape | undefined {
  return source.slides[0]?.shapes.find(
    (node): node is SourceShape =>
      node.kind === "shape" &&
      node.textBody?.paragraphs.some((paragraph) =>
        paragraph.runs.some((run) => run.text === text),
      ) === true,
  );
}

function firstRunProperties(source: PptxSourceModel) {
  const run = firstShape(source).textBody?.paragraphs[0]?.runs[0];
  if (run === undefined) throw new Error("fixture text run not found");
  return run.properties;
}

async function findFreePort(): Promise<number> {
  const server = createServer();
  await new Promise<void>((resolve) => {
    server.listen(0, resolve);
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
