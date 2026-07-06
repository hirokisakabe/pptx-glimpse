import { Buffer } from "node:buffer";
import { mkdtemp, readFile, rm, symlink, writeFile } from "node:fs/promises";
import { createServer, type Server } from "node:http";
import { tmpdir } from "node:os";
import { join } from "node:path";

import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";

import {
  type PptxSourceModel,
  readPptx,
  type SourceConnector,
  type SourceTextRun,
} from "../packages/document/src/index.js";
import { createDevServerRequestHandler, DevEditorBackend } from "../scripts/dev-server.js";
import { unsafeScriptInputAssertion } from "../scripts/unsafe-type-assertion.js";

const encoder = new TextEncoder();
const RED_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEUlEQVR4nGP8z4AATEhsPBwAM9EBBzDn4UwAAAAASUVORK5CYII=";
const BLUE_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAE0lEQVR4nGNkYPjPAANMcBZeDgAx0wEH1s7nlgAAAABJRU5ErkJggg==";
const RED_PNG = pngBytes(RED_PNG_BASE64);
const BLUE_PNG = pngBytes(BLUE_PNG_BASE64);

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

function pngBytes(base64: string): Uint8Array {
  return new Uint8Array(Buffer.from(base64, "base64"));
}

describe("dev server editor API", () => {
  const servers: Server[] = [];

  afterEach(async () => {
    await Promise.all(servers.map((server) => closeServer(server)));
    servers.length = 0;
  });

  it("applies text commands, undo/redo, exposes bounds, and saves edited PPTX", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-dev-server-test-"));
    let outsideDir: string | undefined;
    try {
      const sourcePath = join(dir, "fixture.pptx");
      const savedPath = join(dir, "saved.pptx");
      await writeFile(sourcePath, await buildTextEditFixture());

      const defaultRenderedBackend = await DevEditorBackend.load(sourcePath);
      expect(defaultRenderedBackend.slides[0].svg).toContain("<svg");
      expect(defaultRenderedBackend.slides[0].svg).toContain("<text");

      const backend = await DevEditorBackend.load(sourcePath, renderPreview);
      const server = createServer(createDevServerRequestHandler(backend, "fixture.pptx"));
      servers.push(server);
      const baseUrl = await listen(server);

      const shapes = await getJson<ShapesResponse>(`${baseUrl}/api/editor/shapes?slide=1`);
      expect(shapes.shapes[0]).toMatchObject({
        id: "10",
        name: "Title",
        bounds: { x: 96, y: 192, width: 288, height: 96 },
      });
      expect(shapes.shapes[0].textRuns[0]).toMatchObject({ text: "Original" });
      expect(shapes.shapes[0].editableTextBody?.docJson).toMatchObject({
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [{ type: "text", text: "Original" }],
          },
        ],
      });

      const edited = await postJson<SlidesResponse>(`${baseUrl}/api/editor/command`, {
        command: {
          kind: "replaceTextRunPlainText",
          handle: shapes.shapes[0].textRuns[0].handle,
          text: "Edited",
        },
      });
      expect(edited.slides[0].svg).toContain("Edited");
      expect(edited.history).toMatchObject({ canUndo: true, canRedo: false });

      const undone = await postJson<SlidesResponse>(`${baseUrl}/api/editor/undo`, {});
      expect(undone.slides[0].svg).toContain("Original");
      expect(undone.history).toMatchObject({ canUndo: false, canRedo: true });

      const redone = await postJson<SlidesResponse>(`${baseUrl}/api/editor/redo`, {});
      expect(redone.slides[0].svg).toContain("Edited");
      expect(redone.history).toMatchObject({ canUndo: true, canRedo: false });

      const overlayEdited = await postJson<SlidesResponse>(`${baseUrl}/api/editor/text-body`, {
        handle: shapes.shapes[0].handle,
        docJson: {
          ...shapes.shapes[0].editableTextBody?.docJson,
          content: [
            {
              ...shapes.shapes[0].editableTextBody?.docJson.content?.[0],
              content: [
                {
                  ...shapes.shapes[0].editableTextBody?.docJson.content?.[0]?.content?.[0],
                  text: "Overlay edited",
                },
              ],
            },
          ],
        },
      });
      expect(overlayEdited.slides[0].svg).toContain("Overlay edited");

      await postJson<SlidesResponse>(`${baseUrl}/api/editor/text-body`, {
        handle: shapes.shapes[0].handle,
        docJson: {
          ...shapes.shapes[0].editableTextBody?.docJson,
          content: [
            {
              ...shapes.shapes[0].editableTextBody?.docJson.content?.[0],
              content: [],
            },
          ],
        },
      });

      const saved = await postJson<SaveResponse>(`${baseUrl}/api/editor/save`, { path: savedPath });
      expect(saved).toMatchObject({ ok: true, path: savedPath });
      expect(firstRun(readPptx(await readFile(savedPath)))).toBe("");

      const rejectedSave = await postJsonError(`${baseUrl}/api/editor/save`, {
        path: join(tmpdir(), "outside-dev-server-save.pptx"),
      });
      expect(rejectedSave.error).toMatch(/inside the source PPTX directory/);

      const rejectedExtension = await postJsonError(`${baseUrl}/api/editor/save`, {
        path: join(dir, "saved.txt"),
      });
      expect(rejectedExtension.error).toMatch(/\.pptx extension/);

      outsideDir = await mkdtemp(join(tmpdir(), "pptx-glimpse-outside-save-"));
      const linkPath = join(dir, "outside-link");
      await symlink(outsideDir, linkPath);
      const rejectedSymlinkDir = await postJsonError(`${baseUrl}/api/editor/save`, {
        path: join(linkPath, "saved.pptx"),
      });
      expect(rejectedSymlinkDir.error).toMatch(/inside the source PPTX directory/);
    } finally {
      await rm(dir, { recursive: true, force: true });
      if (outsideDir !== undefined) {
        await rm(outsideDir, { recursive: true, force: true });
      }
    }
  });

  it("applies, undoes, redoes, clears, and saves text run property commands", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-dev-server-run-property-test-"));
    try {
      const sourcePath = join(dir, "fixture.pptx");
      const savedPath = join(dir, "saved.pptx");
      const clearedPath = join(dir, "cleared.pptx");
      await writeFile(sourcePath, await buildTextEditFixture());

      const backend = await DevEditorBackend.load(sourcePath, renderPreview);
      const server = createServer(createDevServerRequestHandler(backend, "fixture.pptx"));
      servers.push(server);
      const baseUrl = await listen(server);

      const shapes = await getJson<ShapesResponse>(`${baseUrl}/api/editor/shapes?slide=1`);
      const handle = shapes.shapes[0].textRuns[0].handle;

      const decorated = await postJson<SlidesResponse>(`${baseUrl}/api/editor/command`, {
        command: {
          kind: "setTextRunProperties",
          handle,
          properties: {
            bold: true,
            italic: true,
            underline: true,
            fontSize: 32,
            color: { kind: "srgb", hex: "9c0000" },
            typeface: "Liberation Sans",
          },
        },
      });
      expect(decorated.history).toMatchObject({ canUndo: true, canRedo: false });
      expect(decorated.slides[0].svg).toContain('data-bold="true"');
      expect(decorated.slides[0].svg).toContain('data-italic="true"');
      expect(decorated.slides[0].svg).toContain('data-underline="true"');
      expect(decorated.slides[0].svg).toContain('data-font-size="32"');
      expect(decorated.slides[0].svg).toContain('data-color="9C0000"');
      expect(decorated.slides[0].svg).toContain('data-typeface="Liberation Sans"');

      const undone = await postJson<SlidesResponse>(`${baseUrl}/api/editor/undo`, {});
      expect(undone.history).toMatchObject({ canUndo: false, canRedo: true });
      expect(undone.slides[0].svg).not.toContain('data-bold="true"');
      expect(undone.slides[0].svg).not.toContain('data-italic="true"');
      expect(undone.slides[0].svg).not.toContain('data-underline="true"');
      expect(undone.slides[0].svg).not.toContain('data-font-size="32"');
      expect(undone.slides[0].svg).not.toContain('data-color="9C0000"');
      expect(undone.slides[0].svg).not.toContain('data-typeface="Liberation Sans"');

      const redone = await postJson<SlidesResponse>(`${baseUrl}/api/editor/redo`, {});
      expect(redone.history).toMatchObject({ canUndo: true, canRedo: false });
      expect(redone.slides[0].svg).toContain('data-bold="true"');
      expect(redone.slides[0].svg).toContain('data-italic="true"');
      expect(redone.slides[0].svg).toContain('data-underline="true"');
      expect(redone.slides[0].svg).toContain('data-font-size="32"');
      expect(redone.slides[0].svg).toContain('data-color="9C0000"');
      expect(redone.slides[0].svg).toContain('data-typeface="Liberation Sans"');

      await postJson<SaveResponse>(`${baseUrl}/api/editor/save`, { path: savedPath });
      expect(firstRunProperties(readPptx(await readFile(savedPath)))).toMatchObject({
        bold: true,
        italic: true,
        underline: true,
        fontSize: 32,
        color: { kind: "srgb", hex: "9C0000" },
        typeface: "Liberation Sans",
      });

      const clearedResponse = await postJson<SlidesResponse>(`${baseUrl}/api/editor/command`, {
        command: {
          kind: "clearTextRunProperties",
          handle,
          properties: ["bold", "italic", "underline", "fontSize", "color", "typeface"],
        },
      });
      expect(clearedResponse.slides[0].svg).not.toContain("data-bold=");
      expect(clearedResponse.slides[0].svg).not.toContain("data-italic=");
      expect(clearedResponse.slides[0].svg).not.toContain("data-underline=");
      expect(clearedResponse.slides[0].svg).not.toContain("data-font-size=");
      expect(clearedResponse.slides[0].svg).not.toContain("data-color=");
      expect(clearedResponse.slides[0].svg).not.toContain("data-typeface=");

      await postJson<SaveResponse>(`${baseUrl}/api/editor/save`, { path: clearedPath });

      const cleared = firstRunProperties(readPptx(await readFile(clearedPath)));
      expect(cleared?.bold).toBeUndefined();
      expect(cleared?.italic).toBeUndefined();
      expect(cleared?.underline).toBeUndefined();
      expect(cleared?.fontSize).toBeUndefined();
      expect(cleared?.color).toBeUndefined();
      expect(cleared?.typeface).toBeUndefined();
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("adds, selects, deletes, undoes, and redoes a text box through the editor API", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-dev-server-shape-test-"));
    try {
      const sourcePath = join(dir, "fixture.pptx");
      await writeFile(sourcePath, await buildTextEditFixture());

      const backend = await DevEditorBackend.load(sourcePath, renderPreview);
      const server = createServer(createDevServerRequestHandler(backend, "fixture.pptx"));
      servers.push(server);
      const baseUrl = await listen(server);

      const added = await postJson<SlidesResponse>(`${baseUrl}/api/editor/add-text-box`, {
        slide: 1,
      });
      expect(added.slides[0].svg).toContain("New text box");
      expect(added.selection?.shapeHandle).toBeDefined();
      expect(added.history).toMatchObject({ canUndo: true, canRedo: false });

      const shapes = await getJson<ShapesResponse>(`${baseUrl}/api/editor/shapes?slide=1`);
      const textBox = shapes.shapes.find((shape) =>
        shape.textRuns.some((run) => run.text === "New text box"),
      );
      expect(textBox).toMatchObject({
        bounds: { x: 96, y: 96, width: 288, height: 72 },
        editableDelete: true,
      });

      const deleted = await postJson<SlidesResponse>(`${baseUrl}/api/editor/command`, {
        command: {
          kind: "deleteShape",
          handle: textBox?.handle,
        },
      });
      expect(deleted.selection).toBeUndefined();
      expect(deleted.slides[0].svg).not.toContain("New text box");

      const undone = await postJson<SlidesResponse>(`${baseUrl}/api/editor/undo`, {});
      expect(undone.slides[0].svg).toContain("New text box");
      expect(undone.history).toMatchObject({ canUndo: true, canRedo: true });

      const redone = await postJson<SlidesResponse>(`${baseUrl}/api/editor/redo`, {});
      expect(redone.slides[0].svg).not.toContain("New text box");
      expect(redone.history).toMatchObject({ canUndo: true, canRedo: false });
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("adds, selects, moves, saves, deletes, undoes, and redoes a connector through the editor API", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-dev-server-connector-add-test-"));
    try {
      const sourcePath = join(dir, "fixture.pptx");
      const savedPath = join(dir, "saved.pptx");
      await writeFile(sourcePath, await buildTextEditFixture());

      const backend = await DevEditorBackend.load(sourcePath, renderPreview);
      const server = createServer(createDevServerRequestHandler(backend, "fixture.pptx"));
      servers.push(server);
      const baseUrl = await listen(server);

      const added = await postJson<SlidesResponse>(`${baseUrl}/api/editor/add-connector`, {
        slide: 1,
      });
      expect(added.selection?.shapeHandle).toBeDefined();
      expect(added.history).toMatchObject({ canUndo: true, canRedo: false });

      const shapes = await getJson<ShapesResponse>(`${baseUrl}/api/editor/shapes?slide=1`);
      const connector = shapes.shapes.find((shape) => shape.kind === "connector");
      expect(connector).toMatchObject({
        name: "Connector 11",
        bounds: { x: 144, y: 144, width: 288, height: 96 },
        editableDelete: true,
      });

      const moved = await postJson<SlidesResponse>(`${baseUrl}/api/editor/command`, {
        command: {
          kind: "setShapeTransform",
          handle: connector?.handle,
          offsetX: 168 * 9525,
          offsetY: 160 * 9525,
          width: 336 * 9525,
          height: 120 * 9525,
        },
      });
      expect(moved.history).toMatchObject({ canUndo: true, canRedo: false });

      await postJson<SaveResponse>(`${baseUrl}/api/editor/save`, { path: savedPath });
      expect(
        connectorByName(readPptx(await readFile(savedPath)), "Connector 11").transform,
      ).toMatchObject({
        offsetX: 168 * 9525,
        offsetY: 160 * 9525,
        width: 336 * 9525,
        height: 120 * 9525,
      });

      const deleted = await postJson<SlidesResponse>(`${baseUrl}/api/editor/command`, {
        command: {
          kind: "deleteShape",
          handle: connector?.handle,
        },
      });
      expect(deleted.selection).toBeUndefined();
      expect(
        (await getJson<ShapesResponse>(`${baseUrl}/api/editor/shapes?slide=1`)).shapes.find(
          (shape) => shape.kind === "connector",
        ),
      ).toBeUndefined();

      const undone = await postJson<SlidesResponse>(`${baseUrl}/api/editor/undo`, {});
      expect(undone.history).toMatchObject({ canUndo: true, canRedo: true });
      expect(
        (await getJson<ShapesResponse>(`${baseUrl}/api/editor/shapes?slide=1`)).shapes.find(
          (shape) => shape.kind === "connector",
        ),
      ).toBeDefined();

      const redone = await postJson<SlidesResponse>(`${baseUrl}/api/editor/redo`, {});
      expect(redone.history).toMatchObject({ canUndo: true, canRedo: false });
      expect(
        (await getJson<ShapesResponse>(`${baseUrl}/api/editor/shapes?slide=1`)).shapes.find(
          (shape) => shape.kind === "connector",
        ),
      ).toBeUndefined();
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("does not expose connector-referenced shapes as deletable through the editor API", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-dev-server-connector-test-"));
    try {
      const sourcePath = join(dir, "fixture.pptx");
      await writeFile(sourcePath, await buildTextEditFixture({ includeConnector: true }));

      const backend = await DevEditorBackend.load(sourcePath, renderPreview);
      const server = createServer(createDevServerRequestHandler(backend, "fixture.pptx"));
      servers.push(server);
      const baseUrl = await listen(server);

      const shapes = await getJson<ShapesResponse>(`${baseUrl}/api/editor/shapes?slide=1`);
      const connectedShape = shapes.shapes.find((shape) => shape.name === "Title");
      expect(connectedShape).toMatchObject({
        bounds: { x: 96, y: 192, width: 288, height: 96 },
      });
      expect(connectedShape?.editableDelete).toBeUndefined();

      const rejectedDelete = await postJsonError(`${baseUrl}/api/editor/command`, {
        command: {
          kind: "deleteShape",
          handle: connectedShape?.handle,
        },
      });
      expect(rejectedDelete.error).toMatch(/referenced by connector/);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("applies image replacement commands with shared media warnings, undo, redo, and save", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-dev-server-image-test-"));
    try {
      const sourcePath = join(dir, "fixture.pptx");
      const savedPath = join(dir, "saved.pptx");
      await writeFile(sourcePath, await buildImageFixture());

      const backend = await DevEditorBackend.load(sourcePath, renderImagePreview);
      const server = createServer(createDevServerRequestHandler(backend, "fixture.pptx"));
      servers.push(server);
      const baseUrl = await listen(server);

      const shapes = await getJson<ShapesResponse>(`${baseUrl}/api/editor/shapes?slide=1`);
      const image = shapes.shapes.find((shape) => shape.kind === "image");
      expect(image?.editableImageReplacement).toEqual({
        contentType: "image/png",
        accept: "image/png,.png",
        mediaPartPath: "ppt/media/image1.png",
        sharedReferenceCount: 2,
      });

      const rejected = await postJsonError(`${baseUrl}/api/editor/command`, {
        command: {
          kind: "replaceImage",
          handle: shapes.shapes[0].handle,
          bytes: [0xff, 0xd8, 0xff, 0xd9],
        },
      });
      expect(rejected.error).toContain("does not match existing media content type");

      const replaced = await postJson<SlidesResponse>(`${baseUrl}/api/editor/command`, {
        command: {
          kind: "replaceImage",
          handle: image?.handle,
          bytes: Array.from(BLUE_PNG),
        },
      });
      expect(replaced.slides[0].svg).toContain(BLUE_PNG_BASE64);
      expect(replaced.warnings).toEqual([
        expect.objectContaining({
          code: "shared-media-part",
          mediaPartPath: "ppt/media/image1.png",
          referenceCount: 2,
        }),
      ]);

      const undone = await postJson<SlidesResponse>(`${baseUrl}/api/editor/undo`, {});
      expect(undone.slides[0].svg).toContain(RED_PNG_BASE64);
      const redone = await postJson<SlidesResponse>(`${baseUrl}/api/editor/redo`, {});
      expect(redone.slides[0].svg).toContain(BLUE_PNG_BASE64);

      await postJson<SaveResponse>(`${baseUrl}/api/editor/save`, { path: savedPath });
      expect(mediaBytes(readPptx(await readFile(savedPath)), "ppt/media/image1.png")).toEqual(
        BLUE_PNG,
      );
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});

interface ShapesResponse {
  shapes: Array<{
    handle: unknown;
    id: string;
    kind: string;
    name?: string;
    bounds?: { x: number; y: number; width: number; height: number };
    editableDelete?: boolean;
    textRuns: Array<{ text: string; handle: unknown }>;
    editableTextBody?: {
      docJson: {
        type: "doc";
        content?: Array<{
          type: "paragraph";
          attrs?: unknown;
          content?: Array<{ type: "text"; text: string; marks?: unknown }>;
        }>;
      };
    };
    editableImageReplacement?: {
      contentType: string;
      accept: string;
      mediaPartPath: string;
      sharedReferenceCount: number;
    };
  }>;
}

interface SlidesResponse {
  slides: Array<{ slideNumber: number; svg: string }>;
  history: { canUndo: boolean; canRedo: boolean };
  selection?: { shapeHandle: unknown };
  warnings?: Array<{
    code: string;
    mediaPartPath: string;
    referenceCount: number;
  }>;
}

interface SaveResponse {
  ok: true;
  path: string;
}

function renderPreview(input: Uint8Array): Promise<Array<{ slideNumber: number; svg: string }>> {
  const source = readPptx(input);
  return Promise.resolve([
    {
      slideNumber: 1,
      svg: `<svg>${allRuns(source)
        .map((run) => `<text${runPropertyAttrs(run)}>${escapeXml(run.text)}</text>`)
        .join("")}</svg>`,
    },
  ]);
}

function renderImagePreview(
  input: Uint8Array,
): Promise<Array<{ slideNumber: number; svg: string }>> {
  const source = readPptx(input);
  return Promise.resolve([
    {
      slideNumber: 1,
      svg: `<svg><image href="data:image/png;base64,${Buffer.from(
        mediaBytes(source, "ppt/media/image1.png"),
      ).toString("base64")}"/></svg>`,
    },
  ]);
}

function runPropertyAttrs(run: SourceTextRun): string {
  const properties = run.properties;
  return [
    properties?.bold === undefined ? "" : ` data-bold="${String(properties.bold)}"`,
    properties?.italic === undefined ? "" : ` data-italic="${String(properties.italic)}"`,
    properties?.underline === undefined ? "" : ` data-underline="${String(properties.underline)}"`,
    properties?.fontSize === undefined ? "" : ` data-font-size="${String(properties.fontSize)}"`,
    properties?.color === undefined ? "" : ` data-color="${escapeXml(properties.color.hex)}"`,
    properties?.typeface === undefined ? "" : ` data-typeface="${escapeXml(properties.typeface)}"`,
  ].join("");
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

async function listen(server: Server): Promise<string> {
  await new Promise<void>((resolve) => {
    server.listen(0, "127.0.0.1", resolve);
  });
  const address = server.address();
  if (address === null || typeof address === "string") {
    throw new Error("test server did not expose a TCP port");
  }
  return `http://127.0.0.1:${String(address.port)}`;
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

async function getJson<T>(url: string): Promise<T> {
  const response = await fetch(url);
  return parseJson<T>(response);
}

async function postJson<T>(url: string, body: unknown): Promise<T> {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return parseJson<T>(response);
}

async function postJsonError(url: string, body: unknown): Promise<{ error: string }> {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const json: unknown = await response.json();
  expect(response.ok).toBe(false);
  return unsafeScriptInputAssertion<{ error: string }>(json);
}

async function parseJson<T>(response: Response): Promise<T> {
  const json: unknown = await response.json();
  if (!response.ok) {
    throw new Error(JSON.stringify(json));
  }
  return unsafeScriptInputAssertion<T>(json);
}

async function buildTextEditFixture(
  options: { includeConnector?: boolean } = {},
): Promise<Uint8Array> {
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
        `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/>` +
        `<a:p><a:r><a:t>Original</a:t></a:r></a:p>` +
        `</p:txBody></p:sp>` +
        (options.includeConnector
          ? `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="11" name="Connector"/><p:cNvCxnSpPr>` +
            `<a:stCxn id="10" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>` +
            `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
            `<a:prstGeom prst="straightConnector1"/></p:spPr></p:cxnSp>`
          : "") +
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

function firstRun(source: PptxSourceModel): string {
  const run = allRuns(source)[0];
  if (run === undefined) throw new Error("fixture text run not found");
  return run.text;
}

function firstRunProperties(source: PptxSourceModel) {
  const run = allRuns(source)[0];
  if (run === undefined) throw new Error("fixture text run not found");
  return run.properties;
}

function allRuns(source: PptxSourceModel) {
  return source.slides[0].shapes.flatMap((node) => {
    if (node.kind !== "shape") return [];
    return node.textBody?.paragraphs.flatMap((paragraph) => paragraph.runs) ?? [];
  });
}

function mediaBytes(source: PptxSourceModel, partPath: string): Uint8Array {
  const media = source.packageGraph.media.find((part) => part.partPath === partPath);
  if (media === undefined) throw new Error(`media not found: ${partPath}`);
  return media.bytes;
}

function connectorByName(source: PptxSourceModel, name: string): SourceConnector {
  const connector = source.slides[0]?.shapes.find(
    (node): node is SourceConnector => node.kind === "connector" && node.name === name,
  );
  if (connector === undefined) throw new Error(`connector not found: ${name}`);
  return connector;
}
