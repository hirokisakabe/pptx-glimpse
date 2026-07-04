import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { createServer, type Server } from "node:http";
import { tmpdir } from "node:os";
import { join } from "node:path";

import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";

import {
  type PptxSourceModel,
  readPptx,
  type SourceShape,
} from "../packages/document/src/index.js";
import { createDevServerRequestHandler, DevEditorBackend } from "../scripts/dev-server.js";
import { unsafeScriptInputAssertion } from "../scripts/unsafe-type-assertion.js";

const encoder = new TextEncoder();

function xml(content: string): Uint8Array {
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`);
}

describe("dev server editor API", () => {
  const servers: Server[] = [];

  afterEach(async () => {
    await Promise.all(servers.map((server) => closeServer(server)));
    servers.length = 0;
  });

  it("applies text commands, undo/redo, exposes bounds, and saves edited PPTX", async () => {
    const dir = await mkdtemp(join(tmpdir(), "pptx-glimpse-dev-server-test-"));
    try {
      const sourcePath = join(dir, "fixture.pptx");
      const savedPath = join(dir, "saved.pptx");
      await writeFile(sourcePath, await buildTextEditFixture());

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

      const saved = await postJson<SaveResponse>(`${baseUrl}/api/editor/save`, { path: savedPath });
      expect(saved).toMatchObject({ ok: true, path: savedPath });
      expect(firstRun(readPptx(await readFile(savedPath)))).toBe("Edited");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});

interface ShapesResponse {
  shapes: Array<{
    id: string;
    name?: string;
    bounds?: { x: number; y: number; width: number; height: number };
    textRuns: Array<{ text: string; handle: unknown }>;
  }>;
}

interface SlidesResponse {
  slides: Array<{ slideNumber: number; svg: string }>;
  history: { canUndo: boolean; canRedo: boolean };
}

interface SaveResponse {
  ok: true;
  path: string;
}

function renderPreview(input: Uint8Array): Promise<Array<{ slideNumber: number; svg: string }>> {
  const source = readPptx(input);
  return Promise.resolve([
    { slideNumber: 1, svg: `<svg><text>${escapeXml(firstRun(source))}</text></svg>` },
  ]);
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

async function parseJson<T>(response: Response): Promise<T> {
  const json: unknown = await response.json();
  if (!response.ok) {
    throw new Error(JSON.stringify(json));
  }
  return unsafeScriptInputAssertion<T>(json);
}

async function buildTextEditFixture(): Promise<Uint8Array> {
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
        `</p:spTree></p:cSld>` +
        `</p:sld>`,
    ),
  );

  return zip.generateAsync({ type: "uint8array" });
}

function firstRun(source: PptxSourceModel): string {
  const shape = source.slides[0].shapes.find((node): node is SourceShape => node.kind === "shape");
  const run = shape?.textBody?.paragraphs[0].runs[0];
  if (run === undefined) throw new Error("fixture text run not found");
  return run.text;
}
