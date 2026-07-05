import { execFile } from "node:child_process";
import { rm, writeFile } from "node:fs/promises";
import { createServer, type Server } from "node:http";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { expect, test } from "@playwright/test";
import { build } from "esbuild";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const corePackageRoot = resolve(repoRoot, "packages/core");
const sharedFixtures = resolve(repoRoot, "shared-fixtures");
const execFileAsync = promisify(execFile);
let coreDistBuildPromise: Promise<void> | null = null;

test("runs a browser-only PPTX to SVG viewer for shared fixtures", async ({ page }) => {
  const viewer = await startStandaloneViewer();
  try {
    await page.goto(viewer.url);

    for (const fixtureName of ["real-basic-theme.pptx", "real-product-page.pptx"]) {
      await page.getByTestId("pptx-input").setInputFiles(resolve(sharedFixtures, fixtureName));
      await expect(page.getByTestId("status")).toContainText("slides rendered");
      await expect(page.getByTestId("slide").first()).toBeVisible();
      await expect(page.locator("svg").first()).toBeVisible();

      const slideCount = await page.getByTestId("slide").count();
      expect(slideCount).toBeGreaterThan(0);
    }
  } finally {
    await viewer.close();
  }
});

test("passes browser-loaded font ArrayBuffers into conversion", async ({ page }) => {
  const viewer = await startStandaloneViewer();
  const fontPath = join(tmpdir(), `pptx-glimpse-browser-font-${Date.now().toString()}.ttf`);
  try {
    await writeFile(fontPath, Buffer.from(await createTestFontBuffer("BrowserSmokeFont")));
    await page.goto(viewer.url);

    await page.getByTestId("font-input").setInputFiles(fontPath);
    await expect(page.getByTestId("font-count")).toContainText("1 font file ready");

    await page
      .getByTestId("pptx-input")
      .setInputFiles(resolve(sharedFixtures, "real-basic-theme.pptx"));
    await expect(page.getByTestId("status")).toContainText("slides rendered with 1 font file");

    const fontCount = await page.evaluate(() => window.__pptxGlimpseSmoke?.fontCount);
    expect(fontCount).toBe(1);
  } finally {
    await rm(fontPath, { force: true });
    await viewer.close();
  }
});

interface ViewerServer {
  readonly url: string;
  close(): Promise<void>;
}

async function startStandaloneViewer(): Promise<ViewerServer> {
  const appBundle = await buildStandaloneViewerBundle();
  const server = createServer((request, response) => {
    const url = new URL(request.url ?? "/", "http://127.0.0.1");
    if (url.pathname === "/" || url.pathname === "/index.html") {
      response.setHeader("Content-Type", "text/html; charset=utf-8");
      response.end(viewerHtml);
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
  return {
    url,
    close: () => closeServer(server),
  };
}

async function buildStandaloneViewerBundle(): Promise<string> {
  await ensureCoreDist();
  const result = await build({
    stdin: {
      contents: viewerAppSource,
      resolveDir: here,
      sourcefile: "browser-standalone-viewer.ts",
      loader: "ts",
    },
    bundle: true,
    write: false,
    format: "esm",
    platform: "browser",
    conditions: ["browser", "import"],
    logLevel: "silent",
    absWorkingDir: repoRoot,
    plugins: [
      {
        name: "workspace-pptx-glimpse-browser-entry",
        setup(buildContext) {
          buildContext.onResolve({ filter: /^pptx-glimpse$/ }, () => ({
            path: resolve(corePackageRoot, "dist/browser.js"),
          }));
        },
      },
    ],
  });
  const bundled = result.outputFiles[0].text;
  expect(bundled).not.toMatch(
    /(?:node:fs|node:path|node:os|node:buffer|fs\/promises|from "fs"|from "path"|from "os"|from "module")/,
  );
  return bundled;
}

async function ensureCoreDist(): Promise<void> {
  coreDistBuildPromise ??= execFileAsync("pnpm", ["--filter", "pptx-glimpse", "build"], {
    cwd: repoRoot,
    maxBuffer: 10 * 1024 * 1024,
  }).then(() => undefined);
  await coreDistBuildPromise;
}

const viewerHtml = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>pptx-glimpse browser smoke viewer</title>
    <style>
      body { font-family: system-ui, sans-serif; margin: 24px; background: #f7f8f4; color: #17212b; }
      main { max-width: 960px; margin: 0 auto; }
      .controls { display: flex; gap: 12px; margin-bottom: 16px; }
      .slide { background: white; border: 1px solid #cfd8dc; margin-top: 16px; }
      .slide svg { display: block; width: 100%; height: auto; }
    </style>
  </head>
  <body>
    <main>
      <div class="controls">
        <input data-testid="pptx-input" id="pptx-input" type="file" accept=".pptx">
        <input data-testid="font-input" id="font-input" type="file" accept=".ttf,.otf,.ttc" multiple>
      </div>
      <p data-testid="font-count" id="font-count">0 font files ready</p>
      <p data-testid="status" id="status">Ready</p>
      <section id="slides"></section>
    </main>
    <script type="module" src="/app.js"></script>
  </body>
</html>`;

const viewerAppSource = `
import { convertPptxToSvg } from "pptx-glimpse";

const pptxInput = document.getElementById("pptx-input");
const fontInput = document.getElementById("font-input");
const fontCount = document.getElementById("font-count");
const status = document.getElementById("status");
const slides = document.getElementById("slides");

let fonts = [];

fontInput.addEventListener("change", async () => {
  fonts = await Promise.all(
    Array.from(fontInput.files ?? []).map(async (file) => ({
      name: file.name.replace(/\\.(?:ttf|otf|ttc)$/i, ""),
      data: await file.arrayBuffer(),
    })),
  );
  fontCount.textContent = fonts.length + " font file" + (fonts.length === 1 ? "" : "s") + " ready";
});

pptxInput.addEventListener("change", async () => {
  const file = pptxInput.files?.[0];
  if (!file) return;
  status.textContent = "Converting";
  slides.textContent = "";
  const input = new Uint8Array(await file.arrayBuffer());
  const report = await convertPptxToSvg(input, { fonts, skipSystemFonts: true });
  window.__pptxGlimpseSmoke = { fontCount: fonts.length, slideCount: report.slides.length };
  slides.innerHTML = report.slides
    .map((slide) => '<article class="slide" data-testid="slide">' + slide.svg + "</article>")
    .join("");
  status.textContent =
    report.slides.length +
    " slides rendered with " +
    fonts.length +
    " font file" +
    (fonts.length === 1 ? "" : "s");
});
`;

async function createTestFontBuffer(familyName: string): Promise<ArrayBuffer> {
  const opentype: OpentypeTestModule = await import("opentype.js");
  const notdefGlyph = new opentype.Glyph({
    name: ".notdef",
    advanceWidth: 650,
    path: new opentype.Path(),
  });
  const spaceGlyph = new opentype.Glyph({
    name: "space",
    unicode: 32,
    advanceWidth: 250,
    path: new opentype.Path(),
  });
  const glyphPath = new opentype.Path();
  glyphPath.moveTo(0, 0);
  glyphPath.lineTo(300, 800);
  glyphPath.lineTo(600, 0);
  glyphPath.close();
  const letterGlyph = new opentype.Glyph({
    name: "A",
    unicode: 65,
    advanceWidth: 600,
    path: glyphPath,
  });
  const font = new opentype.Font({
    familyName,
    styleName: "Regular",
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    glyphs: [notdefGlyph, spaceGlyph, letterGlyph],
  });
  return font.toArrayBuffer();
}

interface OpentypeTestModule {
  Glyph: new (opts: Record<string, unknown>) => unknown;
  Path: new () => {
    moveTo(x: number, y: number): void;
    lineTo(x: number, y: number): void;
    close(): void;
  };
  Font: new (opts: Record<string, unknown>) => { toArrayBuffer(): ArrayBuffer };
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
    __pptxGlimpseSmoke?: {
      fontCount: number;
      slideCount: number;
    };
  }
}
