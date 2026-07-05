import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { build } from "esbuild";
import { describe, expect, it, vi } from "vitest";

const resvgMocks = vi.hoisted(() => ({
  initWasm: vi.fn().mockResolvedValue(undefined),
  MockResvg: class {
    render() {
      return {
        asPng: () => new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
        width: 1,
        height: 1,
      };
    }
  },
}));

vi.mock("@resvg/resvg-wasm", () => ({
  initWasm: resvgMocks.initWasm,
  Resvg: resvgMocks.MockResvg,
}));

const here = dirname(fileURLToPath(import.meta.url));
const packageRoot = resolve(here, "..");

interface CorePackageJson {
  exports: {
    ".": {
      browser: {
        import: string;
      };
    };
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isCorePackageJson(value: unknown): value is CorePackageJson {
  if (!isRecord(value) || !isRecord(value.exports)) return false;
  const rootExport = value.exports["."];
  if (!isRecord(rootExport) || !isRecord(rootExport.browser)) return false;
  return typeof rootExport.browser.import === "string";
}

describe("browser entry", () => {
  it("forwards explicit WASM input to the renderer PNG initializer", async () => {
    resvgMocks.initWasm.mockClear();
    const { initResvgWasm } = await import("./browser.js");
    const wasm = new Uint8Array([1, 2, 3]);

    await initResvgWasm(wasm);

    expect(resvgMocks.initWasm).toHaveBeenCalledWith(wasm);
  });

  it("bundles browser-safe entry APIs without Node built-ins", async () => {
    const packageJson: unknown = JSON.parse(
      readFileSync(resolve(packageRoot, "package.json"), "utf8"),
    );
    if (!isCorePackageJson(packageJson)) {
      throw new Error("packages/core/package.json does not expose a browser import target");
    }
    const browserExport = packageJson.exports["."].browser.import;
    expect(browserExport).toBe("./dist/browser.js");

    const result = await build({
      stdin: {
        contents:
          'import { convertPptxToPng, convertPptxToSvg, initResvgWasm } from "pptx-glimpse"; console.log(convertPptxToPng, convertPptxToSvg, initResvgWasm);',
        resolveDir: here,
        sourcefile: "browser-entry-smoke.ts",
        loader: "ts",
      },
      bundle: true,
      write: false,
      format: "esm",
      platform: "browser",
      conditions: ["browser", "import"],
      logLevel: "silent",
      absWorkingDir: resolve(here, "../../.."),
      plugins: [
        {
          name: "workspace-pptx-glimpse-browser-entry",
          setup(build) {
            build.onResolve({ filter: /^pptx-glimpse$/ }, () => ({
              path: resolve(
                packageRoot,
                browserExport.replace("./dist/", "src/").replace(/\.js$/, ".ts"),
              ),
            }));
            build.onResolve({ filter: /^@pptx-glimpse\/document$/ }, () => ({
              path: resolve(packageRoot, "../document/src/index.ts"),
            }));
            build.onResolve({ filter: /^@pptx-glimpse\/renderer$/ }, () => ({
              path: resolve(packageRoot, "../renderer/src/index.ts"),
            }));
            build.onResolve({ filter: /^@pptx-glimpse\/renderer\/png$/ }, () => ({
              path: resolve(packageRoot, "../renderer/src/png.ts"),
            }));
            build.onResolve({ filter: /^@pptx-glimpse\/renderer\/png\/browser$/ }, () => ({
              path: resolve(packageRoot, "../renderer/src/png-browser.ts"),
            }));
          },
        },
      ],
    });

    const bundled = result.outputFiles[0].text;
    expect(bundled).not.toMatch(
      /(?:node:fs|node:path|node:os|node:buffer|fs\/promises|from "fs"|from "path"|from "os"|from "module")/,
    );
  });
});
