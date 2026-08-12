import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

import { build } from "esbuild";
import { describe, expect, it } from "vitest";

const packageRoot = resolve(import.meta.dirname, "..");
const corePackageRoot = resolve(packageRoot, "../core");

interface CorePackageJson {
  readonly exports: {
    readonly ".": {
      readonly browser: {
        readonly import: string;
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

describe("editor React browser entry", () => {
  it("evaluates without a DOM and exposes the client root API", async () => {
    const entry = await import("./index.js");
    const source = await readFile(resolve(packageRoot, "src/index.ts"), "utf8");

    expect(source.startsWith('"use client";')).toBe(true);
    expect(entry).toHaveProperty("PptxEditor");
    expect(entry).toHaveProperty("PptxEditorController");
    expect(entry).toHaveProperty("usePptxEditorController");
  });

  it("bundles as browser ESM without Node built-ins or bundled peers", async () => {
    const corePackageJson: unknown = JSON.parse(
      await readFile(resolve(corePackageRoot, "package.json"), "utf8"),
    );
    if (!isCorePackageJson(corePackageJson)) {
      throw new Error("pptx-glimpse does not expose a browser import target");
    }
    const browserExport = corePackageJson.exports["."].browser.import;
    expect(browserExport).toBe("./dist/browser.js");

    const result = await build({
      stdin: {
        contents:
          'import { PptxEditor, PptxEditorController, usePptxEditorController } from "@pptx-glimpse/editor-react"; console.log(PptxEditor, PptxEditorController, usePptxEditorController);',
        resolveDir: import.meta.dirname,
        sourcefile: "editor-react-browser-smoke.ts",
        loader: "ts",
      },
      bundle: true,
      write: false,
      format: "esm",
      platform: "browser",
      conditions: ["browser", "import"],
      external: ["react", "react-dom", "react/jsx-runtime"],
      logLevel: "silent",
      absWorkingDir: resolve(packageRoot, "../.."),
      plugins: [
        {
          name: "editor-react-browser-workspace-entries",
          setup(browserBuild) {
            browserBuild.onResolve({ filter: /^@pptx-glimpse\/editor-react$/ }, () => ({
              path: resolve(packageRoot, "src/index.ts"),
            }));
            browserBuild.onResolve({ filter: /^pptx-glimpse$/ }, () => ({
              path: resolve(
                corePackageRoot,
                browserExport.replace("./dist/", "src/").replace(/\.js$/, ".ts"),
              ),
            }));
            browserBuild.onResolve({ filter: /^@pptx-glimpse\/document$/ }, () => ({
              path: resolve(corePackageRoot, "../document/src/index.ts"),
            }));
            browserBuild.onResolve({ filter: /^@pptx-glimpse\/editor$/ }, () => ({
              path: resolve(corePackageRoot, "../editor/src/index.ts"),
            }));
            browserBuild.onResolve({ filter: /^@pptx-glimpse\/renderer$/ }, () => ({
              path: resolve(corePackageRoot, "../renderer/src/index.ts"),
            }));
            browserBuild.onResolve({ filter: /^@pptx-glimpse\/renderer\/png$/ }, () => ({
              path: resolve(corePackageRoot, "../renderer/src/png.ts"),
            }));
            browserBuild.onResolve({ filter: /^@pptx-glimpse\/renderer\/png\/browser$/ }, () => ({
              path: resolve(corePackageRoot, "../renderer/src/png-browser.ts"),
            }));
          },
        },
      ],
    });

    const bundled = result.outputFiles[0]?.text ?? "";
    expect(bundled).toMatch(/from "react(?:\/jsx-runtime)?"/);
    expect(bundled).not.toContain('from "pptx-glimpse"');
    expect(bundled).not.toMatch(
      /(?:from|import\()\s*["']node:|require\(["'](?:fs|path|os|buffer|module)(?:\/|["'])/,
    );
  });
});
