import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { build } from "esbuild";
import { describe, expect, it } from "vitest";

const here = dirname(fileURLToPath(import.meta.url));

describe("browser entry", () => {
  it("bundles convertPptxToSvg for browser without Node built-ins", async () => {
    const result = await build({
      stdin: {
        contents: 'import { convertPptxToSvg } from "./browser.ts"; console.log(convertPptxToSvg);',
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
    });

    const bundled = result.outputFiles[0].text;
    expect(bundled).not.toMatch(
      /(?:node:fs|node:path|node:os|node:buffer|fs\/promises|from "fs"|from "path"|from "os"|from "module")/,
    );
  });
});
