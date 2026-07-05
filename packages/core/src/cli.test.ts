import { mkdtemp, readdir, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { runCli, type RunCliOptions } from "./cli-runner.js";
import type { ConvertOptions, PngConversionReport, SvgConversionReport } from "./converter.js";

describe("pptx-glimpse CLI", () => {
  it("writes SVG files for converted slides", async () => {
    const workspace = await createWorkspace();
    const pptxPath = await writeInput(workspace);
    const streams = createStreams();
    const calls: ConvertOptions[] = [];

    const exitCode = await runCli(["convert", pptxPath], {
      cwd: workspace,
      streams: streams.streams,
      converters: {
        convertPptxToSvg: (_input, options) => {
          calls.push(options ?? {});
          return Promise.resolve(
            svgReport([
              [1, "<svg>one</svg>"],
              [2, "<svg>two</svg>"],
            ]),
          );
        },
        convertPptxToPng: failPngConverter,
      },
    });

    expect(exitCode).toBe(0);
    expect(calls).toEqual([{ logLevel: "off", skipSystemFonts: true }]);
    expect(await sortedFiles(workspace)).toEqual([
      "deck-slide1.svg",
      "deck-slide2.svg",
      "deck.pptx",
    ]);
    await expect(readFile(join(workspace, "deck-slide1.svg"), "utf8")).resolves.toBe(
      "<svg>one</svg>",
    );
    expect(streams.stdout).toContain("deck-slide1.svg");
    expect(streams.stderr).toBe("");
  });

  it("switches to PNG output with --format png", async () => {
    const workspace = await createWorkspace();
    const pptxPath = await writeInput(workspace);

    const exitCode = await runCli(["convert", pptxPath, "--format", "png"], {
      cwd: workspace,
      streams: createStreams().streams,
      converters: {
        convertPptxToSvg: failSvgConverter,
        convertPptxToPng: () =>
          Promise.resolve(pngReport([[1, new Uint8Array([0x89, 0x50, 0x4e, 0x47])]])),
      },
    });

    expect(exitCode).toBe(0);
    expect(await readFile(join(workspace, "deck-slide1.png"))).toEqual(
      Buffer.from([0x89, 0x50, 0x4e, 0x47]),
    );
  });

  it("supports --slides and creates nested --out directories", async () => {
    const workspace = await createWorkspace();
    const pptxPath = await writeInput(workspace);
    const outDir = join(workspace, "nested", "out");
    const streams = createStreams();
    const calls: ConvertOptions[] = [];

    const exitCode = await runCli(
      ["convert", pptxPath, "--slides", "1,3", "--out", outDir, "--log-level", "debug"],
      {
        cwd: workspace,
        streams: streams.streams,
        converters: {
          convertPptxToSvg: (_input, options) => {
            calls.push(options ?? {});
            return Promise.resolve({
              ...svgReport([[1, "<svg />"]]),
              diagnostics: [
                {
                  source: "renderer",
                  severity: "warning",
                  code: "renderer.test",
                  message: "Test warning",
                  slideNumber: 1,
                },
              ],
            });
          },
          convertPptxToPng: failPngConverter,
        },
      },
    );

    expect(exitCode).toBe(0);
    expect(calls).toEqual([{ logLevel: "off", skipSystemFonts: true, slides: [1, 3] }]);
    expect(await sortedFiles(outDir)).toEqual(["deck-slide1.svg"]);
    expect(streams.stderr).toContain("pptx-glimpse: 1 warning(s)");
    expect(streams.stderr).toContain("pptx-glimpse: slide 1 renderer.test: Test warning");
  });

  it("prints help", async () => {
    const streams = createStreams();

    const exitCode = await runCli(["convert", "--help"], { streams: streams.streams });

    expect(exitCode).toBe(0);
    expect(streams.stdout).toContain("pptx-glimpse convert <file.pptx>");
    expect(streams.stderr).toBe("");
  });

  it("returns a non-zero exit code for invalid input", async () => {
    const workspace = await createWorkspace();
    const streams = createStreams();

    const exitCode = await runCli(["convert", "missing.pptx"], {
      cwd: workspace,
      streams: streams.streams,
      converters: {
        convertPptxToSvg: failSvgConverter,
        convertPptxToPng: failPngConverter,
      },
    });

    expect(exitCode).toBe(1);
    expect(streams.stderr).toContain("Input file not found");
  });

  it("returns a non-zero exit code for invalid arguments", async () => {
    const workspace = await createWorkspace();
    const pptxPath = await writeInput(workspace);
    const streams = createStreams();

    const exitCode = await runCli(["convert", pptxPath, "--slides", "0,a"], {
      cwd: workspace,
      streams: streams.streams,
    });

    expect(exitCode).toBe(1);
    expect(streams.stderr).toContain("Invalid --slides value");
  });

  it("registers the public package bin", async () => {
    const packageJson = parsePackageJson(
      await readFile(new URL("../package.json", import.meta.url), "utf8"),
    );

    expect(packageJson.bin).toEqual({ "pptx-glimpse": "./dist/cli.js" });
  });
});

async function createWorkspace(): Promise<string> {
  const workspace = await mkdtemp(join(tmpdir(), "pptx-glimpse-cli-"));
  return workspace;
}

async function writeInput(workspace: string): Promise<string> {
  const pptxPath = join(workspace, "deck.pptx");
  await writeFile(pptxPath, "pptx bytes");
  return pptxPath;
}

async function sortedFiles(dir: string): Promise<string[]> {
  return (await readdir(dir)).sort();
}

function createStreams(): {
  readonly streams: RunCliOptions["streams"];
  readonly stdout: string;
  readonly stderr: string;
} {
  let stdout = "";
  let stderr = "";
  return {
    streams: {
      stdout: {
        write(chunk: string) {
          stdout += chunk;
        },
      },
      stderr: {
        write(chunk: string) {
          stderr += chunk;
        },
      },
    },
    get stdout() {
      return stdout;
    },
    get stderr() {
      return stderr;
    },
  };
}

function svgReport(slides: Array<[number, string]>): SvgConversionReport {
  return {
    slides: slides.map(([slideNumber, svg]) => ({ slideNumber, svg })),
    diagnostics: [],
    supportCoverage: emptySupportCoverage(),
  };
}

function pngReport(slides: Array<[number, Uint8Array]>): PngConversionReport {
  return {
    slides: slides.map(([slideNumber, png]) => ({ slideNumber, png, width: 1, height: 1 })),
    diagnostics: [],
    supportCoverage: emptySupportCoverage(),
  };
}

function emptySupportCoverage(): SvgConversionReport["supportCoverage"] {
  return {
    overall: {
      inputElements: 0,
      outputElements: 0,
      skippedElements: 0,
      unresolvedElements: 0,
      fallbackElements: 0,
      warnings: 0,
    },
    slides: [],
  };
}

function parsePackageJson(json: string): { bin?: Record<string, string> } {
  const parsed: unknown = JSON.parse(json);
  if (typeof parsed !== "object" || parsed === null) {
    throw new Error("Expected package.json object");
  }
  const bin = "bin" in parsed ? parsed.bin : undefined;
  if (bin !== undefined && !isStringRecord(bin)) {
    throw new Error("Expected package.json bin record");
  }
  return bin === undefined ? {} : { bin };
}

function isStringRecord(value: unknown): value is Record<string, string> {
  if (typeof value !== "object" || value === null) return false;
  return Object.values(value).every((item) => typeof item === "string");
}

function failSvgConverter(): Promise<SvgConversionReport> {
  return Promise.reject(new Error("Unexpected SVG conversion"));
}

function failPngConverter(): Promise<PngConversionReport> {
  return Promise.reject(new Error("Unexpected PNG conversion"));
}
