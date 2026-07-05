import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import { basename, extname, resolve } from "node:path";
import { parseArgs } from "node:util";

import {
  type ConvertOptions,
  convertPptxToPng,
  convertPptxToSvg,
  type PngConversionReport,
  type SvgConversionReport,
} from "./converter.js";

type CliFormat = "svg" | "png";
type CliLogLevel = "off" | "warn" | "debug";

interface CliWritable {
  write(chunk: string): unknown;
}

interface CliStreams {
  readonly stdout: CliWritable;
  readonly stderr: CliWritable;
}

interface CliConverters {
  readonly convertPptxToSvg: (
    input: Uint8Array,
    options?: ConvertOptions,
  ) => Promise<SvgConversionReport>;
  readonly convertPptxToPng: (
    input: Uint8Array,
    options?: ConvertOptions,
  ) => Promise<PngConversionReport>;
}

export interface RunCliOptions {
  readonly cwd?: string;
  readonly streams?: CliStreams;
  readonly converters?: CliConverters;
}

interface ConvertCommandOptions {
  readonly inputPath: string;
  readonly outputDir: string;
  readonly format: CliFormat;
  readonly slides?: number[];
  readonly logLevel: CliLogLevel;
  readonly systemFonts: boolean;
}

const defaultConverters: CliConverters = {
  convertPptxToSvg,
  convertPptxToPng,
};

const helpText = `Usage:
  pptx-glimpse convert <file.pptx> [options]

Options:
  --format <svg|png>       Output format. Defaults to svg.
  --slides <list>          Comma-separated 1-based slide numbers, such as 1,3.
  --out <dir>              Output directory. Defaults to the current directory.
  --log-level <level>      Diagnostic output: off, warn, or debug. Defaults to warn.
  --system-fonts           Scan OS system font directories for better text fidelity.
  -h, --help               Show this help message.
`;

export async function runCli(
  argv: readonly string[],
  options: RunCliOptions = {},
): Promise<number> {
  const cwd = options.cwd ?? process.cwd();
  const streams = options.streams ?? { stdout: process.stdout, stderr: process.stderr };
  const converters = options.converters ?? defaultConverters;

  try {
    if (argv.length === 0 || argv[0] === "--help" || argv[0] === "-h" || argv[0] === "help") {
      streams.stdout.write(helpText);
      return 0;
    }

    const [command, ...commandArgs] = argv;
    if (command !== "convert") {
      throw new CliUsageError(`Unknown command: ${command}`);
    }

    const commandOptions = parseConvertCommand(commandArgs, cwd);
    await runConvertCommand(commandOptions, streams, converters);
    return 0;
  } catch (error) {
    if (error instanceof CliHelpRequested) {
      streams.stdout.write(helpText);
      return 0;
    }
    streams.stderr.write(`pptx-glimpse: ${formatError(error)}\n`);
    return 1;
  }
}

function parseConvertCommand(args: readonly string[], cwd: string): ConvertCommandOptions {
  let parsed: ReturnType<typeof parseArgs>;
  try {
    parsed = parseArgs({
      args: [...args],
      allowPositionals: true,
      options: {
        format: { type: "string" },
        slides: { type: "string" },
        out: { type: "string" },
        "log-level": { type: "string" },
        "system-fonts": { type: "boolean" },
        help: { type: "boolean", short: "h" },
      },
    });
  } catch (error) {
    throw new CliUsageError(formatError(error));
  }

  if (parsed.values.help === true) {
    throw new CliHelpRequested();
  }

  if (parsed.positionals.length !== 1) {
    throw new CliUsageError("Expected exactly one PPTX input file.");
  }

  const format = parseFormat(stringOption(parsed.values.format, "--format"));
  const logLevel = parseLogLevel(stringOption(parsed.values["log-level"], "--log-level"));
  const slides = parseSlides(stringOption(parsed.values.slides, "--slides"));
  const outputDir = resolve(cwd, stringOption(parsed.values.out, "--out") ?? ".");

  return {
    inputPath: resolve(cwd, parsed.positionals[0]),
    outputDir,
    format,
    ...(slides !== undefined ? { slides } : {}),
    logLevel,
    systemFonts: parsed.values["system-fonts"] === true,
  };
}

async function runConvertCommand(
  options: ConvertCommandOptions,
  streams: CliStreams,
  converters: CliConverters,
): Promise<void> {
  await assertReadableFile(options.inputPath);
  await mkdir(options.outputDir, { recursive: true });

  const input = await readFile(options.inputPath);
  const basenameWithoutExtension = basename(options.inputPath, extname(options.inputPath));
  const conversionOptions: ConvertOptions = {
    // Route diagnostics through the CLI streams below instead of converter console output.
    logLevel: "off",
    skipSystemFonts: !options.systemFonts,
    ...(options.slides !== undefined ? { slides: options.slides } : {}),
  };

  if (options.format === "png") {
    const report = await converters.convertPptxToPng(input, conversionOptions);
    for (const slide of report.slides) {
      const outputPath = resolve(
        options.outputDir,
        `${basenameWithoutExtension}-slide${slide.slideNumber}.png`,
      );
      await writeFile(outputPath, slide.png);
      streams.stdout.write(`${outputPath}\n`);
    }
    printDiagnostics(report, options.logLevel, streams.stderr);
    return;
  }

  const report = await converters.convertPptxToSvg(input, conversionOptions);
  for (const slide of report.slides) {
    const outputPath = resolve(
      options.outputDir,
      `${basenameWithoutExtension}-slide${slide.slideNumber}.svg`,
    );
    await writeFile(outputPath, slide.svg, "utf8");
    streams.stdout.write(`${outputPath}\n`);
  }
  printDiagnostics(report, options.logLevel, streams.stderr);
}

async function assertReadableFile(inputPath: string): Promise<void> {
  try {
    const inputStat = await stat(inputPath);
    if (!inputStat.isFile()) {
      throw new CliUsageError(`Input is not a file: ${inputPath}`);
    }
  } catch (error) {
    if (error instanceof CliUsageError) throw error;
    throw new CliUsageError(`Input file not found: ${inputPath}`);
  }
}

function parseFormat(value: string | undefined): CliFormat {
  if (value === undefined) return "svg";
  if (value === "svg" || value === "png") return value;
  throw new CliUsageError(`Invalid --format value: ${value}. Expected svg or png.`);
}

function stringOption(
  value: string | boolean | Array<string | boolean> | undefined,
  name: string,
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value === "string") return value;
  throw new CliUsageError(`Invalid ${name} value.`);
}

function parseLogLevel(value: string | undefined): CliLogLevel {
  if (value === undefined) return "warn";
  if (value === "off" || value === "warn" || value === "debug") return value;
  throw new CliUsageError(`Invalid --log-level value: ${value}. Expected off, warn, or debug.`);
}

function parseSlides(value: string | undefined): number[] | undefined {
  if (value === undefined) return undefined;
  const slides = value.split(",").map((part) => {
    const trimmed = part.trim();
    if (!/^[1-9]\d*$/.test(trimmed)) {
      throw new CliUsageError(`Invalid --slides value: ${value}. Expected values like 1,3.`);
    }
    return Number(trimmed);
  });
  if (slides.length === 0) {
    throw new CliUsageError("Expected at least one slide number.");
  }
  return slides;
}

function printDiagnostics(
  report: SvgConversionReport | PngConversionReport,
  logLevel: CliLogLevel,
  stderr: CliWritable,
): void {
  if (logLevel === "off") return;

  const warnings = report.diagnostics.filter((diagnostic) => diagnostic.severity === "warning");
  if (warnings.length === 0) return;

  stderr.write(`pptx-glimpse: ${warnings.length} warning(s)\n`);
  if (logLevel !== "debug") return;

  for (const warning of warnings) {
    const slide = warning.slideNumber !== undefined ? ` slide ${warning.slideNumber}` : "";
    stderr.write(`pptx-glimpse:${slide} ${warning.code}: ${warning.message}\n`);
  }
}

function formatError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

class CliUsageError extends Error {}

class CliHelpRequested extends Error {
  constructor() {
    super(helpText);
  }
}
