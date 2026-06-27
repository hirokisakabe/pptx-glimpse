import { readdir, readFile } from "node:fs/promises";
import { dirname, extname, isAbsolute, join, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { documentExperimentalApi } from "./experimental.js";

interface PackageJson {
  readonly dependencies?: Record<string, string>;
  readonly devDependencies?: Record<string, string>;
  readonly exports?: Record<string, PackageJsonExport>;
  readonly peerDependencies?: Record<string, string>;
}

interface PackageJsonExport {
  readonly import?: string;
  readonly require?: string;
  readonly types?: string;
}

const FORBIDDEN_DEPENDENCIES = new Set([
  "@hirokisakabe/pom",
  "@pptx-glimpse/core",
  "@pptx-glimpse/editor-core",
  "@pptx-glimpse/renderer",
  "pptx-glimpse",
  "pptx-glimpse-renderer",
]);

const SOURCE_ROOT = fileURLToPath(new URL(".", import.meta.url));
const PACKAGE_ROOT = fileURLToPath(new URL("..", import.meta.url));
const PACKAGE_JSON = new URL("../package.json", import.meta.url);

const IMPORT_PATTERN =
  /\b(?:import\s*\(\s*["']([^"']+)["']|(?:import|export)\s+(?:type\s+)?(?:[^'"]*?\s+from\s+)?["']([^"']+)["'])/g;

function isForbiddenDependency(specifier: string): boolean {
  for (const dependency of FORBIDDEN_DEPENDENCIES) {
    if (specifier === dependency || specifier.startsWith(`${dependency}/`)) {
      return true;
    }
  }

  return false;
}

function isExternalRelativeImport(sourceFile: string, specifier: string): boolean {
  if (!specifier.startsWith(".")) {
    return false;
  }

  const resolvedImportPath = resolve(dirname(sourceFile), specifier);
  const relativeImportPath = relative(PACKAGE_ROOT, resolvedImportPath);

  return (
    relativeImportPath.startsWith(`..${sep}`) ||
    relativeImportPath === ".." ||
    isAbsolute(relativeImportPath)
  );
}

async function listTypeScriptFiles(dir: string): Promise<string[]> {
  const entries = await readdir(dir, { withFileTypes: true });
  const files = await Promise.all(
    entries.map(async (entry) => {
      const path = join(dir, entry.name);

      if (entry.isDirectory()) {
        return listTypeScriptFiles(path);
      }

      return extname(entry.name) === ".ts" && !entry.name.endsWith(".test.ts") ? [path] : [];
    }),
  );

  return files.flat();
}

function parsePackageJson(text: string): PackageJson {
  const parsed: unknown = JSON.parse(text);

  return parsed as PackageJson;
}

describe("@pptx-glimpse/document package boundary", () => {
  it("exposes the experimental entry point", async () => {
    const packageJson = parsePackageJson(await readFile(PACKAGE_JSON, "utf8"));

    expect(documentExperimentalApi).toEqual({
      packageName: "@pptx-glimpse/document",
      status: "experimental",
    });
    expect(packageJson.exports?.["./experimental"]).toEqual({
      types: "./src/experimental.ts",
      default: "./src/experimental.ts",
    });
  });

  it("does not declare dependencies on higher-level packages", async () => {
    const packageJson = parsePackageJson(await readFile(PACKAGE_JSON, "utf8"));
    const declaredDependencies = {
      ...packageJson.dependencies,
      ...packageJson.devDependencies,
      ...packageJson.peerDependencies,
    };

    expect(Object.keys(declaredDependencies).filter(isForbiddenDependency)).toEqual([]);
  });

  it("does not import higher-level packages from source", async () => {
    const sourceFiles = await listTypeScriptFiles(SOURCE_ROOT);
    const violations: string[] = [];

    await Promise.all(
      sourceFiles.map(async (sourceFile) => {
        const source = await readFile(sourceFile, "utf8");

        for (const match of source.matchAll(IMPORT_PATTERN)) {
          const specifier = match[1] ?? match[2];

          if (
            specifier !== undefined &&
            (isForbiddenDependency(specifier) || isExternalRelativeImport(sourceFile, specifier))
          ) {
            violations.push(`${relative(SOURCE_ROOT, sourceFile)} imports ${specifier}`);
          }
        }
      }),
    );

    expect(violations).toEqual([]);
  });
});
