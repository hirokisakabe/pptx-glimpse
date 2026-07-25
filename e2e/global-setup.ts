import { execFile } from "node:child_process";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");

export default async function globalSetup(): Promise<void> {
  for (const packagePath of [
    "packages/document",
    "packages/editor",
    "packages/renderer",
    "packages/core",
  ]) {
    await execFileAsync(
      resolve(repoRoot, "node_modules/.bin/tsup"),
      ["--config", "tsup.config.ts"],
      {
        cwd: resolve(repoRoot, packagePath),
        maxBuffer: 10 * 1024 * 1024,
      },
    );
  }
}
