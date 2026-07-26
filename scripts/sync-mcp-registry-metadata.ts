import { readFile, writeFile } from "node:fs/promises";

const packageJsonPath = new URL("../packages/mcp/package.json", import.meta.url);
const serverJsonPath = new URL("../packages/mcp/server.json", import.meta.url);

const packageJson: unknown = JSON.parse(await readFile(packageJsonPath, "utf8"));
const serverJson: unknown = JSON.parse(await readFile(serverJsonPath, "utf8"));

if (!isRecord(packageJson) || typeof packageJson.version !== "string") {
  throw new Error("packages/mcp/package.json must contain a string version");
}
if (!isRecord(serverJson) || !Array.isArray(serverJson.packages)) {
  throw new Error("packages/mcp/server.json must contain a packages array");
}

const npmPackage = serverJson.packages[0];
if (!isRecord(npmPackage)) {
  throw new Error("packages/mcp/server.json must contain npm package metadata");
}

serverJson.version = packageJson.version;
npmPackage.version = packageJson.version;
await writeFile(serverJsonPath, `${JSON.stringify(serverJson, null, 2)}\n`);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
