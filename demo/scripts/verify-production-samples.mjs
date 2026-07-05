import { execFile, spawn } from "node:child_process";
import { readFile } from "node:fs/promises";
import { createServer } from "node:net";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { chromium, expect } from "@playwright/test";

const execFileAsync = promisify(execFile);
const here = dirname(fileURLToPath(import.meta.url));
const demoRoot = resolve(here, "..");
const repoRoot = resolve(demoRoot, "..");
const sampleNames = ["real-basic-theme.pptx", "real-product-page.pptx"];

const port = await getFreePort();
const baseUrl = `http://127.0.0.1:${port.toString()}`;
let server;
let browser;

try {
  await execFileAsync("npm", ["run", "build"], {
    cwd: demoRoot,
    maxBuffer: 20 * 1024 * 1024,
  });

  server = spawn("npm", ["run", "start", "--", "--hostname", "127.0.0.1", "--port", String(port)], {
    cwd: demoRoot,
    stdio: ["ignore", "pipe", "pipe"],
  });

  let serverOutput = "";
  server.stdout.on("data", (chunk) => {
    serverOutput += chunk.toString();
  });
  server.stderr.on("data", (chunk) => {
    serverOutput += chunk.toString();
  });

  await waitForHttpOk(baseUrl, () => serverOutput);

  for (const sampleName of sampleNames) {
    await assertSampleServed(baseUrl, sampleName);
  }

  browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto(baseUrl);
  await page.getByTestId("sample-basic-theme").click();
  await expect(page.getByTestId("viewer-status")).toContainText("slides rendered", {
    timeout: 30_000,
  });
  await expect(page.locator("svg").first()).toBeVisible();
} finally {
  await browser?.close();
  if (server !== undefined) {
    server.kill("SIGTERM");
    await new Promise((resolve) => {
      server.once("exit", resolve);
      setTimeout(resolve, 5_000);
    });
  }
}

async function assertSampleServed(baseUrl, sampleName) {
  const response = await fetch(`${baseUrl}/samples/${sampleName}`);
  if (!response.ok) {
    throw new Error(`/samples/${sampleName} returned ${response.status.toString()}`);
  }

  const actual = Buffer.from(await response.arrayBuffer());
  const expected = await readFile(resolve(repoRoot, "shared-fixtures", sampleName));

  if (!actual.equals(expected)) {
    throw new Error(`/samples/${sampleName} did not match shared-fixtures/${sampleName}`);
  }
}

async function waitForHttpOk(url, getOutput) {
  const deadline = Date.now() + 60_000;
  let lastError;

  while (Date.now() < deadline) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
      lastError = new Error(`HTTP ${response.status.toString()}`);
    } catch (error) {
      lastError = error;
    }
    await new Promise((resolve) => setTimeout(resolve, 500));
  }

  throw new Error(`next start did not become ready: ${String(lastError)}\n${getOutput()}`);
}

async function getFreePort() {
  const server = createServer();
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const address = server.address();
  await new Promise((resolve, reject) => {
    server.close((error) => {
      if (error) reject(error);
      else resolve();
    });
  });

  if (address === null || typeof address === "string") {
    throw new Error("failed to allocate a local port");
  }

  return address.port;
}
