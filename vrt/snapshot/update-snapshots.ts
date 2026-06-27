import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { convertPptxToPng } from "../../packages/core/src/converter.js";
import { getVrtRenderOptions } from "./render-options.js";
import { SHARED_FIXTURE_CASES } from "./vrt-cases.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURE_DIR = join(__dirname, "fixtures");
const SHARED_FIXTURE_DIR = join(__dirname, "..", "..", "shared-fixtures");
const SNAPSHOT_DIR = join(__dirname, "snapshots");

const CONCURRENCY = 4;

async function processFixture(fixturePath: string, name: string): Promise<number> {
  const input = readFileSync(fixturePath);
  console.log(`Processing: ${name}`);
  const results = await convertPptxToPng(input, await getVrtRenderOptions());
  let count = 0;

  for (const result of results) {
    const snapshotPath = join(SNAPSHOT_DIR, `${name}-slide${result.slideNumber}.png`);
    writeFileSync(snapshotPath, result.png);
    console.log(`  Updated: ${snapshotPath} (${result.width}x${result.height})`);
    count++;
  }

  return count;
}

async function runWithConcurrency<T>(
  tasks: (() => Promise<T>)[],
  concurrency: number,
): Promise<T[]> {
  const results: T[] = new Array(tasks.length);
  let nextIndex = 0;

  async function worker(): Promise<void> {
    while (nextIndex < tasks.length) {
      const index = nextIndex++;
      results[index] = await tasks[index]();
    }
  }

  await Promise.all(Array.from({ length: Math.min(concurrency, tasks.length) }, () => worker()));
  return results;
}

async function main(): Promise<void> {
  mkdirSync(SNAPSHOT_DIR, { recursive: true });

  const tasks: (() => Promise<number>)[] = [];

  // Generated fixtures
  const fixtures = existsSync(FIXTURE_DIR)
    ? readdirSync(FIXTURE_DIR).filter((f) => f.endsWith(".pptx"))
    : [];

  for (const fixture of fixtures.sort()) {
    const name = fixture.replace(".pptx", "");
    const fixturePath = join(FIXTURE_DIR, fixture);
    tasks.push(() => processFixture(fixturePath, name));
  }

  // Shared fixtures
  for (const { name, fixture } of SHARED_FIXTURE_CASES) {
    const fixturePath = join(SHARED_FIXTURE_DIR, fixture);
    if (!existsSync(fixturePath)) {
      console.warn(`  Skipped (not found): ${fixturePath}`);
      continue;
    }
    tasks.push(() => processFixture(fixturePath, name));
  }

  if (tasks.length === 0) {
    console.error("No fixtures found. Run 'npm run vrt:snapshot:fixtures' first.");
    process.exit(1);
  }

  const counts = await runWithConcurrency(tasks, CONCURRENCY);
  const totalSlides = counts.reduce((sum, c) => sum + c, 0);

  console.log(`\nDone! Updated ${totalSlides} snapshot(s).`);
}

main().catch(console.error);
