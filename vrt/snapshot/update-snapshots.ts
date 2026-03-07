import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { convertPptxToPng } from "../../src/converter.js";
import { SHARED_FIXTURE_CASES } from "./vrt-cases.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURE_DIR = join(__dirname, "fixtures");
const SHARED_FIXTURE_DIR = join(__dirname, "..", "..", "shared-fixtures");
const SNAPSHOT_DIR = join(__dirname, "snapshots");

async function processFixture(fixturePath: string, name: string): Promise<number> {
  const input = readFileSync(fixturePath);
  console.log(`Processing: ${name}`);
  const results = await convertPptxToPng(input);
  let count = 0;

  for (const result of results) {
    const snapshotPath = join(SNAPSHOT_DIR, `${name}-slide${result.slideNumber}.png`);
    writeFileSync(snapshotPath, result.png);
    console.log(`  Updated: ${snapshotPath} (${result.width}x${result.height})`);
    count++;
  }

  return count;
}

async function main(): Promise<void> {
  mkdirSync(SNAPSHOT_DIR, { recursive: true });

  let totalSlides = 0;

  // Generated fixtures
  const fixtures = existsSync(FIXTURE_DIR)
    ? readdirSync(FIXTURE_DIR).filter((f) => f.endsWith(".pptx"))
    : [];

  for (const fixture of fixtures.sort()) {
    const name = fixture.replace(".pptx", "");
    totalSlides += await processFixture(join(FIXTURE_DIR, fixture), name);
  }

  // Shared fixtures
  for (const { name, fixture } of SHARED_FIXTURE_CASES) {
    const fixturePath = join(SHARED_FIXTURE_DIR, fixture);
    if (!existsSync(fixturePath)) {
      console.warn(`  Skipped (not found): ${fixturePath}`);
      continue;
    }
    totalSlides += await processFixture(fixturePath, name);
  }

  if (totalSlides === 0) {
    console.error("No fixtures found. Run 'npm run vrt:snapshot:fixtures' first.");
    process.exit(1);
  }

  console.log(`\nDone! Updated ${totalSlides} snapshot(s).`);
}

main().catch(console.error);
