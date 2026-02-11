import { readFileSync, writeFileSync, mkdirSync, readdirSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { convertPptxToPng } from "../../src/converter.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURE_DIR = join(__dirname, "fixtures");
const SNAPSHOT_DIR = join(__dirname, "snapshots");

async function main(): Promise<void> {
  mkdirSync(SNAPSHOT_DIR, { recursive: true });

  const fixtures = readdirSync(FIXTURE_DIR).filter(
    (f) => f.startsWith("vrt-") && f.endsWith(".pptx"),
  );

  if (fixtures.length === 0) {
    console.error('No VRT fixtures found. Run "npm run test:vrt:fixtures" first.');
    process.exit(1);
  }

  let totalSlides = 0;

  for (const fixture of fixtures.sort()) {
    const name = fixture.replace(".pptx", "").replace("vrt-", "");
    const fixturePath = join(FIXTURE_DIR, fixture);
    const input = readFileSync(fixturePath);

    console.log(`Processing: ${fixture}`);
    const results = await convertPptxToPng(input);

    for (const result of results) {
      const snapshotPath = join(SNAPSHOT_DIR, `${name}-slide${result.slideNumber}.png`);
      writeFileSync(snapshotPath, result.png);
      console.log(`  Updated: ${snapshotPath} (${result.width}x${result.height})`);
      totalSlides++;
    }
  }

  console.log(`\nDone! Updated ${totalSlides} snapshot(s).`);
}

main().catch(console.error);
