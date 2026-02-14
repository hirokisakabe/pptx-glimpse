import { readFileSync, writeFileSync, mkdirSync, readdirSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { convertPptxToPng } from "../../src/converter.js";
import { loadSystemFontBuffers, VRT_FONT_MAPPING } from "./load-fonts.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURE_DIR = join(__dirname, "fixtures");
const SNAPSHOT_DIR = join(__dirname, "snapshots");

async function main(): Promise<void> {
  mkdirSync(SNAPSHOT_DIR, { recursive: true });

  const fixtures = readdirSync(FIXTURE_DIR).filter((f) => f.endsWith(".pptx"));

  if (fixtures.length === 0) {
    console.error('No VRT fixtures found. Run "npm run vrt:snapshot:fixtures" first.');
    process.exit(1);
  }

  const fontBuffers = loadSystemFontBuffers();
  console.log(`Loaded ${fontBuffers.length} font(s) for text-to-path conversion`);

  let totalSlides = 0;

  for (const fixture of fixtures.sort()) {
    const name = fixture.replace(".pptx", "");
    const fixturePath = join(FIXTURE_DIR, fixture);
    const input = readFileSync(fixturePath);

    console.log(`Processing: ${fixture}`);
    const results = await convertPptxToPng(input, {
      fonts: { fontBuffers },
      fontMapping: VRT_FONT_MAPPING,
    });

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
