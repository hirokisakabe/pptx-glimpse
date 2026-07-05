import { copyFile, mkdir } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const demoRoot = resolve(here, "..");
const repoRoot = resolve(demoRoot, "..");
const sampleNames = ["real-basic-theme.pptx", "real-product-page.pptx"];
const sampleDir = resolve(demoRoot, "public/samples");

await mkdir(sampleDir, { recursive: true });

await Promise.all(
  sampleNames.map((name) =>
    copyFile(resolve(repoRoot, "shared-fixtures", name), resolve(sampleDir, name)),
  ),
);

console.log(`Synced ${sampleNames.length.toString()} sample PPTX files from shared-fixtures.`);
