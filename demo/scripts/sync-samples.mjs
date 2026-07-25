import { copyFile, mkdir, rm } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const demoRoot = resolve(here, "..");
const sampleNames = ["editor-demo.pptx"];
const obsoleteSampleNames = ["real-basic-theme.pptx", "real-product-page.pptx"];
const assetDir = resolve(demoRoot, "assets");
const sampleDir = resolve(demoRoot, "public/samples");

await mkdir(sampleDir, { recursive: true });

await Promise.all(obsoleteSampleNames.map((name) => rm(resolve(sampleDir, name), { force: true })));

await Promise.all(
  sampleNames.map((name) => copyFile(resolve(assetDir, name), resolve(sampleDir, name))),
);

console.log(`Synced ${sampleNames.length.toString()} demo PPTX sample.`);
