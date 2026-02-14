import { readFileSync, writeFileSync, mkdirSync } from "fs";
import { basename, join, resolve } from "path";
import { convertPptxToSvg, convertPptxToPng } from "../src/converter.js";

async function main() {
  const filePath = process.argv[2];

  if (!filePath) {
    console.error("Usage: npx tsx scripts/test-render.ts <pptx-file> [output-dir]");
    console.error("");
    console.error("Examples:");
    console.error("  npx tsx scripts/test-render.ts presentation.pptx");
    console.error("  npx tsx scripts/test-render.ts presentation.pptx ./output");
    process.exit(1);
  }

  const outputDir = resolve(process.argv[3] ?? "./output");
  mkdirSync(outputDir, { recursive: true });

  const input = readFileSync(filePath);
  const name = basename(filePath, ".pptx");

  console.log(`Converting: ${filePath}`);
  console.log(`Output dir: ${outputDir}`);
  console.log("");

  const svgResults = await convertPptxToSvg(input, { logLevel: "warn" });
  const pngResults = await convertPptxToPng(input);

  for (const svg of svgResults) {
    const svgPath = join(outputDir, `${name}-slide${svg.slideNumber}.svg`);
    writeFileSync(svgPath, svg.svg);
    console.log(`  SVG: ${svgPath}`);
  }

  for (const png of pngResults) {
    const pngPath = join(outputDir, `${name}-slide${png.slideNumber}.png`);
    writeFileSync(pngPath, png.png);
    console.log(`  PNG: ${pngPath} (${png.width}x${png.height})`);
  }

  console.log("");
  console.log(`Done! ${svgResults.length} slide(s) converted.`);
}

main().catch(console.error);
