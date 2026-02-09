import { readFileSync, writeFileSync } from "fs";
import { convertPptxToSvg, convertPptxToPng } from "../src/converter.js";

async function main() {
  const input = readFileSync("tests/fixtures/basic-shapes.pptx");

  const svgResults = await convertPptxToSvg(input);
  writeFileSync("/tmp/test-slide.svg", svgResults[0].svg);
  console.log("SVG written to /tmp/test-slide.svg");

  const pngResults = await convertPptxToPng(input);
  writeFileSync("/tmp/test-slide.png", pngResults[0].png);
  console.log(`PNG written to /tmp/test-slide.png (${pngResults[0].width}x${pngResults[0].height})`);
}

main().catch(console.error);
