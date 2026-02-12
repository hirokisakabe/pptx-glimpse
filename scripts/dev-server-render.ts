import { readFileSync } from "fs";
import { convertPptxToSvg } from "../src/converter.js";

async function main(): Promise<void> {
  const pptxPath = process.argv[2];
  if (!pptxPath) {
    process.exit(1);
  }

  const input = readFileSync(pptxPath);
  const slides = await convertPptxToSvg(input, { logLevel: "warn" });

  process.stdout.write(JSON.stringify(slides));
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
