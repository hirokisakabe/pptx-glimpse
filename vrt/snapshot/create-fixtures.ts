/**
 * PPTX fixture generation script for VRT (Visual Regression Testing)
 *
 * Usage: npx tsx vrt/snapshot/create-fixtures.ts [case-name...]
 */
import { pathToFileURL } from "url";

import type { FixtureCreatorMap } from "./fixture-builder.js";
import { backgroundFixtureCreators } from "./fixtures-src/backgrounds.js";
import { chartFixtureCreators } from "./fixtures-src/charts.js";
import { imageFixtureCreators } from "./fixtures-src/images.js";
import { miscFixtureCreators } from "./fixtures-src/misc.js";
import { placeholderFixtureCreators } from "./fixtures-src/placeholder.js";
import { shapeFixtureCreators } from "./fixtures-src/shapes.js";
import { tableFixtureCreators } from "./fixtures-src/tables.js";
import { textFixtureCreators } from "./fixtures-src/text.js";
import { resolveGeneratedVrtCases } from "./vrt-cases.js";

export {
  buildPptx,
  shapeXml,
  slideRelsXml,
  textBodyXmlHelper,
  wrapSlideXml,
} from "./fixture-builder.js";

const FIXTURE_CREATORS: FixtureCreatorMap = {
  ...shapeFixtureCreators,
  ...backgroundFixtureCreators,
  ...chartFixtureCreators,
  ...imageFixtureCreators,
  ...tableFixtureCreators,
  ...textFixtureCreators,
  ...placeholderFixtureCreators,
  ...miscFixtureCreators,
};

async function main(): Promise<void> {
  const caseNames = process.argv.slice(2);
  const selectedCases = resolveGeneratedVrtCases(caseNames);

  console.log("Creating VRT fixtures...\n");

  for (const { fixture } of selectedCases) {
    const creator = FIXTURE_CREATORS[fixture];
    if (!creator) {
      throw new Error(
        'No fixture creator found for "' +
          fixture +
          '". Add a creator function to the appropriate fixtures-src module and register it in FIXTURE_CREATORS.',
      );
    }
    await creator();
  }

  if (caseNames.length > 0 && selectedCases.length === 0) {
    console.log("No generated VRT fixtures selected.");
  }

  console.log("\nDone!");
}

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error: unknown) => {
    console.error(error);
    process.exit(1);
  });
}
