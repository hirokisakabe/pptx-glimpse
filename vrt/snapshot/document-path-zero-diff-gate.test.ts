import { existsSync, readFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { convertPptxToPngViaDocumentPath } from "../../packages/core/src/experimental-document-renderer.js";
import { convertPptxToPngViaParserPath } from "../../packages/core/src/parser-path-oracle.js";
import { compareImageBuffers } from "../compare-utils.js";
import {
  DOCUMENT_PATH_VRT_CASES,
  DOCUMENT_PATH_VRT_RENDER_WIDTH,
  type DocumentPathVrtFixtureGroup,
} from "./document-path-cases.js";
import { DOCUMENT_PATH_VRT_RENDER_OPTIONS } from "./render-options.js";
import { SHARED_FIXTURE_CASES, VRT_CASES } from "./vrt-cases.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SHARED_FIXTURE_DIR = join(__dirname, "..", "..", "shared-fixtures");
const GENERATED_FIXTURE_DIR = join(__dirname, "fixtures");
const DIFF_DIR = join(__dirname, "diffs");

const PIXEL_THRESHOLD = 0;
const MISMATCH_TOLERANCE = 0;

interface ZeroDiffCase {
  readonly group: DocumentPathVrtFixtureGroup;
  readonly name: string;
  readonly fixture: string;
}

const ZERO_DIFF_CASES = [
  ...SHARED_FIXTURE_CASES.map((testCase) => ({ ...testCase, group: "shared" as const })),
  ...VRT_CASES.map((testCase) => ({ ...testCase, group: "generated" as const })),
] as const satisfies readonly ZeroDiffCase[];

describe("Document path zero-diff default switch gate", { timeout: 60000 }, () => {
  it("covers the full shared fixture and generated VRT sets", () => {
    expect(ZERO_DIFF_CASES.filter((testCase) => testCase.group === "shared")).toHaveLength(
      SHARED_FIXTURE_CASES.length,
    );
    expect(ZERO_DIFF_CASES.filter((testCase) => testCase.group === "generated")).toHaveLength(
      VRT_CASES.length,
    );
    expect(ZERO_DIFF_CASES.map(({ name }) => name).sort()).toEqual(
      [
        ...SHARED_FIXTURE_CASES.map(({ name }) => name),
        ...VRT_CASES.map(({ name }) => name),
      ].sort(),
    );
    expect(ZERO_DIFF_CASES.map(({ name }) => name).sort()).toEqual(
      DOCUMENT_PATH_VRT_CASES.map(({ name }) => name).sort(),
    );
  });

  for (const testCase of ZERO_DIFF_CASES) {
    describe(testCase.name, () => {
      it("matches the current parser path PNG output exactly", async () => {
        const fixturePath = join(fixtureDir(testCase.group), testCase.fixture);
        if (!existsSync(fixturePath)) {
          throw new Error(`VRT fixture not found: ${fixturePath}.`);
        }

        const input = readFileSync(fixturePath);
        const options = {
          ...DOCUMENT_PATH_VRT_RENDER_OPTIONS,
          width: DOCUMENT_PATH_VRT_RENDER_WIDTH,
        };
        const parserResults = await convertPptxToPngViaParserPath(input, options);
        const documentResults = await convertPptxToPngViaDocumentPath(input, options);

        expect(
          parserResults.length,
          `${testCase.name}: current parser rendered no slides`,
        ).toBeGreaterThan(0);
        expect(
          documentResults.slides.length,
          `${testCase.name}: document path slide count should match current parser`,
        ).toBe(parserResults.length);
        expect(documentResults.slides.map((slide) => slide.slideNumber)).toEqual(
          parserResults.map((slide) => slide.slideNumber),
        );

        for (const documentResult of documentResults.slides) {
          const parserResult = parserResults.find(
            (slide) => slide.slideNumber === documentResult.slideNumber,
          );
          if (parserResult === undefined) {
            throw new Error(
              `Current parser path did not render ${testCase.name} slide ${documentResult.slideNumber}`,
            );
          }

          const diffPath = join(
            DIFF_DIR,
            `document-path-zero-diff-${testCase.name}-slide${documentResult.slideNumber}-diff.png`,
          );
          const comparison = await compareImageBuffers(
            documentResult.png,
            parserResult.png,
            diffPath,
            {
              pixelThreshold: PIXEL_THRESHOLD,
              mismatchTolerance: MISMATCH_TOLERANCE,
              includeAntiAliased: true,
            },
          );

          expect(
            comparison.passed,
            `${testCase.name} slide ${documentResult.slideNumber}: ` +
              `${comparison.mismatchedPixels}/${comparison.totalPixels} pixels differ. ` +
              "The document path must be pixel-identical to the current parser path.",
          ).toBe(true);
        }
      });
    });
  }
});

function fixtureDir(group: DocumentPathVrtFixtureGroup): string {
  return group === "shared" ? SHARED_FIXTURE_DIR : GENERATED_FIXTURE_DIR;
}
