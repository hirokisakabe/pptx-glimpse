import { existsSync, readFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { convertPptxToPng } from "../../packages/pptx-glimpse/src/converter.js";
import { convertPptxToPngViaDocumentPath } from "../../packages/pptx-glimpse/src/experimental-document-renderer.js";
import { compareImageBuffers } from "../compare-utils.js";
import {
  DOCUMENT_PATH_VRT_EXCLUDED_CASES,
  DOCUMENT_PATH_VRT_OPT_IN_CASES,
  DOCUMENT_PATH_VRT_RENDER_WIDTH,
  DOCUMENT_PATH_VRT_SNAPSHOT_POLICY,
} from "./document-path-cases.js";
import { SHARED_FIXTURE_CASES } from "./vrt-cases.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SHARED_FIXTURE_DIR = join(__dirname, "..", "..", "shared-fixtures");
const DIFF_DIR = join(__dirname, "diffs");

const PIXEL_THRESHOLD = 0;

describe("Document path Visual Regression Tests", { timeout: 60000 }, () => {
  it("documents opt-in cases, excluded cases, and snapshot policy", () => {
    const scopedSharedFixtureNames = [
      ...DOCUMENT_PATH_VRT_OPT_IN_CASES,
      ...DOCUMENT_PATH_VRT_EXCLUDED_CASES.filter(({ fixture }) => !fixture.includes("*")),
    ]
      .map(({ fixture }) => fixture)
      .sort();
    const sharedFixtureNames = SHARED_FIXTURE_CASES.map(({ fixture }) => fixture).sort();

    expect(scopedSharedFixtureNames).toEqual([...new Set(scopedSharedFixtureNames)]);
    expect(scopedSharedFixtureNames).toEqual(sharedFixtureNames);
    expect(DOCUMENT_PATH_VRT_SNAPSHOT_POLICY).toMatchInlineSnapshot(
      `"No committed snapshot update is required: document path VRT compares against the current parser path in-memory until the public default path changes."`,
    );
    expect(
      DOCUMENT_PATH_VRT_OPT_IN_CASES.map(({ name, fixture, slides, tolerance, reason }) => ({
        name,
        fixture,
        slides,
        tolerance,
        reason,
      })),
    ).toMatchInlineSnapshot(`
      [
        {
          "fixture": "real-basic-theme.pptx",
          "name": "real-basic-theme",
          "reason": "Shared fixture already covered by focused document render tests; slide 1 stays within the current CleanDoc shape/text subset.",
          "slides": [
            1,
          ],
          "tolerance": 0.001,
        },
        {
          "fixture": "real-product-page.pptx",
          "name": "real-product-page",
          "reason": "Shared fixture exercises shape/text rendering and intentionally records the current visual parity gap before default-path migration.",
          "slides": [
            1,
          ],
          "tolerance": 0.04,
        },
      ]
    `);
    expect(DOCUMENT_PATH_VRT_EXCLUDED_CASES).toMatchInlineSnapshot(`
      [
        {
          "fixture": "real-financial-report.pptx",
          "name": "real-financial-report",
          "reason": "Deferred because the fixture is broader than the initial selected shared fixture slice and would mix document path dogfood with unrelated parity gaps.",
        },
        {
          "fixture": "sample.pptx",
          "name": "sample",
          "reason": "Deferred until the selected real app fixtures have stable document path parity metrics.",
        },
        {
          "fixture": "sample-issue-387.pptx",
          "name": "sample-issue-387",
          "reason": "Deferred because issue-specific regression fixtures should only opt in after their document path support scope is explicitly reviewed.",
        },
        {
          "fixture": "vrt/snapshot/fixtures/*.pptx",
          "name": "generated snapshot VRT cases",
          "reason": "Deferred because generated cases cover many unsupported tables, charts, groups, effects, and advanced text features outside this issue's selected fixture scope.",
        },
      ]
    `);
  });

  for (const testCase of DOCUMENT_PATH_VRT_OPT_IN_CASES) {
    describe(testCase.name, () => {
      it("tracks current parser vs document path PNG parity", async () => {
        const fixturePath = join(SHARED_FIXTURE_DIR, testCase.fixture);
        if (!existsSync(fixturePath)) {
          throw new Error(`Shared fixture not found: ${fixturePath}.`);
        }

        const input = readFileSync(fixturePath);
        const options = {
          slides: [...testCase.slides],
          width: DOCUMENT_PATH_VRT_RENDER_WIDTH,
          skipSystemFonts: true,
        };
        const currentResults = await convertPptxToPng(input, options);
        const documentResults = await convertPptxToPngViaDocumentPath(input, options);

        expect(currentResults.map((slide) => slide.slideNumber)).toEqual([...testCase.slides]);
        expect(documentResults.slides.map((slide) => slide.slideNumber)).toEqual([
          ...testCase.slides,
        ]);
        expect(uniqueSortedCodes(documentResults.diagnostics)).toEqual(
          [...testCase.expectedDiagnosticCodes].sort(),
        );

        for (const documentResult of documentResults.slides) {
          const currentResult = currentResults.find(
            (slide) => slide.slideNumber === documentResult.slideNumber,
          );
          if (currentResult === undefined) {
            throw new Error(
              `Current parser path did not render ${testCase.name} slide ${documentResult.slideNumber}`,
            );
          }

          const diffPath = join(
            DIFF_DIR,
            `document-path-${testCase.name}-slide${documentResult.slideNumber}-diff.png`,
          );
          const comparison = await compareImageBuffers(
            documentResult.png,
            currentResult.png,
            diffPath,
            {
              pixelThreshold: PIXEL_THRESHOLD,
              mismatchTolerance: testCase.tolerance,
            },
          );

          console.log(
            `[document-path-vrt] ${testCase.name} slide${documentResult.slideNumber}: ` +
              `${(comparison.mismatchPercentage * 100).toFixed(3)}% ` +
              `(tolerance ${(testCase.tolerance * 100).toFixed(1)}%)`,
          );

          expect(
            comparison.passed,
            `${testCase.name} slide ${documentResult.slideNumber}: ` +
              `${(comparison.mismatchPercentage * 100).toFixed(2)}% pixels differ ` +
              `(${comparison.mismatchedPixels}/${comparison.totalPixels}). ` +
              `Tolerance: ${testCase.tolerance * 100}%`,
          ).toBe(true);
        }
      });
    });
  }
});

function uniqueSortedCodes(diagnostics: readonly { readonly code: string }[]): string[] {
  return [...new Set(diagnostics.map((diagnostic) => diagnostic.code))].sort();
}
