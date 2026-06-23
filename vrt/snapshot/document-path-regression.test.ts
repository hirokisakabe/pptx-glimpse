import { existsSync, readFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { convertPptxToPng } from "../../packages/pptx-glimpse/src/converter.js";
import { convertPptxToPngViaDocumentPath } from "../../packages/pptx-glimpse/src/experimental-document-renderer.js";
import { compareImageBuffers } from "../compare-utils.js";
import {
  DOCUMENT_PATH_VRT_BLOCKER_ISSUES,
  DOCUMENT_PATH_VRT_CASES,
  DOCUMENT_PATH_VRT_GENERATED_CASES,
  DOCUMENT_PATH_VRT_RENDER_WIDTH,
  DOCUMENT_PATH_VRT_SHARED_CASES,
  DOCUMENT_PATH_VRT_SNAPSHOT_POLICY,
  type DocumentPathVrtFixtureGroup,
} from "./document-path-cases.js";
import { SHARED_FIXTURE_CASES, VRT_CASES } from "./vrt-cases.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SHARED_FIXTURE_DIR = join(__dirname, "..", "..", "shared-fixtures");
const GENERATED_FIXTURE_DIR = join(__dirname, "fixtures");
const DIFF_DIR = join(__dirname, "diffs");

const PIXEL_THRESHOLD = 0;

describe("Document path Visual Regression Tests", { timeout: 60000 }, () => {
  it("covers the full shared fixture and generated VRT sets", () => {
    const scopedSharedFixtureNames = DOCUMENT_PATH_VRT_SHARED_CASES.map(
      ({ fixture }) => fixture,
    ).sort();
    const sharedFixtureNames = SHARED_FIXTURE_CASES.map(({ fixture }) => fixture).sort();
    const scopedGeneratedFixtureNames = DOCUMENT_PATH_VRT_GENERATED_CASES.map(
      ({ fixture }) => fixture,
    ).sort();
    const generatedFixtureNames = VRT_CASES.map(({ fixture }) => fixture).sort();
    const scopedCaseNames = DOCUMENT_PATH_VRT_CASES.map(({ name }) => name).sort();

    expect(scopedSharedFixtureNames).toEqual([...new Set(scopedSharedFixtureNames)]);
    expect(scopedGeneratedFixtureNames).toEqual([...new Set(scopedGeneratedFixtureNames)]);
    expect(scopedCaseNames).toEqual([...new Set(scopedCaseNames)]);
    expect(scopedSharedFixtureNames).toEqual(sharedFixtureNames);
    expect(scopedGeneratedFixtureNames).toEqual(generatedFixtureNames);
    expect(DOCUMENT_PATH_VRT_SNAPSHOT_POLICY).toMatchInlineSnapshot(
      `"No committed snapshot update is required: document path VRT compares against the current parser path in-memory until the public default path changes."`,
    );
  });

  it("keeps every non-zero or diagnostic-emitting gap linked to a blocker issue", () => {
    const blockerIssueNumbers = new Set(Object.values(DOCUMENT_PATH_VRT_BLOCKER_ISSUES));

    for (const testCase of DOCUMENT_PATH_VRT_CASES) {
      for (const issue of testCase.blockerIssues) {
        expect(blockerIssueNumbers.has(issue), `${testCase.name}: unknown blocker #${issue}`).toBe(
          true,
        );
      }
      if (testCase.mismatchTolerance > 0 || testCase.expectedDiagnosticCodes.length > 0) {
        expect(
          testCase.blockerIssues.length,
          `${testCase.name}: expected blocker issue for residual gap`,
        ).toBeGreaterThan(0);
      }
    }
  });

  for (const testCase of DOCUMENT_PATH_VRT_CASES) {
    describe(testCase.name, () => {
      it("tracks current parser vs document path PNG parity", async () => {
        const fixturePath = join(fixtureDir(testCase.group), testCase.fixture);
        if (!existsSync(fixturePath)) {
          throw new Error(`VRT fixture not found: ${fixturePath}.`);
        }

        const input = readFileSync(fixturePath);
        const options = {
          width: DOCUMENT_PATH_VRT_RENDER_WIDTH,
          skipSystemFonts: true,
        };
        const currentResults = await convertPptxToPng(input, options);
        const documentResults = await convertPptxToPngViaDocumentPath(input, options);

        expect(
          currentResults.length,
          `${testCase.name}: current parser rendered no slides`,
        ).toBeGreaterThan(0);
        expect(
          documentResults.slides.length,
          `${testCase.name}: document path slide count should match current parser`,
        ).toBe(currentResults.length);
        expect(documentResults.slides.map((slide) => slide.slideNumber)).toEqual(
          currentResults.map((slide) => slide.slideNumber),
        );
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
              mismatchTolerance: testCase.mismatchTolerance,
            },
          );

          console.log(
            `[document-path-vrt] ${testCase.name} slide${documentResult.slideNumber}: ` +
              `${(comparison.mismatchPercentage * 100).toFixed(3)}% ` +
              `(tolerance ${(testCase.mismatchTolerance * 100).toFixed(1)}%)` +
              blockerSuffix(testCase.blockerIssues),
          );

          expect(
            comparison.passed,
            `${testCase.name} slide ${documentResult.slideNumber}: ` +
              `${(comparison.mismatchPercentage * 100).toFixed(2)}% pixels differ ` +
              `(${comparison.mismatchedPixels}/${comparison.totalPixels}). ` +
              `Tolerance: ${testCase.mismatchTolerance * 100}%. ` +
              `Blockers: ${formatBlockers(testCase.blockerIssues)}`,
          ).toBe(true);
        }
      });
    });
  }
});

function uniqueSortedCodes(diagnostics: readonly { readonly code: string }[]): string[] {
  return [...new Set(diagnostics.map((diagnostic) => diagnostic.code))].sort();
}

function fixtureDir(group: DocumentPathVrtFixtureGroup): string {
  return group === "shared" ? SHARED_FIXTURE_DIR : GENERATED_FIXTURE_DIR;
}

function blockerSuffix(blockerIssues: readonly number[]): string {
  if (blockerIssues.length === 0) return "";
  return `; blockers ${formatBlockers(blockerIssues)}`;
}

function formatBlockers(blockerIssues: readonly number[]): string {
  return blockerIssues.length === 0 ? "none" : blockerIssues.map((issue) => `#${issue}`).join(", ");
}
