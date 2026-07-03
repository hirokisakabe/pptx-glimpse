import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { type ConvertOptions, convertPptxToPng } from "../packages/core/src/converter.js";
import { type PptxSourceModel, readPptx, writePptx } from "../packages/document/src/index.js";
import { compareImageBuffers, type CompareOptions, type CompareResult } from "./compare-utils.js";
import { getVrtRenderOptions } from "./snapshot/render-options.js";

type MaybePromise<T> = T | Promise<T>;

export interface EditEquivalenceFixture {
  readonly name: string;
  readonly create: () => MaybePromise<Buffer | Uint8Array>;
}

export interface EditEquivalenceOperation {
  readonly name: string;
  readonly apply: (source: PptxSourceModel) => PptxSourceModel;
}

interface EditEquivalenceCase {
  readonly name: string;
  readonly sourceFixture: EditEquivalenceFixture;
  readonly operations: readonly EditEquivalenceOperation[];
  readonly expectedFixture: EditEquivalenceFixture;
  readonly renderOptionsProvider?: () => MaybePromise<ConvertOptions>;
  readonly renderOptions?: ConvertOptions;
  readonly compareOptions?: Partial<CompareOptions>;
  readonly diffDir?: string;
}

interface EditEquivalenceResult {
  readonly slideNumber: number;
  readonly comparison: CompareResult;
}

const DEFAULT_COMPARE_OPTIONS = {
  pixelThreshold: 0,
  mismatchTolerance: 0,
  includeAntiAliased: true,
} as const satisfies CompareOptions;

const DEFAULT_DIFF_DIR = join(tmpdir(), "pptx-glimpse-edit-equivalence-diffs");

/**
 * Defines edit-equivalence rendering tests for source-model editing workflows.
 *
 * The harness checks the semantic, rendered result of supported edits by comparing
 * "source fixture + edit operations + writer" against "expected fixture authored in
 * the same final state" inside one process and one renderer setup. It intentionally
 * cannot detect damage to non-rendered PPTX content such as animations, notes,
 * comments, unsupported raw OOXML, or package relationships that do not affect the
 * rendered slide. Those concerns remain covered by writer structural-preservation
 * tests, not by this rendering-equivalence oracle. Cases use the deterministic VRT
 * render options by default, but may inject lighter render options when the case does
 * not need VRT font subsets.
 */
export function defineEditEquivalenceTests(cases: readonly EditEquivalenceCase[]): void {
  describe("edit equivalence rendering", { timeout: 60000 }, () => {
    for (const testCase of cases) {
      it(`${testCase.name}: edited fixture renders like expected fixture`, async () => {
        await expect(assertEditEquivalence(testCase)).resolves.toEqual(
          expect.arrayContaining([
            expect.objectContaining({
              comparison: expect.objectContaining({ passed: true }),
            }),
          ]),
        );
      });
    }
  });
}

export async function assertEditEquivalence(
  testCase: EditEquivalenceCase,
): Promise<readonly EditEquivalenceResult[]> {
  const source = await testCase.sourceFixture.create();
  const expected = await testCase.expectedFixture.create();
  const edited = testCase.operations.reduce(
    (document, operation) => operation.apply(document),
    readSourceModel(source),
  );
  const editedPptx = writePptx(edited);
  const baseRenderOptions = await (testCase.renderOptionsProvider?.() ?? getVrtRenderOptions());
  const renderOptions = {
    ...baseRenderOptions,
    ...testCase.renderOptions,
  };
  const actualReport = await convertPptxToPng(editedPptx, renderOptions);
  const expectedReport = await convertPptxToPng(expected, renderOptions);

  expect(actualReport.slides.map((slide) => slide.slideNumber)).toEqual(
    expectedReport.slides.map((slide) => slide.slideNumber),
  );

  const compareOptions = {
    ...DEFAULT_COMPARE_OPTIONS,
    ...testCase.compareOptions,
  };
  const results: EditEquivalenceResult[] = [];

  for (const actualSlide of actualReport.slides) {
    const expectedSlide = expectedReport.slides.find(
      (slide) => slide.slideNumber === actualSlide.slideNumber,
    );
    if (expectedSlide === undefined) {
      throw new Error(`Expected fixture did not render slide ${actualSlide.slideNumber}.`);
    }

    const comparison = await compareImageBuffers(
      actualSlide.png,
      expectedSlide.png,
      join(
        testCase.diffDir ?? DEFAULT_DIFF_DIR,
        `${slugify(testCase.name)}-slide${actualSlide.slideNumber}-diff.png`,
      ),
      compareOptions,
    );
    if (!comparison.passed) {
      throw new Error(
        `${testCase.name} slide ${actualSlide.slideNumber}: ${(
          comparison.mismatchPercentage * 100
        ).toFixed(2)}% pixels differ (${comparison.mismatchedPixels}/${comparison.totalPixels})`,
      );
    }

    results.push({ slideNumber: actualSlide.slideNumber, comparison });
  }

  return results;
}

function readSourceModel(input: Buffer | Uint8Array): PptxSourceModel {
  return readPptx(input);
}

function slugify(value: string): string {
  return value.replace(/[^A-Za-z0-9_-]+/g, "-").replace(/^-+|-+$/g, "") || "case";
}
