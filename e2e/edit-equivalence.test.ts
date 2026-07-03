import { describe, expect, it } from "vitest";

import { replaceTextRunPlainText } from "../packages/document/src/index.js";
import {
  assertEditEquivalence,
  defineEditEquivalenceTests,
  type EditEquivalenceFixture,
  type EditEquivalenceOperation,
} from "../vrt/edit-equivalence.js";
import {
  buildPptx,
  shapeXml,
  slideRelsXml,
  textBodyXmlHelper,
  wrapSlideXml,
} from "../vrt/snapshot/create-fixtures.js";

const originalTextRunFixture = textRunFixture("source text-run fixture", "Original text");
const editedTextRunFixture = textRunFixture("expected edited text-run fixture", "Edited text");
const mismatchedTextRunFixture = textRunFixture(
  "mismatched expected text-run fixture",
  "Unexpected text",
);

const replaceFirstRunWithEditedText = {
  name: "replace first text run with edited text",
  apply: (source) => {
    const run =
      source.slides[0]?.shapes[0]?.kind === "shape"
        ? source.slides[0].shapes[0].textBody?.paragraphs[0]?.runs[0]
        : undefined;
    if (run?.handle === undefined) throw new Error("editable text run not found");
    return replaceTextRunPlainText(source, run.handle, "Edited text");
  },
} as const satisfies EditEquivalenceOperation;

defineEditEquivalenceTests([
  {
    name: "text-run replacement",
    sourceFixture: originalTextRunFixture,
    operations: [replaceFirstRunWithEditedText],
    expectedFixture: editedTextRunFixture,
    renderOptions: { width: 480 },
  },
]);

describe("edit equivalence rendering oracle", { timeout: 60000 }, () => {
  it("fails when edit operations and expected fixture intentionally disagree", async () => {
    await expect(
      assertEditEquivalence({
        name: "text-run replacement mismatch proof",
        sourceFixture: originalTextRunFixture,
        operations: [replaceFirstRunWithEditedText],
        expectedFixture: mismatchedTextRunFixture,
        renderOptions: { width: 480 },
      }),
    ).rejects.toThrow(/pixels differ/);
  });
});

function textRunFixture(name: string, text: string): EditEquivalenceFixture {
  return {
    name,
    create: async () =>
      await buildPptx({
        slides: [
          {
            xml: wrapSlideXml(textRunShapeXml(text)),
            rels: slideRelsXml(),
          },
        ],
      }),
  };
}

function textRunShapeXml(text: string): string {
  return shapeXml(2, "Editable Text", {
    preset: "rect",
    x: 914400,
    y: 914400,
    cx: 5486400,
    cy: 1371600,
    fillXml: `<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`,
    outlineXml: `<a:ln w="12700"><a:solidFill><a:srgbClr val="2F528F"/></a:solidFill></a:ln>`,
    textBodyXml: textBodyXmlHelper(text, {
      fontSize: 28,
      color: "FFFFFF",
      align: "ctr",
    }),
  });
}
