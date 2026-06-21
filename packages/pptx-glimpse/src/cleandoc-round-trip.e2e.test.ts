import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  type CleanDocSource,
  createComputedView,
  findTextRunBySourceHandle,
  type MediaPart,
  readPptx,
  replaceTextRunPlainText,
  type SourceParagraph,
  type SourceShape,
  type SourceTextRun,
  writePptx,
} from "../../pptx-glimpse-document/src/experimental.js";
import { renderSlideToSvg } from "../../pptx-glimpse-renderer/src/index.js";
import { adaptComputedViewToRendererModel } from "./cleandoc-renderer-adapter.js";
import { convertPptxToPng, convertPptxToSvg } from "./converter.js";

const SHARED_FIXTURES = ["real-basic-theme.pptx", "real-product-page.pptx"] as const;
const EDITED_TEXT = "Edited 470";

describe("CleanDoc PoC end-to-end round-trip", () => {
  it.each(SHARED_FIXTURES)(
    "renders %s through read -> computed view -> adapter -> SVG",
    (fixtureName) => {
      const rendered = renderDocumentPath(readFixture(fixtureName));

      expect(rendered.source.presentation.slidePartPaths.length).toBeGreaterThan(0);
      expect(rendered.computed.slides).toHaveLength(
        rendered.source.presentation.slidePartPaths.length,
      );
      expect(rendered.adapted.slides).toHaveLength(rendered.computed.slides.length);
      expect(rendered.adapted.slideSize).toBeDefined();
      expect(rendered.svgs).toHaveLength(rendered.adapted.slides.length);
      for (const { svg } of rendered.svgs) {
        expect(svg).toContain("<svg");
        expect(svg).toMatch(/viewBox="0 0 \d+ \d+"/);
        expect(svg).toContain("</svg>");
      }
      for (const diagnostic of rendered.adapted.diagnostics) {
        expect(diagnostic).toMatchObject({
          severity: "warning",
          code: "cleandoc-adapter.raw-element-skipped",
        });
      }
    },
  );

  it.each(SHARED_FIXTURES)("writes and re-reads %s with structural preservation", (fixtureName) => {
    const input = readFixture(fixtureName);
    const original = readPptx(input);
    const output = writePptx(original);
    const reread = readPptx(output);

    expect(reread.presentation.slidePartPaths).toEqual(original.presentation.slidePartPaths);
    expect(reread.presentation.slideSize).toEqual(original.presentation.slideSize);
    expect(reread.packageGraph.contentTypes).toEqual(original.packageGraph.contentTypes);
    expect(reread.packageGraph.relationships).toEqual(original.packageGraph.relationships);
    expect(mediaBytesByPath(reread.packageGraph.media)).toEqual(
      mediaBytesByPath(original.packageGraph.media),
    );

    const inputEntries = unzipSync(input);
    const outputEntries = unzipSync(output);
    for (const partPath of selectedPreservedPartPaths(original)) {
      expect(outputEntries[partPath]).toEqual(inputEntries[partPath]);
    }

    const originalRendered = renderDocumentPath(input).svgs;
    const roundTrippedRendered = renderDocumentPath(output).svgs;
    expect(roundTrippedRendered).toEqual(originalRendered);
  });

  it("writes, re-reads, and renders one text-run edit while preserving unrelated material", () => {
    const input = readFixture("real-product-page.pptx");
    const source = readPptx(input);
    const { paragraph, run } = firstEditableRun(source);
    const handle = run.handle;
    const edited = replaceTextRunPlainText(source, handle, EDITED_TEXT);
    const output = writePptx(edited);
    const reread = readPptx(output);
    const rereadRun = findTextRunBySourceHandle(reread, handle);

    expect(run.text).not.toBe(EDITED_TEXT);
    expect(rereadRun?.text).toBe(EDITED_TEXT);
    expect(rereadRun?.properties).toEqual(run.properties);
    expect(firstParagraphForRun(reread, handle)?.properties).toEqual(paragraph.properties);
    expect(mediaBytesByPath(reread.packageGraph.media)).toEqual(
      mediaBytesByPath(source.packageGraph.media),
    );
    expect(reread.packageGraph.relationships).toEqual(source.packageGraph.relationships);

    const inputEntries = unzipSync(input);
    const outputEntries = unzipSync(output);
    expect(outputEntries[handle.partPath]).not.toEqual(inputEntries[handle.partPath]);
    expect(textDecoder.decode(outputEntries[handle.partPath])).toContain(EDITED_TEXT);
    for (const partPath of selectedPreservedPartPaths(source)) {
      if (partPath === handle.partPath) continue;
      expect(outputEntries[partPath]).toEqual(inputEntries[partPath]);
    }

    const rendered = renderDocumentPath(output);
    expect(rendered.svgs.some(({ svg }) => svg.includes(EDITED_TEXT))).toBe(true);
  });

  it("keeps the public SVG and PNG conversion APIs working for written PPTX output", async () => {
    const input = readFixture("real-basic-theme.pptx");
    const noEditOutput = writePptx(readPptx(input));

    const originalSvgResults = await convertPptxToSvg(input, {
      slides: [1],
      textOutput: "text",
      skipSystemFonts: true,
    });
    const roundTrippedSvgResults = await convertPptxToSvg(noEditOutput, {
      slides: [1],
      textOutput: "text",
      skipSystemFonts: true,
    });
    const pngResults = await convertPptxToPng(noEditOutput, {
      slides: [1],
      width: 240,
      skipSystemFonts: true,
    });

    expect(roundTrippedSvgResults).toEqual(originalSvgResults);
    expect(pngResults.map((result) => result.slideNumber)).toEqual([1]);
    expect(pngResults[0]).toMatchObject({ width: 240 });
    expect([...pngResults[0].png.subarray(0, 4)]).toEqual([0x89, 0x50, 0x4e, 0x47]);

    const editedInput = readFixture("real-product-page.pptx");
    const editedSource = readPptx(editedInput);
    const editedOutput = writePptx(
      replaceTextRunPlainText(editedSource, firstEditableRun(editedSource).run.handle, EDITED_TEXT),
    );
    const editedSvgResults = await convertPptxToSvg(editedOutput, {
      slides: [1],
      textOutput: "text",
      skipSystemFonts: true,
    });

    expect(editedSvgResults.map((result) => result.slideNumber)).toEqual([1]);
    expect(editedSvgResults[0].svg).toContain(EDITED_TEXT);

    // VRT snapshot updates are unnecessary: this PR does not change production
    // rendering, and no-edit writer output is asserted to keep public SVG stable
    // for a selected shared fixture.
  });
});

const textDecoder = new TextDecoder();

interface RenderedDocumentPath {
  readonly source: CleanDocSource;
  readonly computed: ReturnType<typeof createComputedView>;
  readonly adapted: ReturnType<typeof adaptComputedViewToRendererModel>;
  readonly svgs: readonly { readonly slideNumber: number; readonly svg: string }[];
}

function readFixture(name: (typeof SHARED_FIXTURES)[number]): Buffer {
  return readFileSync(fileURLToPath(new URL(`../../../shared-fixtures/${name}`, import.meta.url)));
}

function renderDocumentPath(input: Uint8Array): RenderedDocumentPath {
  const source = readPptx(input);
  const computed = createComputedView(source);
  const adapted = adaptComputedViewToRendererModel(computed);
  const slideSize = adapted.slideSize;
  if (slideSize === undefined) {
    throw new Error("CleanDoc render path requires a slide size for selected fixtures");
  }

  return {
    source,
    computed,
    adapted,
    svgs: adapted.slides.map((slide) => ({
      slideNumber: slide.slideNumber,
      svg: renderSlideToSvg(slide, slideSize),
    })),
  };
}

function mediaBytesByPath(media: readonly MediaPart[]): Record<string, readonly number[]> {
  return Object.fromEntries(media.map((part) => [part.partPath, [...part.bytes]]));
}

function selectedPreservedPartPaths(source: CleanDocSource): string[] {
  return [
    ...source.presentation.slidePartPaths,
    ...source.slideLayouts.map((layout) => layout.partPath),
    ...source.slideMasters.map((master) => master.partPath),
    ...source.themes.map((theme) => theme.partPath),
  ];
}

function firstEditableRun(source: CleanDocSource): {
  readonly paragraph: SourceParagraph;
  readonly run: EditableTextRun;
} {
  for (const slide of source.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind !== "shape") continue;
      for (const paragraph of shape.textBody?.paragraphs ?? []) {
        const run = paragraph.runs.find(isEditableTextRun);
        if (run?.handle !== undefined) return { paragraph, run };
      }
    }
  }
  throw new Error("No editable text run found in selected CleanDoc fixture");
}

type EditableTextRun = SourceTextRun & {
  readonly handle: NonNullable<SourceTextRun["handle"]>;
};

function isEditableTextRun(run: SourceTextRun): run is EditableTextRun {
  return run.text.trim() !== "" && run.handle !== undefined;
}

function firstParagraphForRun(
  source: CleanDocSource,
  handle: SourceTextRun["handle"],
): SourceParagraph | undefined {
  if (handle === undefined) return undefined;
  for (const slide of source.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind !== "shape") continue;
      const paragraph = findParagraphForRun(shape, handle);
      if (paragraph !== undefined) return paragraph;
    }
  }
  return undefined;
}

function findParagraphForRun(shape: SourceShape, handle: NonNullable<SourceTextRun["handle"]>) {
  for (const paragraph of shape.textBody?.paragraphs ?? []) {
    if (
      paragraph.runs.some(
        (run) =>
          run.handle?.partPath === handle.partPath &&
          run.handle.nodeId === handle.nodeId &&
          run.handle.relationshipId === handle.relationshipId &&
          run.handle.orderingSlot === handle.orderingSlot,
      )
    ) {
      return paragraph;
    }
  }
  return undefined;
}
