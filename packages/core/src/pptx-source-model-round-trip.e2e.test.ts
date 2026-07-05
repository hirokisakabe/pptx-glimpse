import { Buffer } from "node:buffer";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import {
  asEmu,
  createComputedView,
  deleteSlide,
  duplicateSlide,
  findShapeNodeBySourceHandle,
  findTextRunBySourceHandle,
  type MediaPart,
  type PptxSourceModel,
  readPptx,
  replaceImageBytes,
  replaceTextRunPlainText,
  type SourceImage,
  type SourceParagraph,
  type SourceShape,
  type SourceShapeNode,
  type SourceTextRun,
  updateShapeTransform,
  writePptx,
} from "@pptx-glimpse/document";
import { renderSlideToSvg } from "@pptx-glimpse/renderer";
import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { convertPptxToPng, convertPptxToSvg } from "./converter.js";
import { adaptComputedViewToRendererModel } from "./pptx-computed-view-renderer-adapter.js";

const SHARED_FIXTURES = ["real-basic-theme.pptx", "real-product-page.pptx"] as const;
const EDITED_TEXT = "Edited 470";
const BLUE_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAE0lEQVR4nGNkYPjPAANMcBZeDgAx0wEH1s7nlgAAAABJRU5ErkJggg==";
const BLUE_PNG = new Uint8Array(Buffer.from(BLUE_PNG_BASE64, "base64"));

describe("PptxSourceModel PoC end-to-end round-trip", () => {
  it.each(SHARED_FIXTURES)(
    "renders %s through read -> computed view -> adapter -> SVG",
    (fixtureName) => {
      const rendered = renderSourceModelPipeline(readFixture(fixtureName));

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
          code: "pptx-computed-view-adapter.raw-element-skipped",
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

    const originalRendered = renderSourceModelPipeline(input).svgs;
    const roundTrippedRendered = renderSourceModelPipeline(output).svgs;
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

    const rendered = renderSourceModelPipeline(output);
    expect(rendered.svgs.some(({ svg }) => svg.includes(EDITED_TEXT))).toBe(true);
  });

  it("keeps the public SVG and PNG conversion APIs working for written PPTX output", async () => {
    const input = readFixture("real-basic-theme.pptx");
    const noEditOutput = writePptx(readPptx(input));

    const { slides: originalSvgResults } = await convertPptxToSvg(input, {
      textOutput: "text",
      skipSystemFonts: true,
    });
    const { slides: roundTrippedSvgResults } = await convertPptxToSvg(noEditOutput, {
      textOutput: "text",
      skipSystemFonts: true,
    });
    const { slides: pngResults } = await convertPptxToPng(noEditOutput, {
      width: 240,
      skipSystemFonts: true,
    });

    expect(roundTrippedSvgResults).toEqual(originalSvgResults);
    expect(pngResults.map((result) => result.slideNumber)).toEqual(
      originalSvgResults.map((result) => result.slideNumber),
    );
    for (const pngResult of pngResults) {
      expect(pngResult).toMatchObject({ width: 240 });
      expect([...pngResult.png.subarray(0, 4)]).toEqual([0x89, 0x50, 0x4e, 0x47]);
    }

    const editedInput = readFixture("real-product-page.pptx");
    const editedSource = readPptx(editedInput);
    const editable = firstEditableRun(editedSource);
    const editedOutput = writePptx(
      replaceTextRunPlainText(editedSource, editable.run.handle, EDITED_TEXT),
    );
    const { slides: editedSvgResults } = await convertPptxToSvg(editedOutput, {
      slides: [editable.slideNumber],
      textOutput: "text",
      skipSystemFonts: true,
    });

    expect(editedSvgResults.map((result) => result.slideNumber)).toEqual([editable.slideNumber]);
    expect(editedSvgResults[0].svg).toContain(EDITED_TEXT);

    // No-edit writer output is asserted to keep public SVG stable for the
    // selected shared fixture, so this e2e coverage does not need VRT snapshots.
  });

  it("writes, re-reads, and renders one image replacement through convertPptxToSvg", async () => {
    const input = readFixture("real-basic-theme.pptx");
    const source = readPptx(input);
    const image = firstImage(source);
    const editedOutput = writePptx(replaceImageBytes(source, image.handle!, BLUE_PNG));
    const reread = readPptx(editedOutput);
    const rereadMedia = reread.packageGraph.media.find(
      (part) => part.partPath === "ppt/media/image1.png",
    );
    const { slides } = await convertPptxToSvg(editedOutput, {
      textOutput: "text",
      skipSystemFonts: true,
    });

    expect(rereadMedia?.bytes).toEqual(BLUE_PNG);
    expect(
      slides.some((slide) => slide.svg.includes(`data:image/png;base64,${BLUE_PNG_BASE64}`)),
    ).toBe(true);
  });

  it("writes, re-reads, and renders one shape xfrm edit through convertPptxToSvg", async () => {
    const input = readFixture("real-product-page.pptx");
    const source = readPptx(input);
    const editable = firstTransformShape(source);
    const editedOutput = writePptx(
      updateShapeTransform(source, editable.shape.handle, {
        offsetX: asEmu(914400),
        offsetY: asEmu(1828800),
        width: asEmu(2743200),
        height: asEmu(914400),
      }),
    );
    const reread = readPptx(editedOutput);
    const rereadShape = findShapeNodeBySourceHandle(reread, editable.shape.handle);

    expect(rereadShape?.transform).toMatchObject({
      offsetX: 914400,
      offsetY: 1828800,
      width: 2743200,
      height: 914400,
    });

    const { slides } = await convertPptxToSvg(editedOutput, {
      slides: [editable.slideNumber],
      textOutput: "text",
      skipSystemFonts: true,
    });

    expect(slides).toHaveLength(1);
    expect(slides[0].svg).toContain('transform="translate(96, 192)"');
    expect(slides[0].svg).toContain('width="288" height="96"');
  });

  it("duplicates and deletes slides through writer output while keeping render order stable", async () => {
    const input = readFixture("real-basic-theme.pptx");
    const source = readPptx(input);
    const duplicatedOutput = writePptx(duplicateSlide(source, source.slides[0].handle!));
    const duplicated = readPptx(duplicatedOutput);
    const { slides: duplicatedSvgs } = await convertPptxToSvg(duplicatedOutput, {
      textOutput: "text",
      skipSystemFonts: true,
    });

    expect(duplicated.presentation.slidePartPaths).toEqual([
      source.slides[0].partPath,
      "ppt/slides/slide3.xml",
      source.slides[1].partPath,
    ]);
    expect(duplicatedSvgs).toHaveLength(3);
    expect(duplicatedSvgs[1].svg).toEqual(duplicatedSvgs[0].svg);

    const deletedOutput = writePptx(deleteSlide(duplicated, duplicated.slides[0].handle!));
    const deleted = readPptx(deletedOutput);
    const { slides: deletedSvgs } = await convertPptxToSvg(deletedOutput, {
      textOutput: "text",
      skipSystemFonts: true,
    });

    expect(deleted.presentation.slidePartPaths).toEqual([
      "ppt/slides/slide3.xml",
      source.slides[1].partPath,
    ]);
    expect(deletedSvgs).toHaveLength(2);
    expect(deletedSvgs[0].svg).toEqual(duplicatedSvgs[0].svg);
    expect(deletedSvgs[1].svg).toEqual(duplicatedSvgs[2].svg);
  });
});

const textDecoder = new TextDecoder();

interface RenderedSourceModelPipeline {
  readonly source: PptxSourceModel;
  readonly computed: ReturnType<typeof createComputedView>;
  readonly adapted: ReturnType<typeof adaptComputedViewToRendererModel>;
  readonly svgs: readonly { readonly slideNumber: number; readonly svg: string }[];
}

function readFixture(name: (typeof SHARED_FIXTURES)[number]): Buffer {
  return readFileSync(fileURLToPath(new URL(`../../../shared-fixtures/${name}`, import.meta.url)));
}

function renderSourceModelPipeline(input: Uint8Array): RenderedSourceModelPipeline {
  const source = readPptx(input);
  const computed = createComputedView(source);
  const adapted = adaptComputedViewToRendererModel(computed);
  const slideSize = adapted.slideSize;
  if (slideSize === undefined) {
    throw new Error("PptxSourceModel render path requires a slide size for selected fixtures");
  }

  return {
    source,
    computed,
    adapted,
    svgs: adapted.slides.map((slide) => ({
      slideNumber: slide.slideNumber,
      svg: normalizeEffectFilterIds(renderSlideToSvg(slide, slideSize)),
    })),
  };
}

function normalizeEffectFilterIds(svg: string): string {
  const ids = new Map<string, string>();
  return svg.replace(
    /(?:blip-)?effect-[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/g,
    (id) => {
      const existing = ids.get(id);
      if (existing !== undefined) return existing;
      const normalized = `effect-${ids.size}`;
      ids.set(id, normalized);
      return normalized;
    },
  );
}

function mediaBytesByPath(media: readonly MediaPart[]): Record<string, readonly number[]> {
  return Object.fromEntries(media.map((part) => [part.partPath, [...part.bytes]]));
}

function selectedPreservedPartPaths(source: PptxSourceModel): string[] {
  return [
    ...source.presentation.slidePartPaths,
    ...source.slideLayouts.map((layout) => layout.partPath),
    ...source.slideMasters.map((master) => master.partPath),
    ...source.themes.map((theme) => theme.partPath),
  ];
}

function firstEditableRun(source: PptxSourceModel): {
  readonly slideNumber: number;
  readonly paragraph: SourceParagraph;
  readonly run: EditableTextRun;
} {
  for (const slide of source.slides) {
    const slideNumber = source.presentation.slidePartPaths.indexOf(slide.partPath) + 1;
    if (slideNumber <= 0) continue;
    for (const shape of slide.shapes) {
      if (shape.kind !== "shape") continue;
      for (const paragraph of shape.textBody?.paragraphs ?? []) {
        const run = paragraph.runs.find(isEditableTextRun);
        if (run?.handle !== undefined) return { slideNumber, paragraph, run };
      }
    }
  }
  throw new Error("No editable text run found in selected PptxSourceModel fixture");
}

function firstTransformShape(source: PptxSourceModel): {
  readonly slideNumber: number;
  readonly shape: TransformEditableShape;
} {
  for (const slide of source.slides) {
    const slideNumber = source.presentation.slidePartPaths.indexOf(slide.partPath) + 1;
    if (slideNumber <= 0) continue;
    const shape = slide.shapes.find(isTransformEditableShape);
    if (shape !== undefined) return { slideNumber, shape };
  }
  throw new Error("No editable shape transform found in selected PptxSourceModel fixture");
}

function firstImage(source: PptxSourceModel): SourceImage {
  for (const slide of source.slides) {
    const image = slide.shapes.find((shape): shape is SourceImage => shape.kind === "image");
    if (image !== undefined) return image;
  }
  throw new Error("No editable image found in selected PptxSourceModel fixture");
}

type TransformEditableShape = SourceShapeNode & {
  readonly handle: NonNullable<SourceShapeNode["handle"]>;
  readonly transform: NonNullable<SourceShape["transform"]>;
};

function isTransformEditableShape(shape: SourceShapeNode): shape is TransformEditableShape {
  return shape.kind !== "raw" && shape.handle !== undefined && shape.transform !== undefined;
}

type EditableTextRun = SourceTextRun & {
  readonly handle: NonNullable<SourceTextRun["handle"]>;
};

function isEditableTextRun(run: SourceTextRun): run is EditableTextRun {
  return run.text.trim() !== "" && run.handle !== undefined;
}

function firstParagraphForRun(
  source: PptxSourceModel,
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
