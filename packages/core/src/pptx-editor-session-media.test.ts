import { asOoxmlPercent, readPptx } from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { renderPptxSourceModelToSvg } from "./converter.js";
import { createPptxEditorSession } from "./index.js";
import { configurePptxEditorSessionRenderer } from "./pptx-editor-session.js";
import {
  BLUE_PNG,
  buildImageFixture,
  mediaBytes,
  RED_PNG,
} from "./pptx-editor-session.test-helpers.js";

describe("PptxEditorSession - media", () => {
  it("isolates shared media replacements without warnings", async () => {
    const editor = await createPptxEditorSession(await buildImageFixture(), {
      skipSystemFonts: true,
    });
    const image = editor.shapes(1).find((shape) => shape.kind === "image");
    if (image?.handle === undefined) throw new Error("image handle not found");
    expect(image.editableImageReplacement).toEqual({
      contentType: "image/png",
      accept: "image/png,.png",
      mediaPartPath: "ppt/media/image1.png",
      sharedReferenceCount: 2,
    });

    const result = await editor.apply({
      kind: "replaceImage",
      handle: image.handle,
      bytes: BLUE_PNG,
    });

    expect(result.warnings).toBeUndefined();
    expect(mediaBytes(editor.document, "ppt/media/image1.png")).toEqual(RED_PNG);
    expect(mediaBytes(editor.document, "ppt/media/image2.png")).toEqual(BLUE_PNG);
    expect(
      editor.shapes(1).find((shape) => shape.name === "Shared Picture A")?.editableImageReplacement,
    ).toMatchObject({ mediaPartPath: "ppt/media/image2.png", sharedReferenceCount: 1 });
    expect(
      editor.shapes(1).find((shape) => shape.name === "Shared Picture B")?.editableImageReplacement,
    ).toMatchObject({ mediaPartPath: "ppt/media/image1.png", sharedReferenceCount: 1 });
  });

  it("applies picture crop, rerenders the target slide, and saves the srcRect", async () => {
    const renderCalls: Array<readonly number[] | undefined> = [];
    configurePptxEditorSessionRenderer((source, options) => {
      renderCalls.push(options?.slides);
      return renderPptxSourceModelToSvg(source, options);
    });
    try {
      const editor = await createPptxEditorSession(await buildImageFixture(), {
        skipSystemFonts: true,
      });
      const image = editor.shapes(1).find((shape) => shape.kind === "image");
      if (image?.handle === undefined) throw new Error("image handle not found");
      const beforeSvg = editor.slides[0]?.svg;

      await editor.apply({
        kind: "setPictureCrop",
        handle: image.handle,
        left: asOoxmlPercent(25000),
        top: asOoxmlPercent(10000),
      });

      expect(renderCalls).toEqual([undefined, [1]]);
      expect(editor.slides[0]?.svg).not.toBe(beforeSvg);
      const saved = readPptx(editor.save().pptx);
      expect(saved.slides[0]?.shapes.find((shape) => shape.kind === "image")?.crop).toEqual({
        left: 25000,
        top: 10000,
      });
    } finally {
      configurePptxEditorSessionRenderer(renderPptxSourceModelToSvg);
    }
  });

  it("rerenders only the picture owner after copy-on-write replacement", async () => {
    const renderCalls: Array<readonly number[] | undefined> = [];
    configurePptxEditorSessionRenderer((source, options) => {
      renderCalls.push(options?.slides);
      return renderPptxSourceModelToSvg(source, options);
    });
    try {
      const editor = await createPptxEditorSession(
        await buildImageFixture({ includeSecondSlide: true }),
        { skipSystemFonts: true },
      );
      const image = editor.shapes(1).find((shape) => shape.kind === "image");
      if (image?.handle === undefined) throw new Error("image handle not found");

      await editor.apply({ kind: "replaceImage", handle: image.handle, bytes: BLUE_PNG });

      expect(renderCalls).toEqual([undefined, [1]]);
    } finally {
      configurePptxEditorSessionRenderer(renderPptxSourceModelToSvg);
    }
  });

  it("counts unparsed image relationships in image replacement metadata", async () => {
    const editor = await createPptxEditorSession(
      await buildImageFixture({ includeUnusedImageRelationship: true }),
      { skipSystemFonts: true },
    );
    const image = editor.shapes(1).find((shape) => shape.kind === "image");

    expect(image?.editableImageReplacement?.sharedReferenceCount).toBe(3);
  });

  it("counts preserved image fills in image replacement metadata", async () => {
    const editor = await createPptxEditorSession(
      await buildImageFixture({ includeImageFill: true }),
      { skipSystemFonts: true },
    );
    const image = editor.shapes(1).find((shape) => shape.kind === "image");

    expect(image?.editableImageReplacement?.sharedReferenceCount).toBe(3);
  });
});
