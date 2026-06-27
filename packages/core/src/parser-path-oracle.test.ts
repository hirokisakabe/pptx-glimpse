import type { ShapeElement, Slide, SlideElement } from "@pptx-glimpse/renderer";
import { describe, expect, it } from "vitest";

import { buildEffectiveSlideElements } from "./parser-path-oracle.js";
import type { ParsedSlide } from "./pptx-data-parser.js";
import { unsafeTypeAssertion } from "./unsafe-type-assertion.js";

describe("buildEffectiveSlideElements", () => {
  it("omits master elements when the slide hides master shapes", () => {
    const master = shape("master");
    const layout = shape("layout");
    const slideElement = shape("slide");

    const result = buildEffectiveSlideElements(
      parsedSlide({
        showMasterSp: false,
        layoutShowMasterSp: true,
        masterElements: [master],
        layoutElements: [layout],
        slideElements: [slideElement],
      }),
    );

    expect(result).toEqual([layout, slideElement]);
  });

  it("omits master elements when the layout hides master shapes", () => {
    const master = shape("master");
    const layout = shape("layout");
    const slideElement = shape("slide");

    const result = buildEffectiveSlideElements(
      parsedSlide({
        showMasterSp: true,
        layoutShowMasterSp: false,
        masterElements: [master],
        layoutElements: [layout],
        slideElements: [slideElement],
      }),
    );

    expect(result).toEqual([layout, slideElement]);
  });

  it("filters template placeholders and empty slide placeholders", () => {
    const masterPlaceholder = shape("master-placeholder", { placeholderType: "body" });
    const masterShape = shape("master-shape");
    const layoutPlaceholder = shape("layout-placeholder", { placeholderType: "title" });
    const layoutShape = shape("layout-shape");
    const emptySlidePlaceholder = shape("empty-slide-placeholder", { placeholderType: "body" });
    const filledSlidePlaceholder = shape("filled-slide-placeholder", {
      placeholderType: "body",
      text: "Keep me",
    });

    const result = buildEffectiveSlideElements(
      parsedSlide({
        showMasterSp: true,
        layoutShowMasterSp: true,
        masterElements: [masterPlaceholder, masterShape],
        layoutElements: [layoutPlaceholder, layoutShape],
        slideElements: [emptySlidePlaceholder, filledSlidePlaceholder],
      }),
    );

    expect(result).toEqual([masterShape, layoutShape, filledSlidePlaceholder]);
  });
});

function parsedSlide({
  showMasterSp,
  layoutShowMasterSp,
  masterElements,
  layoutElements,
  slideElements,
}: {
  showMasterSp: boolean;
  layoutShowMasterSp: boolean;
  masterElements: SlideElement[];
  layoutElements: SlideElement[];
  slideElements: SlideElement[];
}): ParsedSlide {
  return {
    slide: {
      slideNumber: 1,
      background: null,
      elements: slideElements,
      showMasterSp,
    } satisfies Slide,
    layoutElements,
    layoutShowMasterSp,
    masterElements,
  };
}

function shape(
  id: string,
  options: { placeholderType?: string; text?: string } = {},
): ShapeElement {
  const textBody =
    options.text === undefined
      ? null
      : {
          paragraphs: [{ runs: [{ text: options.text }] }],
        };

  return unsafeTypeAssertion<ShapeElement>({
    type: "shape",
    id,
    placeholderType: options.placeholderType,
    textBody,
  });
}
