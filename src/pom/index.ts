import type { SlideElement } from "../model/shape.js";
import { clearXmlCache, enableXmlCache } from "../parser/xml-parser.js";
import { parsePptxData, parseSlideWithLayout } from "../pptx-data-parser.js";
import type { LogLevel } from "../warning-logger.js";
import { flushWarnings, initWarningLogger } from "../warning-logger.js";
import { convertSlideToPom } from "./pom-converter.js";
import { pomLayerToXml } from "./xml-builder.js";

export interface ConvertPptxToPomOptions {
  /** 変換対象のスライド番号 (1始まり)。未指定で全スライド */
  slides?: number[];
  /** 警告ログレベル。デフォルト: "off" */
  logLevel?: LogLevel;
}

export interface PomSlide {
  slideNumber: number;
  xml: string;
}

export function convertPptxToPom(
  input: Buffer | Uint8Array,
  options?: ConvertPptxToPomOptions,
): PomSlide[] {
  enableXmlCache();
  try {
    initWarningLogger(options?.logLevel ?? "off");

    const data = parsePptxData(input);

    const targetSlides = options?.slides
      ? data.slidePaths.filter((s) => options.slides!.includes(s.slideNumber))
      : data.slidePaths;

    const results: PomSlide[] = [];
    for (const { slideNumber, path } of targetSlides) {
      const parsed = parseSlideWithLayout(slideNumber, path, data);
      if (!parsed) continue;

      const { slide, layoutElements, layoutShowMasterSp } = parsed;

      // Merge shapes: master (back) → layout → slide (front)
      const effectiveMasterElements =
        slide.showMasterSp && layoutShowMasterSp ? data.masterElements : [];
      slide.elements = mergeElements(effectiveMasterElements, layoutElements, slide.elements);

      const node = convertSlideToPom(slide, data.presInfo.slideSize);
      const xml = pomLayerToXml(node);
      results.push({ slideNumber, xml });
    }

    flushWarnings();

    return results;
  } finally {
    clearXmlCache();
  }
}

function mergeElements(
  masterElements: SlideElement[],
  layoutElements: SlideElement[],
  slideElements: SlideElement[],
): SlideElement[] {
  const filterPlaceholders = (elements: SlideElement[]) =>
    elements.filter((el) => {
      if (el.type !== "shape") return true;
      return !el.placeholderType;
    });

  const filteredMaster = filterPlaceholders(masterElements);
  const filteredLayout = filterPlaceholders(layoutElements);

  return [...filteredMaster, ...filteredLayout, ...slideElements];
}
