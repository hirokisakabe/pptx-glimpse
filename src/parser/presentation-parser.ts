import type { SlideSize } from "../model/presentation.js";
import type { DefaultTextStyle } from "../model/text.js";
import { parseXml } from "./xml-parser.js";
import { parseListStyle } from "./text-style-parser.js";

export interface PresentationInfo {
  slideSize: SlideSize;
  slideRIds: string[];
  defaultTextStyle?: DefaultTextStyle;
}

const WARN_PREFIX = "[pptx-glimpse]";
const DEFAULT_SLIDE_WIDTH = 9144000;
const DEFAULT_SLIDE_HEIGHT = 5143500;

export function parsePresentation(xml: string): PresentationInfo {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;
  const pres = parsed.presentation;

  if (!pres) {
    console.warn(`${WARN_PREFIX} Presentation: missing root element "presentation" in XML`);
    return {
      slideSize: { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT },
      slideRIds: [],
    };
  }

  const sldSz = pres.sldSz;
  let slideSize: SlideSize;
  if (!sldSz || sldSz["@_cx"] === undefined || sldSz["@_cy"] === undefined) {
    console.warn(
      `${WARN_PREFIX} Presentation: sldSz missing, using default ${DEFAULT_SLIDE_WIDTH}x${DEFAULT_SLIDE_HEIGHT} EMU`,
    );
    slideSize = { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT };
  } else {
    slideSize = {
      width: Number(sldSz["@_cx"]),
      height: Number(sldSz["@_cy"]),
    };
  }

  const sldIdLst = pres.sldIdLst?.sldId ?? [];
  const slideRIds: string[] = sldIdLst
    .map((s: Record<string, string>) => s["@_r:id"] ?? s["@_id"])
    .filter((id: string | undefined) => {
      if (id === undefined) {
        console.warn(
          `${WARN_PREFIX} Presentation: undefined slide relationship ID found, skipping`,
        );
        return false;
      }
      return true;
    });

  const defaultTextStyle = parseListStyle(pres.defaultTextStyle);

  return { slideSize, slideRIds, defaultTextStyle };
}
