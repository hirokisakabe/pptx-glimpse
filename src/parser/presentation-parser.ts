import type { SlideSize } from "../model/presentation.js";
import type { DefaultTextStyle } from "../model/text.js";
import { parseXml, type XmlNode } from "./xml-parser.js";
import { parseListStyle } from "./text-style-parser.js";
import { debug } from "../warning-logger.js";

export interface PresentationInfo {
  slideSize: SlideSize;
  slideRIds: string[];
  defaultTextStyle?: DefaultTextStyle;
}

const DEFAULT_SLIDE_WIDTH = 9144000;
const DEFAULT_SLIDE_HEIGHT = 5143500;

export function parsePresentation(xml: string): PresentationInfo {
  const parsed = parseXml(xml);
  const pres = parsed.presentation as XmlNode | undefined;

  if (!pres) {
    debug("presentation.missing", `missing root element "presentation" in XML`);
    return {
      slideSize: { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT },
      slideRIds: [],
    };
  }

  const sldSz = pres.sldSz as XmlNode | undefined;
  let slideSize: SlideSize;
  if (!sldSz || sldSz["@_cx"] === undefined || sldSz["@_cy"] === undefined) {
    debug(
      "presentation.sldSz",
      `sldSz missing, using default ${DEFAULT_SLIDE_WIDTH}x${DEFAULT_SLIDE_HEIGHT} EMU`,
    );
    slideSize = { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT };
  } else {
    slideSize = {
      width: Number(sldSz["@_cx"]),
      height: Number(sldSz["@_cy"]),
    };
  }

  const sldIdLst = pres.sldIdLst as XmlNode | undefined;
  const sldIdArr = (sldIdLst?.sldId as XmlNode[] | undefined) ?? [];
  const slideRIds: string[] = sldIdArr
    .map((s) => (s["@_r:id"] as string | undefined) ?? (s["@_id"] as string | undefined))
    .filter((id): id is string => {
      if (id === undefined) {
        debug("presentation.slideRId", "undefined slide relationship ID found, skipping");
        return false;
      }
      return true;
    });

  const defaultTextStyle = parseListStyle(pres.defaultTextStyle as XmlNode);

  return { slideSize, slideRIds, defaultTextStyle };
}
