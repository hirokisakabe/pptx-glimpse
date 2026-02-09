import type { SlideSize } from "../model/presentation.js";
import { parseXml } from "./xml-parser.js";

export interface PresentationInfo {
  slideSize: SlideSize;
  slideRIds: string[];
}

export function parsePresentation(xml: string): PresentationInfo {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parsed = parseXml(xml) as any;
  const pres = parsed.presentation;

  const sldSz = pres.sldSz;
  const slideSize: SlideSize = {
    width: Number(sldSz["@_cx"]),
    height: Number(sldSz["@_cy"]),
  };

  const sldIdLst = pres.sldIdLst?.sldId ?? [];
  const slideRIds = sldIdLst.map((s: Record<string, string>) => s["@_r:id"] ?? s["@_id"]);

  return { slideSize, slideRIds };
}
