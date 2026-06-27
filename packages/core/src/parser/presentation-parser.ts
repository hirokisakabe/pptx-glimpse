import type { EmbeddedFont, Protection, SlideSize } from "@pptx-glimpse/renderer";
import type { DefaultTextStyle } from "@pptx-glimpse/renderer";
import { asEmu } from "@pptx-glimpse/renderer";
import { debug } from "@pptx-glimpse/renderer";

import { unsafeTypeAssertion } from "../unsafe-type-assertion.js";
import { parseListStyle } from "./text-style-parser.js";
import { parseXml, type XmlNode } from "./xml-parser.js";

export interface PresentationInfo {
  slideSize: SlideSize;
  slideRIds: string[];
  defaultTextStyle?: DefaultTextStyle;
  embeddedFonts?: EmbeddedFont[];
  protection?: Protection;
}

const DEFAULT_SLIDE_WIDTH = asEmu(9144000);
const DEFAULT_SLIDE_HEIGHT = asEmu(5143500);

export function parsePresentation(xml: string): PresentationInfo {
  const parsed = parseXml(xml);
  const pres = unsafeTypeAssertion<XmlNode | undefined>(parsed.presentation);

  if (!pres) {
    debug("presentation.missing", `missing root element "presentation" in XML`);
    return {
      slideSize: { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT },
      slideRIds: [],
    };
  }

  const sldSz = unsafeTypeAssertion<XmlNode | undefined>(pres.sldSz);
  let slideSize: SlideSize;
  if (!sldSz || sldSz["@_cx"] === undefined || sldSz["@_cy"] === undefined) {
    debug(
      "presentation.sldSz",
      `sldSz missing, using default ${DEFAULT_SLIDE_WIDTH}x${DEFAULT_SLIDE_HEIGHT} EMU`,
    );
    slideSize = { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT };
  } else {
    slideSize = {
      width: asEmu(Number(sldSz["@_cx"])),
      height: asEmu(Number(sldSz["@_cy"])),
    };
  }

  const sldIdLst = unsafeTypeAssertion<XmlNode | undefined>(pres.sldIdLst);
  const sldIdArr = unsafeTypeAssertion<XmlNode[] | undefined>(sldIdLst?.sldId) ?? [];
  const slideRIds: string[] = sldIdArr
    .map(
      (s) =>
        unsafeTypeAssertion<string | undefined>(s["@_r:id"]) ??
        unsafeTypeAssertion<string | undefined>(s["@_id"]),
    )
    .filter((id): id is string => {
      if (id === undefined) {
        debug("presentation.slideRId", "undefined slide relationship ID found, skipping");
        return false;
      }
      return true;
    });

  const defaultTextStyle = parseListStyle(unsafeTypeAssertion<XmlNode>(pres.defaultTextStyle));
  const embeddedFonts = parseEmbeddedFontList(
    unsafeTypeAssertion<XmlNode | undefined>(pres.embeddedFontLst),
  );
  const protection = parseProtection(unsafeTypeAssertion<XmlNode | undefined>(pres.modifyVerifier));

  return {
    slideSize,
    slideRIds,
    defaultTextStyle,
    ...(embeddedFonts && { embeddedFonts }),
    ...(protection && { protection }),
  };
}

function parseEmbeddedFontList(node: XmlNode | undefined): EmbeddedFont[] | undefined {
  if (!node) return undefined;
  const fonts = node.embeddedFont;
  if (!fonts) return undefined;
  const fontArr = unsafeTypeAssertion<XmlNode[]>(Array.isArray(fonts) ? fonts : [fonts]);
  const result: EmbeddedFont[] = [];
  for (const f of fontArr) {
    const fontRaw = f.font;
    const fontNode = unsafeTypeAssertion<XmlNode | undefined>(
      Array.isArray(fontRaw) ? fontRaw[0] : fontRaw,
    );
    if (!fontNode) continue;
    const font: EmbeddedFont = {
      typeface: unsafeTypeAssertion<string | undefined>(fontNode["@_typeface"]) ?? "",
    };
    if (fontNode["@_panose"]) font.panose = unsafeTypeAssertion<string>(fontNode["@_panose"]);
    if (fontNode["@_pitchFamily"] !== undefined)
      font.pitchFamily = Number(fontNode["@_pitchFamily"]);
    if (fontNode["@_charset"] !== undefined) font.charset = Number(fontNode["@_charset"]);
    result.push(font);
  }
  return result.length > 0 ? result : undefined;
}

function parseProtection(node: XmlNode | undefined): Protection | undefined {
  if (!node) return undefined;
  const verifier: NonNullable<Protection["modifyVerifier"]> = {};
  if (node["@_algorithmName"])
    verifier.algorithmName = unsafeTypeAssertion<string>(node["@_algorithmName"]);
  if (node["@_hashValue"]) verifier.hashValue = unsafeTypeAssertion<string>(node["@_hashValue"]);
  if (node["@_saltValue"]) verifier.saltValue = unsafeTypeAssertion<string>(node["@_saltValue"]);
  if (node["@_spinCount"] !== undefined) verifier.spinCount = Number(node["@_spinCount"]);
  return { modifyVerifier: verifier };
}
