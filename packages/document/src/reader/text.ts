/**
 * Read `p:txBody` into text body / paragraph / run of PptxSourceModel source.
 *
 * The current typed subset is plain run text and basic run properties (bold/italic/
 * Underline / font size / typeface / solid color). bullet / field /
 * Unsupported nodes such as line breaks are retained as raw sidecars.
 */

import type {
  PartPath,
  RawSidecarId,
  SourceAutoNumScheme,
  SourceNodeId,
  SourceParagraph,
  SourceParagraphProperties,
  SourceRunProperties,
  SourceSpacingValue,
  SourceTextAlign,
  SourceTextBody,
  SourceTextBodyProperties,
  SourceTextRun,
  SourceTextStyle,
  SourceTextVerticalType,
  SourceTextWrap,
  SourceUnderlineStyle,
  SourceVerticalAnchor,
} from "../source/index.js";
import { asEmu, asHundredthPt, asPt, asSourceNodeId } from "../source/index.js";
import { unsafeOoxmlBoundaryAssertion } from "../unsafe-type-assertion.js";
import { parseColorElement } from "./drawing.js";
import { isTrue, numericAttr } from "./drawing.js";
import { parseEnumValue, parseEnumValueWithDefault } from "./ooxml-values.js";
import { collectUnknownSidecars } from "./raw-node.js";
import {
  getAttr,
  getChild,
  getChildArray,
  getChildText,
  localName,
  type XmlNode,
  type XmlOrderedNode,
} from "./xml.js";

const KNOWN_TXBODY_CHILDREN: ReadonlySet<string> = new Set(["bodyPr", "lstStyle", "p"]);
const KNOWN_PARAGRAPH_CHILDREN: ReadonlySet<string> = new Set([
  "pPr",
  "r",
  "fld",
  "br",
  "endParaRPr",
]);
const KNOWN_RUN_CHILDREN: ReadonlySet<string> = new Set(["rPr", "t"]);
const KNOWN_RUN_PROPERTY_CHILDREN: ReadonlySet<string> = new Set([
  "latin",
  "ea",
  "cs",
  "solidFill",
  "highlight",
  "uFill",
]);
const UNDERLINE_STYLES: ReadonlySet<SourceUnderlineStyle> = new Set([
  "sng",
  "dbl",
  "heavy",
  "dotted",
  "dottedHeavy",
  "dash",
  "dashHeavy",
  "dashLong",
  "dashLongHeavy",
  "dotDash",
  "dotDashHeavy",
  "dotDotDash",
  "dotDotDashHeavy",
  "wavy",
  "wavyHeavy",
  "wavyDbl",
  "none",
]);

const ALIGN_MAP: Readonly<Record<string, SourceTextAlign>> = {
  l: "left",
  ctr: "center",
  r: "right",
  just: "justify",
};

const ANCHOR_MAP: Readonly<Record<string, SourceVerticalAnchor>> = {
  t: "top",
  ctr: "middle",
  b: "bottom",
};

const WRAP_VALUES: ReadonlySet<SourceTextWrap> = new Set(["square", "none"]);
const VERTICAL_VALUES: ReadonlySet<SourceTextVerticalType> = new Set([
  "horz",
  "vert",
  "vert270",
  "eaVert",
  "wordArtVert",
  "mongolianVert",
]);

const VALID_AUTO_NUM_SCHEMES: ReadonlySet<SourceAutoNumScheme> = new Set([
  "arabicPeriod",
  "arabicParenR",
  "romanUcPeriod",
  "romanLcPeriod",
  "alphaUcPeriod",
  "alphaLcPeriod",
  "alphaLcParenR",
  "alphaUcParenR",
  "arabicPlain",
]);

/** Read `p:txBody`. Undefined for shapes without `txBody`. */
export function parseTextBody(
  txBody: XmlNode | undefined,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  ownerNodeId: SourceNodeId | undefined,
  ownerOrderingSlot: number,
  orderedTxBody?: readonly XmlOrderedNode[],
): SourceTextBody | undefined {
  if (!txBody) return undefined;

  const properties = parseBodyProperties(getChild(txBody, "bodyPr"));
  const listStyle = parseTextStyle(getChild(txBody, "lstStyle"));
  const orderedParagraphs = orderedTxBody
    ?.filter((child) => orderedLocalName(child) === "p")
    .map((child) => {
      const key = orderedKey(child);
      const value = key !== undefined ? child[key] : undefined;
      return Array.isArray(value) ? (value as readonly XmlOrderedNode[]) : undefined;
    });
  const paragraphs: SourceParagraph[] = [];
  let logicalParagraphIndex = 0;
  getChildArray(txBody, "p").forEach((p, paragraphIndex) => {
    const orderedChildren = orderedParagraphs?.[paragraphIndex];
    if (orderedChildren !== undefined && hasMultipleBulletPPr(p, orderedChildren)) {
      const split = splitInterleavedParagraph(
        p,
        orderedChildren,
        partPath,
        nextId,
        ownerNodeId,
        ownerOrderingSlot,
        logicalParagraphIndex,
      );
      paragraphs.push(...split);
      logicalParagraphIndex += split.length;
      return;
    }
    paragraphs.push(
      parseParagraph(
        p,
        partPath,
        nextId,
        ownerNodeId,
        ownerOrderingSlot,
        logicalParagraphIndex,
        orderedChildren,
      ),
    );
    logicalParagraphIndex++;
  });
  const rawSidecars = collectUnknownSidecars(txBody, KNOWN_TXBODY_CHILDREN, nextId);

  return {
    paragraphs,
    ...(properties !== undefined ? { properties } : {}),
    ...(listStyle !== undefined ? { listStyle } : {}),
    handle: { partPath },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

export function parseTextStyle(node: XmlNode | undefined): SourceTextStyle | undefined {
  if (node === undefined) return undefined;
  const defaultParagraph = parseParagraphProperties(getChild(node, "defPPr"));
  const levels = Array.from({ length: 9 }, (_, index) =>
    parseParagraphProperties(getChild(node, `lvl${index + 1}pPr`)),
  );
  if (defaultParagraph === undefined && levels.every((level) => level === undefined)) {
    return undefined;
  }
  return { ...(defaultParagraph !== undefined ? { defaultParagraph } : {}), levels };
}

function hasMultipleBulletPPr(p: XmlNode, orderedChildren: readonly XmlOrderedNode[]): boolean {
  const pPrList = getChildArray(p, "pPr");
  let bulletPPrCount = 0;
  let pPrCounter = 0;
  for (const child of orderedChildren) {
    if (orderedLocalName(child) !== "pPr") continue;
    const pPr = pPrList[pPrCounter];
    if (
      pPr !== undefined &&
      (getChild(pPr, "buChar") !== undefined || getChild(pPr, "buAutoNum") !== undefined)
    ) {
      bulletPPrCount++;
      if (bulletPPrCount >= 2) return true;
    }
    pPrCounter++;
  }
  return false;
}

function splitInterleavedParagraph(
  p: XmlNode,
  orderedChildren: readonly XmlOrderedNode[],
  partPath: PartPath,
  nextId: () => RawSidecarId,
  ownerNodeId: SourceNodeId | undefined,
  ownerOrderingSlot: number,
  paragraphIndex: number,
): SourceParagraph[] {
  const pPrList = getChildArray(p, "pPr");
  const rList = getChildArray(p, "r");
  const fldList = getChildArray(p, "fld");
  const brList = getChildArray(p, "br");
  const groups: {
    pPrIndex: number;
    rNodes: XmlNode[];
    fldNodes: XmlNode[];
    brNodes: XmlNode[];
    orderedChildren: XmlOrderedNode[];
  }[] = [];
  let currentGroup:
    | {
        pPrIndex: number;
        rNodes: XmlNode[];
        fldNodes: XmlNode[];
        brNodes: XmlNode[];
        orderedChildren: XmlOrderedNode[];
      }
    | undefined;
  let pPrCounter = 0;
  let runCounter = 0;
  let fieldCounter = 0;
  let breakCounter = 0;

  for (const child of orderedChildren) {
    const tag = orderedLocalName(child);
    if (tag === "pPr") {
      const pPr = pPrList[pPrCounter];
      const hasBullet =
        pPr !== undefined &&
        (getChild(pPr, "buChar") !== undefined || getChild(pPr, "buAutoNum") !== undefined);
      if (hasBullet || currentGroup === undefined) {
        if (currentGroup !== undefined) groups.push(currentGroup);
        currentGroup = {
          pPrIndex: pPrCounter,
          rNodes: [],
          fldNodes: [],
          brNodes: [],
          orderedChildren: [child],
        };
      } else {
        currentGroup.orderedChildren.push(child);
      }
      pPrCounter++;
    } else if (tag === "r") {
      currentGroup ??= {
        pPrIndex: -1,
        rNodes: [],
        fldNodes: [],
        brNodes: [],
        orderedChildren: [],
      };
      currentGroup.orderedChildren.push(child);
      if (rList[runCounter] !== undefined) currentGroup.rNodes.push(rList[runCounter]);
      runCounter++;
    } else if (tag === "fld") {
      currentGroup ??= {
        pPrIndex: -1,
        rNodes: [],
        fldNodes: [],
        brNodes: [],
        orderedChildren: [],
      };
      currentGroup.orderedChildren.push(child);
      if (fldList[fieldCounter] !== undefined) currentGroup.fldNodes.push(fldList[fieldCounter]);
      fieldCounter++;
    } else if (tag === "br") {
      currentGroup ??= {
        pPrIndex: -1,
        rNodes: [],
        fldNodes: [],
        brNodes: [],
        orderedChildren: [],
      };
      currentGroup.orderedChildren.push(child);
      if (brList[breakCounter] !== undefined) currentGroup.brNodes.push(brList[breakCounter]);
      breakCounter++;
    }
  }
  if (currentGroup !== undefined) groups.push(currentGroup);

  return groups.map((group, groupIndex) => {
    const synthetic: XmlNode = {};
    if (group.pPrIndex >= 0) synthetic.pPr = pPrList[group.pPrIndex];
    synthetic.r = group.rNodes;
    synthetic.fld = group.fldNodes;
    synthetic.br = group.brNodes;
    if (groupIndex === groups.length - 1 && getChild(p, "endParaRPr") !== undefined) {
      synthetic.endParaRPr = getChild(p, "endParaRPr");
    }
    const paragraph = parseParagraph(
      synthetic,
      partPath,
      nextId,
      ownerNodeId,
      ownerOrderingSlot,
      paragraphIndex + groupIndex,
      group.orderedChildren,
    );
    if (groupIndex < groups.length - 1 && paragraph.runs.length > 0) {
      const lastRun = paragraph.runs[paragraph.runs.length - 1];
      if (lastRun.text.endsWith("\n")) {
        const trimmed = lastRun.text.slice(0, -1);
        const precedingRuns = paragraph.runs.slice(0, -1);
        return {
          ...paragraph,
          runs: trimmed === "" ? precedingRuns : [...precedingRuns, { ...lastRun, text: trimmed }],
        };
      }
    }
    return paragraph;
  });
}

function parseBodyProperties(bodyPr: XmlNode | undefined): SourceTextBodyProperties | undefined {
  if (!bodyPr) return undefined;

  const marginLeft = numericAttr(bodyPr, "lIns");
  const marginRight = numericAttr(bodyPr, "rIns");
  const marginTop = numericAttr(bodyPr, "tIns");
  const marginBottom = numericAttr(bodyPr, "bIns");
  const anchorToken = getAttr(bodyPr, "anchor");
  const anchor = anchorToken !== undefined ? ANCHOR_MAP[anchorToken] : undefined;
  const wrap = parseWrap(getAttr(bodyPr, "wrap"));
  const vert = parseVerticalType(getAttr(bodyPr, "vert"));
  const numCol = numericAttr(bodyPr, "numCol");
  const normAutofit = getChild(bodyPr, "normAutofit");
  const hasSpAutofit = getChild(bodyPr, "spAutoFit") !== undefined;
  const hasNoAutofit = getChild(bodyPr, "noAutofit") !== undefined;
  const fontScale = numericAttr(normAutofit, "fontScale");
  const lnSpcReduction = numericAttr(normAutofit, "lnSpcReduction");

  const properties: SourceTextBodyProperties = {
    ...(marginLeft !== undefined ? { marginLeft: emu(marginLeft) } : {}),
    ...(marginRight !== undefined ? { marginRight: emu(marginRight) } : {}),
    ...(marginTop !== undefined ? { marginTop: emu(marginTop) } : {}),
    ...(marginBottom !== undefined ? { marginBottom: emu(marginBottom) } : {}),
    ...(anchor !== undefined ? { anchor } : {}),
    ...(wrap !== undefined ? { wrap } : {}),
    ...(normAutofit !== undefined
      ? {
          autoFit: "normAutofit",
          fontScale: fontScale !== undefined ? fontScale / 100000 : 1,
          lnSpcReduction: lnSpcReduction !== undefined ? lnSpcReduction / 100000 : 0,
        }
      : hasSpAutofit
        ? { autoFit: "spAutofit", fontScale: 1, lnSpcReduction: 0 }
        : hasNoAutofit
          ? { autoFit: "noAutofit", fontScale: 1, lnSpcReduction: 0 }
          : {}),
    ...(numCol !== undefined ? { numCol: Math.max(1, numCol) } : {}),
    ...(vert !== undefined ? { vert } : {}),
  };
  return Object.keys(properties).length > 0 ? properties : undefined;
}

function parseWrap(value: string | undefined): SourceTextWrap | undefined {
  return parseEnumValue(value, WRAP_VALUES);
}

function parseVerticalType(value: string | undefined): SourceTextVerticalType | undefined {
  return parseEnumValue(value, VERTICAL_VALUES);
}

function parseParagraph(
  p: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  ownerNodeId: SourceNodeId | undefined,
  ownerOrderingSlot: number,
  paragraphIndex: number,
  orderedChildren?: readonly XmlOrderedNode[],
): SourceParagraph {
  const properties = parseParagraphProperties(getChild(p, "pPr"));
  const runs = parseRunsInOrder(
    p,
    partPath,
    nextId,
    ownerNodeId,
    ownerOrderingSlot,
    paragraphIndex,
    orderedChildren,
  );
  const rawSidecars = collectUnknownSidecars(p, KNOWN_PARAGRAPH_CHILDREN, nextId);

  return {
    runs,
    ...(properties !== undefined ? { properties } : {}),
    handle: {
      partPath,
      nodeId: textNodeId("paragraph", ownerNodeId, ownerOrderingSlot, paragraphIndex),
      orderingSlot: paragraphIndex,
    },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseRunsInOrder(
  p: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  ownerNodeId: SourceNodeId | undefined,
  ownerOrderingSlot: number,
  paragraphIndex: number,
  orderedChildren?: readonly XmlOrderedNode[],
): SourceTextRun[] {
  const rList = getChildArray(p, "r");
  const fldList = getChildArray(p, "fld");
  const brList = getChildArray(p, "br");
  if (orderedChildren === undefined) {
    return [
      ...rList.map((r, runIndex) =>
        parseRun(r, partPath, nextId, ownerNodeId, ownerOrderingSlot, paragraphIndex, runIndex),
      ),
      ...fldList.map((fld, index) =>
        parseRun(
          fld,
          partPath,
          nextId,
          ownerNodeId,
          ownerOrderingSlot,
          paragraphIndex,
          rList.length + index,
        ),
      ),
      ...brList.map((br, index) =>
        parseBreakRun(
          br,
          partPath,
          nextId,
          ownerNodeId,
          ownerOrderingSlot,
          paragraphIndex,
          rList.length + fldList.length + index,
        ),
      ),
    ];
  }

  const counters: Record<string, number> = {};
  const runs: SourceTextRun[] = [];
  for (const child of orderedChildren) {
    const tag = orderedLocalName(child);
    if (tag !== "r" && tag !== "fld" && tag !== "br") continue;
    const index = counters[tag] ?? 0;
    counters[tag] = index + 1;
    const runIndex = runs.length;
    if (tag === "r" && rList[index] !== undefined) {
      runs.push(
        parseRun(
          rList[index],
          partPath,
          nextId,
          ownerNodeId,
          ownerOrderingSlot,
          paragraphIndex,
          runIndex,
        ),
      );
    } else if (tag === "fld" && fldList[index] !== undefined) {
      runs.push(
        parseRun(
          fldList[index],
          partPath,
          nextId,
          ownerNodeId,
          ownerOrderingSlot,
          paragraphIndex,
          runIndex,
        ),
      );
    } else if (tag === "br" && brList[index] !== undefined) {
      runs.push(
        parseBreakRun(
          brList[index],
          partPath,
          nextId,
          ownerNodeId,
          ownerOrderingSlot,
          paragraphIndex,
          runIndex,
        ),
      );
    }
  }
  return runs;
}

function parseParagraphProperties(pPr: XmlNode | undefined): SourceParagraphProperties | undefined {
  if (!pPr) return undefined;

  const alignToken = getAttr(pPr, "algn");
  const align = alignToken !== undefined ? ALIGN_MAP[alignToken] : undefined;
  const level = numericAttr(pPr, "lvl");
  const lineSpacing = parseSpacing(getChild(pPr, "lnSpc"));
  const spaceBefore = parseSpacing(getChild(pPr, "spcBef"));
  const spaceAfter = parseSpacing(getChild(pPr, "spcAft"));
  const marginLeft = numericAttr(pPr, "marL");
  const indent = numericAttr(pPr, "indent");
  const bullet = parseBullet(pPr);
  const bulletFont = getAttr(getChild(pPr, "buFont"), "typeface");
  const bulletColor = parseColorElement(getChild(pPr, "buClr"));
  const bulletSizePct = numericAttr(getChild(pPr, "buSzPct"), "val");
  const tabStops = parseTabStops(pPr);
  const defaultRunProperties = parseRunProperties(getChild(pPr, "defRPr"));

  const properties: SourceParagraphProperties = {
    ...(align !== undefined ? { align } : {}),
    ...(level !== undefined ? { level } : {}),
    ...(lineSpacing !== undefined ? { lineSpacing } : {}),
    ...(spaceBefore !== undefined ? { spaceBefore } : {}),
    ...(spaceAfter !== undefined ? { spaceAfter } : {}),
    ...(marginLeft !== undefined ? { marginLeft: emu(marginLeft) } : {}),
    ...(indent !== undefined ? { indent: emu(indent) } : {}),
    ...(bullet !== undefined ? { bullet } : {}),
    ...(bulletFont !== undefined ? { bulletFont } : {}),
    ...(bulletColor !== undefined ? { bulletColor } : {}),
    ...(bulletSizePct !== undefined ? { bulletSizePct } : {}),
    ...(tabStops.length > 0 ? { tabStops } : {}),
    ...(defaultRunProperties !== undefined ? { defaultRunProperties } : {}),
  };
  return Object.keys(properties).length > 0 ? properties : undefined;
}

function parseTabStops(
  pPr: XmlNode | undefined,
): NonNullable<SourceParagraphProperties["tabStops"]> {
  return getChildArray(getChild(pPr, "tabLst"), "tab").map((tab) => ({
    position: emu(numericAttr(tab, "pos") ?? 0),
    alignment:
      unsafeOoxmlBoundaryAssertion<"l" | "ctr" | "r" | "dec" | undefined>(getAttr(tab, "algn")) ??
      "l",
  }));
}

function parseSpacing(node: XmlNode | undefined): SourceSpacingValue | undefined {
  const points = numericAttr(getChild(node, "spcPts"), "val");
  if (points !== undefined) return { type: "pts", value: asHundredthPt(points) };
  const percent = numericAttr(getChild(node, "spcPct"), "val");
  if (percent !== undefined) return { type: "pct", value: percent };
  return undefined;
}

function parseBullet(pPr: XmlNode | undefined): SourceParagraphProperties["bullet"] | undefined {
  if (pPr === undefined) return undefined;
  if (getChild(pPr, "buNone") !== undefined) return { type: "none" };
  const buChar = getChild(pPr, "buChar");
  if (buChar !== undefined) {
    return { type: "char", char: decodeXmlCharRef(getAttr(buChar, "char") ?? "\u2022") };
  }
  const buAutoNum = getChild(pPr, "buAutoNum");
  if (buAutoNum !== undefined) {
    const scheme = getAttr(buAutoNum, "type") ?? "arabicPeriod";
    return {
      type: "autoNum",
      scheme: parseEnumValueWithDefault(scheme, VALID_AUTO_NUM_SCHEMES, "arabicPeriod"),
      startAt: numericAttr(buAutoNum, "startAt") ?? 1,
    };
  }
  return undefined;
}

function decodeXmlCharRef(value: string): string {
  return value.replace(
    /&#x([0-9a-fA-F]+);|&#([0-9]+);/g,
    (_match, hex?: string, decimal?: string) =>
      String.fromCodePoint(parseInt(hex ?? decimal ?? "0", hex !== undefined ? 16 : 10)),
  );
}

function parseRun(
  r: XmlNode,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  ownerNodeId: SourceNodeId | undefined,
  ownerOrderingSlot: number,
  paragraphIndex: number,
  runIndex: number,
): SourceTextRun {
  const properties = parseRunProperties(getChild(r, "rPr"));
  const rawSidecars = collectRunSidecars(r, nextId);

  return {
    kind: "textRun",
    text: decodeXmlCharRef(getChildText(r, "t") ?? ""),
    ...(properties !== undefined ? { properties } : {}),
    handle: {
      partPath,
      nodeId: textNodeId("run", ownerNodeId, ownerOrderingSlot, paragraphIndex, runIndex),
      orderingSlot: runIndex,
    },
    ...(rawSidecars.length > 0 ? { rawSidecars } : {}),
  };
}

function parseBreakRun(
  br: XmlNode | undefined,
  partPath: PartPath,
  nextId: () => RawSidecarId,
  ownerNodeId: SourceNodeId | undefined,
  ownerOrderingSlot: number,
  paragraphIndex: number,
  runIndex: number,
): SourceTextRun {
  const properties = parseRunProperties(getChild(br, "rPr"));
  return {
    kind: "textRun",
    text: "\n",
    ...(properties !== undefined ? { properties } : {}),
    handle: {
      partPath,
      nodeId: textNodeId("run", ownerNodeId, ownerOrderingSlot, paragraphIndex, runIndex),
      orderingSlot: runIndex,
    },
  };
}

function parseRunProperties(rPr: XmlNode | undefined): SourceRunProperties | undefined {
  if (!rPr) return undefined;

  const bold = getAttr(rPr, "b");
  const italic = getAttr(rPr, "i");
  const underline = getAttr(rPr, "u");
  const underlineStyle = parseEnumValue(underline, UNDERLINE_STYLES);
  const strike = getAttr(rPr, "strike");
  const baseline = numericAttr(rPr, "baseline");
  const size = numericAttr(rPr, "sz");
  const typeface = getAttr(getChild(rPr, "latin"), "typeface");
  const typefaceEa = getAttr(getChild(rPr, "ea"), "typeface");
  const typefaceCs = getAttr(getChild(rPr, "cs"), "typeface");
  const hasHyperlink = getChild(rPr, "hlinkClick") !== undefined;
  const color = parseColorElement(getChild(rPr, "solidFill"));
  const underlineColor = parseColorElement(getChild(getChild(rPr, "uFill"), "solidFill"));
  const highlight = parseColorElement(getChild(rPr, "highlight"));

  const properties: SourceRunProperties = {
    ...(bold !== undefined ? { bold: isTrue(bold) } : {}),
    ...(italic !== undefined ? { italic: isTrue(italic) } : {}),
    ...(underline !== undefined
      ? {
          underline: underline !== "none",
          ...(underlineStyle !== undefined ? { underlineStyle } : {}),
        }
      : hasHyperlink
        ? { underline: true }
        : {}),
    ...(underlineColor !== undefined ? { underlineColor } : {}),
    ...(strike !== undefined ? { strikethrough: strike !== "noStrike" } : {}),
    ...(baseline !== undefined ? { baseline: baseline / 1000 } : {}),
    ...(highlight !== undefined ? { highlight } : {}),
    // `a:rPr@sz` is in 1/100 pt unit. Convert to pt and save.
    ...(size !== undefined ? { fontSize: asPt(size / 100) } : {}),
    ...(typeface !== undefined ? { typeface } : {}),
    ...(typefaceEa !== undefined ? { typefaceEa } : {}),
    ...(typefaceCs !== undefined ? { typefaceCs } : {}),
    ...(color !== undefined
      ? { color }
      : hasHyperlink
        ? { color: { kind: "scheme", scheme: "hlink" } }
        : {}),
  };
  return Object.keys(properties).length > 0 ? properties : undefined;
}

/**
 * Collect unsupported run materials in sidecar. In addition to directly under run (other than `a:rPr` / `a:t`),
 * Unsupported child elements in `a:rPr` (`a:hlinkClick` / `a:ln`, etc.) are also retained.
 */
function collectRunSidecars(
  r: XmlNode,
  nextId: () => RawSidecarId,
): ReturnType<typeof collectUnknownSidecars> {
  const runLevel = collectUnknownSidecars(r, KNOWN_RUN_CHILDREN, nextId);
  const propLevel = collectUnknownSidecars(getChild(r, "rPr"), KNOWN_RUN_PROPERTY_CHILDREN, nextId);
  return [...runLevel, ...propLevel];
}

function emu(value: number) {
  return asEmu(value);
}

function textNodeId(
  kind: "paragraph" | "run",
  ownerNodeId: SourceNodeId | undefined,
  ownerOrderingSlot: number,
  paragraphIndex: number,
  runIndex?: number,
): SourceNodeId {
  const owner =
    ownerNodeId !== undefined ? `shape:${ownerNodeId}` : `shapeSlot:${ownerOrderingSlot}`;
  const suffix = kind === "paragraph" ? `p:${paragraphIndex}` : `p:${paragraphIndex}:r:${runIndex}`;
  return asSourceNodeId(`text:${owner}:${suffix}`);
}

function orderedKey(node: XmlOrderedNode): string | undefined {
  return Object.keys(node).find((key) => key !== ":@");
}

function orderedLocalName(node: XmlOrderedNode): string | undefined {
  const key = orderedKey(node);
  return key !== undefined ? localName(key) : undefined;
}
