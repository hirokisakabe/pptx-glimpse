import type { AutoNumScheme, BulletType, Paragraph, TextBody } from "../model/text.js";
import type { SlideScaleContext } from "./pom-converter.js";
import { stripHash } from "./pom-converter.js";
import type {
  PomAlignText,
  PomLiNode,
  PomNode,
  PomOlNode,
  PomTextNode,
  PomUlNode,
  PomVStackNode,
} from "./pom-types.js";

/**
 * Convert a TextBody into one or more POM nodes.
 * - Plain text paragraphs → TextNode (or VStack of TextNodes if mixed alignment/style)
 * - Bullet char paragraphs → UlNode
 * - Auto-numbered paragraphs → OlNode
 */
export function convertTextBodyToNodes(textBody: TextBody, _ctx: SlideScaleContext): PomNode {
  const paragraphs = textBody.paragraphs.filter(
    (p) => p.runs.length > 0 && p.runs.some((r) => r.text.length > 0),
  );

  if (paragraphs.length === 0) {
    return { type: "text", text: "" } as PomTextNode;
  }

  // Group consecutive paragraphs by bullet type
  const groups = groupByBulletType(paragraphs);

  if (groups.length === 1) {
    return convertParagraphGroup(groups[0]);
  }

  // Multiple groups → VStack
  const children: PomNode[] = groups.map(convertParagraphGroup);
  return {
    type: "vstack",
    children,
  } as PomVStackNode;
}

interface ParagraphGroup {
  bulletType: "none" | "char" | "autoNum";
  paragraphs: Paragraph[];
}

function groupByBulletType(paragraphs: Paragraph[]): ParagraphGroup[] {
  const groups: ParagraphGroup[] = [];
  let current: ParagraphGroup | null = null;

  for (const para of paragraphs) {
    const bt = getBulletCategory(para.properties.bullet);
    if (!current || current.bulletType !== bt) {
      current = { bulletType: bt, paragraphs: [para] };
      groups.push(current);
    } else {
      current.paragraphs.push(para);
    }
  }

  return groups;
}

function getBulletCategory(bullet: BulletType | null): "none" | "char" | "autoNum" {
  if (!bullet) return "none";
  return bullet.type;
}

function convertParagraphGroup(group: ParagraphGroup): PomNode {
  switch (group.bulletType) {
    case "char":
      return convertToUlNode(group.paragraphs);
    case "autoNum":
      return convertToOlNode(group.paragraphs);
    default:
      return convertToTextNode(group.paragraphs);
  }
}

function convertToTextNode(paragraphs: Paragraph[]): PomTextNode {
  const textParts: string[] = [];
  let fontPx: number | undefined;
  let color: string | undefined;
  let bold: boolean | undefined;
  let italic: boolean | undefined;
  let underline: boolean | undefined;
  let strike: boolean | undefined;
  let alignText: PomAlignText | undefined;
  let fontFamily: string | undefined;
  let lineSpacing: number | undefined;

  for (const para of paragraphs) {
    const paraText = para.runs.map((r) => r.text).join("");
    textParts.push(paraText);

    if (fontPx === undefined) {
      for (const run of para.runs) {
        if (run.properties.fontSize !== null) {
          fontPx = ptToPx(run.properties.fontSize as number);
          break;
        }
      }
    }
    if (color === undefined) {
      for (const run of para.runs) {
        if (run.properties.color) {
          color = stripHash(run.properties.color.hex);
          break;
        }
      }
    }
    if (bold === undefined && para.runs.some((r) => r.properties.bold)) {
      bold = true;
    }
    if (italic === undefined && para.runs.some((r) => r.properties.italic)) {
      italic = true;
    }
    if (underline === undefined && para.runs.some((r) => r.properties.underline)) {
      underline = true;
    }
    if (strike === undefined && para.runs.some((r) => r.properties.strikethrough)) {
      strike = true;
    }
    if (alignText === undefined) {
      alignText = convertAlignment(para.properties.alignment);
    }
    if (fontFamily === undefined) {
      for (const run of para.runs) {
        if (run.properties.fontFamily) {
          fontFamily = run.properties.fontFamily;
          break;
        }
      }
    }
    if (lineSpacing === undefined && para.properties.lineSpacing !== null) {
      lineSpacing = para.properties.lineSpacing;
    }
  }

  const node: PomTextNode = {
    type: "text",
    text: textParts.join("\n"),
  };
  if (fontPx !== undefined) node.fontPx = fontPx;
  if (color !== undefined) node.color = color;
  if (bold) node.bold = true;
  if (italic) node.italic = true;
  if (underline) node.underline = true;
  if (strike) node.strike = true;
  if (alignText && alignText !== "left") node.alignText = alignText;
  if (fontFamily) node.fontFamily = fontFamily;
  if (lineSpacing !== undefined) node.lineSpacingMultiple = lineSpacing;

  return node;
}

function convertToUlNode(paragraphs: Paragraph[]): PomUlNode {
  const items: PomLiNode[] = paragraphs.map(convertParagraphToLi);

  const node: PomUlNode = {
    type: "ul",
    items,
  };

  // Inherit common styles from first item
  if (items.length > 0) {
    const first = items[0];
    if (first.fontPx !== undefined) node.fontPx = first.fontPx;
    if (first.color !== undefined) node.color = first.color;
    if (first.fontFamily !== undefined) node.fontFamily = first.fontFamily;
  }

  return node;
}

function convertToOlNode(paragraphs: Paragraph[]): PomOlNode {
  const items: PomLiNode[] = paragraphs.map(convertParagraphToLi);

  const node: PomOlNode = {
    type: "ol",
    items,
  };

  // Inherit common styles from first item
  if (items.length > 0) {
    const first = items[0];
    if (first.fontPx !== undefined) node.fontPx = first.fontPx;
    if (first.color !== undefined) node.color = first.color;
    if (first.fontFamily !== undefined) node.fontFamily = first.fontFamily;
  }

  // Number type
  const bullet = paragraphs[0].properties.bullet;
  if (bullet && bullet.type === "autoNum") {
    const numberType = convertAutoNumScheme(bullet.scheme);
    if (numberType) node.numberType = numberType;
    if (bullet.startAt > 1) node.numberStartAt = bullet.startAt;
  }

  return node;
}

function convertParagraphToLi(para: Paragraph): PomLiNode {
  const text = para.runs.map((r) => r.text).join("");
  const li: PomLiNode = { text };

  for (const run of para.runs) {
    if (run.properties.fontSize !== null && li.fontPx === undefined) {
      li.fontPx = ptToPx(run.properties.fontSize as number);
    }
    if (run.properties.color && li.color === undefined) {
      li.color = stripHash(run.properties.color.hex);
    }
    if (run.properties.fontFamily && li.fontFamily === undefined) {
      li.fontFamily = run.properties.fontFamily;
    }
    if (run.properties.bold && li.bold === undefined) li.bold = true;
    if (run.properties.italic && li.italic === undefined) li.italic = true;
    if (run.properties.underline && li.underline === undefined) li.underline = true;
    if (run.properties.strikethrough && li.strike === undefined) li.strike = true;
  }

  return li;
}

function convertAutoNumScheme(scheme: AutoNumScheme): string | undefined {
  // Map pptx-glimpse scheme to pom numberType
  const mapping: Record<string, string> = {
    arabicPeriod: "arabicPeriod",
    arabicParenR: "arabicParenR",
    romanUcPeriod: "romanUcPeriod",
    romanLcPeriod: "romanLcPeriod",
    alphaUcPeriod: "alphaUcPeriod",
    alphaLcPeriod: "alphaLcPeriod",
    alphaLcParenR: "alphaLcParenR",
    alphaUcParenR: "alphaUcParenR",
    arabicPlain: "arabicPlain",
  };
  return mapping[scheme];
}

function convertAlignment(alignment: "l" | "ctr" | "r" | "just"): PomAlignText {
  switch (alignment) {
    case "ctr":
      return "center";
    case "r":
      return "right";
    default:
      return "left";
  }
}

function ptToPx(pt: number): number {
  return Math.round((pt * 4) / 3);
}
