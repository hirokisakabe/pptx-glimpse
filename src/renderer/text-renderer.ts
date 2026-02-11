import type {
  AutoNumScheme,
  Paragraph,
  ParagraphProperties,
  RunProperties,
  TextBody,
} from "../model/text.js";
import type { Transform } from "../model/shape.js";
import { emuToPixels } from "../utils/emu.js";
import { wrapParagraph } from "../utils/text-wrap.js";

const PX_PER_PT = 96 / 72;
const DEFAULT_LINE_SPACING = 1.2;
const DEFAULT_FONT_SIZE_PT = 18;

export function renderTextBody(textBody: TextBody, transform: Transform): string {
  const { bodyProperties, paragraphs } = textBody;
  const width = emuToPixels(transform.extentWidth);
  const height = emuToPixels(transform.extentHeight);
  const marginLeft = emuToPixels(bodyProperties.marginLeft);
  const marginRight = emuToPixels(bodyProperties.marginRight);
  const marginTop = emuToPixels(bodyProperties.marginTop);
  const marginBottom = emuToPixels(bodyProperties.marginBottom);

  const hasText = paragraphs.some((p) => p.runs.some((r) => r.text.length > 0));
  if (!hasText) return "";

  const textWidth = width - marginLeft - marginRight;
  const defaultFontSize = getDefaultFontSize(paragraphs);
  const shouldWrap = bodyProperties.wrap !== "none";

  const fontScale = bodyProperties.fontScale;
  const lnSpcReduction = bodyProperties.lnSpcReduction;
  const scaledDefaultFontSize = defaultFontSize * fontScale;

  const tspans: string[] = [];
  let isFirstLine = true;

  // 連番管理用
  const autoNumCounters = new Map<string, number>();

  for (const para of paragraphs) {
    const paraMarginLeft = emuToPixels(para.properties.marginLeft);
    const paraIndent = emuToPixels(para.properties.indent);

    // テキスト開始位置 = bodyMarginLeft + paraMarginLeft
    const textStartX = marginLeft + paraMarginLeft;
    // 箇条書き記号位置 = textStartX + indent (indent は通常負値)
    const bulletX = textStartX + paraIndent;

    const effectiveTextWidth = textWidth - paraMarginLeft;

    const bulletText = resolveBulletText(para.properties, autoNumCounters);

    const { xPos, anchorValue } = getAlignmentInfo(
      para.properties.alignment,
      textStartX,
      effectiveTextWidth,
      width,
      marginRight,
    );

    if (para.runs.length === 0 || !para.runs.some((r) => r.text.length > 0)) {
      const dy = computeDy(isFirstLine, scaledDefaultFontSize, DEFAULT_LINE_SPACING, para, true);
      tspans.push(`<tspan x="${xPos}" dy="${dy}" text-anchor="${anchorValue}"> </tspan>`);
      isFirstLine = false;
      continue;
    }

    if (shouldWrap) {
      const wrappedLines = wrapParagraph(
        para,
        effectiveTextWidth,
        scaledDefaultFontSize,
        fontScale,
      );
      for (let lineIdx = 0; lineIdx < wrappedLines.length; lineIdx++) {
        const line = wrappedLines[lineIdx];
        if (line.segments.length === 0) {
          const dy = computeDy(
            isFirstLine,
            scaledDefaultFontSize,
            getLineSpacing(para, lnSpcReduction),
            para,
            lineIdx === 0,
          );
          tspans.push(`<tspan x="${xPos}" dy="${dy}" text-anchor="${anchorValue}"> </tspan>`);
          isFirstLine = false;
          continue;
        }

        // 最初の行に箇条書き記号を挿入
        if (lineIdx === 0 && bulletText) {
          const lineFontSize = getLineFontSize(line.segments, defaultFontSize) * fontScale;
          const dy = computeDy(
            isFirstLine,
            lineFontSize,
            getLineSpacing(para, lnSpcReduction),
            para,
            true,
          );
          const bulletStyles = buildBulletStyleAttrs(para.properties, lineFontSize, fontScale);
          tspans.push(
            `<tspan x="${bulletX}" dy="${dy}" text-anchor="start" ${bulletStyles}>${escapeXml(bulletText)}</tspan>`,
          );
          // 箇条書き記号の後にテキストを続ける（同じ行で x をテキスト開始位置に設定）
          for (let segIdx = 0; segIdx < line.segments.length; segIdx++) {
            const seg = line.segments[segIdx];
            const prefix = segIdx === 0 ? `x="${xPos}" text-anchor="${anchorValue}" ` : "";
            tspans.push(renderSegment(seg.text, seg.properties, fontScale, prefix));
          }
        } else {
          for (let segIdx = 0; segIdx < line.segments.length; segIdx++) {
            const seg = line.segments[segIdx];

            if (segIdx === 0) {
              const lineFontSize = getLineFontSize(line.segments, defaultFontSize) * fontScale;
              const dy = computeDy(
                isFirstLine,
                lineFontSize,
                getLineSpacing(para, lnSpcReduction),
                para,
                lineIdx === 0,
              );
              const prefix = `x="${xPos}" dy="${dy}" text-anchor="${anchorValue}" `;
              tspans.push(renderSegment(seg.text, seg.properties, fontScale, prefix));
            } else {
              tspans.push(renderSegment(seg.text, seg.properties, fontScale, ""));
            }
          }
        }
        isFirstLine = false;
      }
    } else {
      // wrap="none": 折り返しなし
      let firstRunRendered = false;
      if (bulletText) {
        const firstRun = para.runs.find((r) => r.text.length > 0);
        const fontSize = (firstRun?.properties.fontSize ?? defaultFontSize) * fontScale;
        const dy = computeDy(
          isFirstLine,
          fontSize,
          getLineSpacing(para, lnSpcReduction),
          para,
          true,
        );
        const bulletStyles = buildBulletStyleAttrs(para.properties, fontSize, fontScale);
        tspans.push(
          `<tspan x="${bulletX}" dy="${dy}" text-anchor="start" ${bulletStyles}>${escapeXml(bulletText)}</tspan>`,
        );
      }

      for (let i = 0; i < para.runs.length; i++) {
        const run = para.runs[i];
        if (run.text.length === 0) continue;

        if (!firstRunRendered) {
          if (bulletText) {
            // 箇条書きがある場合、テキストは x 位置を指定（dy なし、箇条書きと同じ行）
            const prefix = `x="${xPos}" text-anchor="${anchorValue}" `;
            tspans.push(renderSegment(run.text, run.properties, fontScale, prefix));
          } else {
            const dy = computeDy(
              isFirstLine,
              (run.properties.fontSize ?? defaultFontSize) * fontScale,
              getLineSpacing(para, lnSpcReduction),
              para,
              true,
            );
            const prefix = `x="${xPos}" dy="${dy}" text-anchor="${anchorValue}" `;
            tspans.push(renderSegment(run.text, run.properties, fontScale, prefix));
          }
          firstRunRendered = true;
        } else {
          tspans.push(renderSegment(run.text, run.properties, fontScale, ""));
        }
      }
      isFirstLine = false;
    }
  }

  // 垂直位置の計算
  let yStart = marginTop;
  const totalTextHeight = estimateTextHeight(
    paragraphs,
    scaledDefaultFontSize,
    shouldWrap,
    textWidth,
    lnSpcReduction,
    fontScale,
  );
  if (bodyProperties.anchor === "ctr") {
    yStart = Math.max(marginTop, (height - totalTextHeight) / 2);
  } else if (bodyProperties.anchor === "b") {
    yStart = Math.max(marginTop, height - totalTextHeight - marginBottom);
  }
  yStart += scaledDefaultFontSize;

  return `<text x="0" y="${yStart}">${tspans.join("")}</text>`;
}

function resolveBulletText(
  props: ParagraphProperties,
  autoNumCounters: Map<string, number>,
): string | null {
  if (!props.bullet) return null;
  if (props.bullet.type === "none") return null;

  if (props.bullet.type === "char") {
    return props.bullet.char;
  }

  if (props.bullet.type === "autoNum") {
    const { scheme, startAt } = props.bullet;
    const counterKey = `${scheme}-${props.level}`;
    const current = autoNumCounters.get(counterKey) ?? 0;
    const nextVal = current + 1;
    autoNumCounters.set(counterKey, nextVal);
    const index = startAt + nextVal - 1;
    return formatAutoNum(scheme, index);
  }

  return null;
}

export function formatAutoNum(scheme: AutoNumScheme, index: number): string {
  switch (scheme) {
    case "arabicPeriod":
      return `${index}.`;
    case "arabicParenR":
      return `${index})`;
    case "arabicPlain":
      return `${index}`;
    case "romanUcPeriod":
      return `${toRoman(index)}.`;
    case "romanLcPeriod":
      return `${toRoman(index).toLowerCase()}.`;
    case "alphaUcPeriod":
      return `${toAlpha(index)}.`;
    case "alphaLcPeriod":
      return `${toAlpha(index).toLowerCase()}.`;
    case "alphaUcParenR":
      return `${toAlpha(index)})`;
    case "alphaLcParenR":
      return `${toAlpha(index).toLowerCase()})`;
    default:
      return `${index}.`;
  }
}

function toRoman(num: number): string {
  const romanNumerals: [number, string][] = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ];
  let result = "";
  let remaining = num;
  for (const [value, symbol] of romanNumerals) {
    while (remaining >= value) {
      result += symbol;
      remaining -= value;
    }
  }
  return result;
}

function toAlpha(num: number): string {
  let result = "";
  let remaining = num;
  while (remaining > 0) {
    remaining--;
    result = String.fromCharCode(65 + (remaining % 26)) + result;
    remaining = Math.floor(remaining / 26);
  }
  return result;
}

function buildBulletStyleAttrs(
  props: ParagraphProperties,
  textFontSizePt: number,
  fontScale: number,
): string {
  const styles: string[] = [];

  if (props.bulletSizePct !== null) {
    const size = textFontSizePt * (props.bulletSizePct / 100000);
    styles.push(`font-size="${size}pt"`);
  }

  if (props.bulletFont) {
    styles.push(`font-family="${escapeXml(props.bulletFont)}"`);
  }

  if (props.bulletColor) {
    styles.push(`fill="${props.bulletColor.hex}"`);
    if (props.bulletColor.alpha < 1) {
      styles.push(`fill-opacity="${props.bulletColor.alpha}"`);
    }
  }

  // fontScale が 1 でない場合、サイズ未指定なら textFontSizePt をそのまま使用
  // (textFontSizePt は既に fontScale 適用済み)
  void fontScale;

  return styles.join(" ");
}

function getAlignmentInfo(
  alignment: "l" | "ctr" | "r" | "just",
  marginLeft: number,
  textWidth: number,
  width: number,
  marginRight: number,
): { xPos: number; anchorValue: string } {
  switch (alignment) {
    case "ctr":
      return { xPos: marginLeft + textWidth / 2, anchorValue: "middle" };
    case "r":
      return { xPos: width - marginRight, anchorValue: "end" };
    default:
      return { xPos: marginLeft, anchorValue: "start" };
  }
}

function getLineSpacing(para: Paragraph, lnSpcReduction: number = 0): number {
  let spacing: number;
  if (para.properties.lineSpacing !== null) {
    const factor = para.properties.lineSpacing / 100000;
    spacing = Math.max(0.5, factor);
  } else {
    spacing = DEFAULT_LINE_SPACING;
  }
  return spacing * (1 - lnSpcReduction);
}

function computeDy(
  isFirstLine: boolean,
  fontSizePt: number,
  lineSpacingFactor: number,
  para: Paragraph,
  isFirstLineOfParagraph: boolean,
): string {
  if (isFirstLine) return "0";

  const lineHeight = fontSizePt * PX_PER_PT * lineSpacingFactor;
  let dy = lineHeight;

  if (isFirstLineOfParagraph && para.properties.spaceBefore > 0) {
    // spaceBefore は 1/100 ポイント単位
    dy += (para.properties.spaceBefore / 100) * PX_PER_PT;
  }

  return dy.toFixed(2);
}

function getLineFontSize(
  segments: { text: string; properties: RunProperties }[],
  defaultFontSize: number,
): number {
  for (const seg of segments) {
    if (seg.properties.fontSize) return seg.properties.fontSize;
  }
  return defaultFontSize;
}

function isCjkCodePoint(codePoint: number): boolean {
  return (
    (codePoint >= 0x3000 && codePoint <= 0x9fff) ||
    (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
    (codePoint >= 0xff01 && codePoint <= 0xff60) ||
    (codePoint >= 0x20000 && codePoint <= 0x2a6df)
  );
}

interface ScriptSegment {
  text: string;
  isEa: boolean;
}

function splitByScript(text: string): ScriptSegment[] {
  const segments: ScriptSegment[] = [];
  let current = "";
  let currentIsEa: boolean | null = null;

  for (const char of text) {
    const cp = char.codePointAt(0)!;
    const isEa = isCjkCodePoint(cp);

    if (currentIsEa !== null && isEa !== currentIsEa) {
      segments.push({ text: current, isEa: currentIsEa });
      current = "";
    }
    currentIsEa = isEa;
    current += char;
  }

  if (current && currentIsEa !== null) {
    segments.push({ text: current, isEa: currentIsEa });
  }

  return segments;
}

function needsScriptSplit(props: RunProperties): boolean {
  return (
    props.fontFamily !== null &&
    props.fontFamilyEa !== null &&
    props.fontFamily !== props.fontFamilyEa
  );
}

function buildStyleAttrs(
  props: RunProperties,
  fontScale: number = 1,
  fontOverride?: string | null,
): string {
  const styles: string[] = [];

  if (props.fontSize) {
    const scaledSize = props.fontSize * fontScale;
    styles.push(`font-size="${scaledSize}pt"`);
  }
  const fontFamily = fontOverride !== undefined ? fontOverride : props.fontFamily;
  if (fontFamily) {
    styles.push(`font-family="${escapeXml(fontFamily)}"`);
  }
  if (props.bold) {
    styles.push(`font-weight="bold"`);
  }
  if (props.italic) {
    styles.push(`font-style="italic"`);
  }
  if (props.color) {
    styles.push(`fill="${props.color.hex}"`);
    if (props.color.alpha < 1) {
      styles.push(`fill-opacity="${props.color.alpha}"`);
    }
  }

  const decorations: string[] = [];
  if (props.underline) decorations.push("underline");
  if (props.strikethrough) decorations.push("line-through");
  if (decorations.length > 0) {
    styles.push(`text-decoration="${decorations.join(" ")}"`);
  }

  if (props.baseline > 0) {
    styles.push(`baseline-shift="super"`);
  } else if (props.baseline < 0) {
    styles.push(`baseline-shift="sub"`);
  }

  return styles.join(" ");
}

function renderSegment(
  text: string,
  props: RunProperties,
  fontScale: number,
  prefix: string,
): string {
  if (!needsScriptSplit(props)) {
    const styles = buildStyleAttrs(props, fontScale);
    return `<tspan ${prefix}${styles}>${escapeXml(text)}</tspan>`;
  }
  const parts = splitByScript(text);
  const result: string[] = [];
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    const font = part.isEa ? props.fontFamilyEa : props.fontFamily;
    const styles = buildStyleAttrs(props, fontScale, font);
    if (i === 0) {
      result.push(`<tspan ${prefix}${styles}>${escapeXml(part.text)}</tspan>`);
    } else {
      result.push(`<tspan ${styles}>${escapeXml(part.text)}</tspan>`);
    }
  }
  return result.join("");
}

function getDefaultFontSize(paragraphs: TextBody["paragraphs"]): number {
  for (const p of paragraphs) {
    for (const r of p.runs) {
      if (r.properties.fontSize) return r.properties.fontSize;
    }
  }
  return DEFAULT_FONT_SIZE_PT;
}

function estimateTextHeight(
  paragraphs: TextBody["paragraphs"],
  defaultFontSize: number,
  shouldWrap: boolean,
  textWidth: number,
  lnSpcReduction: number = 0,
  fontScale: number = 1,
): number {
  let totalHeight = 0;

  for (let pIdx = 0; pIdx < paragraphs.length; pIdx++) {
    const para = paragraphs[pIdx];
    const lineSpacing = getLineSpacing(para, lnSpcReduction);
    const lineHeight = defaultFontSize * PX_PER_PT * lineSpacing;

    let lineCount: number;
    if (shouldWrap && para.runs.length > 0 && para.runs.some((r) => r.text.length > 0)) {
      const wrappedLines = wrapParagraph(para, textWidth, defaultFontSize, fontScale);
      lineCount = wrappedLines.length;
    } else {
      lineCount = para.runs.some((r) => r.text.length > 0) ? 1 : 1;
    }

    totalHeight += lineCount * lineHeight;

    if (pIdx > 0 && para.properties.spaceBefore > 0) {
      totalHeight += (para.properties.spaceBefore / 100) * PX_PER_PT;
    }
  }

  return totalHeight;
}

function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
