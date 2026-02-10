import type { Paragraph, RunProperties, TextBody } from "../model/text.js";
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

  for (const para of paragraphs) {
    const { xPos, anchorValue } = getAlignmentInfo(
      para.properties.alignment,
      marginLeft,
      textWidth,
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
      const wrappedLines = wrapParagraph(para, textWidth, scaledDefaultFontSize);
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

        for (let segIdx = 0; segIdx < line.segments.length; segIdx++) {
          const seg = line.segments[segIdx];
          const styles = buildStyleAttrs(seg.properties, fontScale);

          if (segIdx === 0) {
            const lineFontSize = getLineFontSize(line.segments, defaultFontSize) * fontScale;
            const dy = computeDy(
              isFirstLine,
              lineFontSize,
              getLineSpacing(para, lnSpcReduction),
              para,
              lineIdx === 0,
            );
            tspans.push(
              `<tspan x="${xPos}" dy="${dy}" text-anchor="${anchorValue}" ${styles}>${escapeXml(seg.text)}</tspan>`,
            );
          } else {
            tspans.push(`<tspan ${styles}>${escapeXml(seg.text)}</tspan>`);
          }
        }
        isFirstLine = false;
      }
    } else {
      // wrap="none": 折り返しなし
      for (let i = 0; i < para.runs.length; i++) {
        const run = para.runs[i];
        if (run.text.length === 0) continue;
        const styles = buildStyleAttrs(run.properties, fontScale);

        if (i === 0) {
          const dy = computeDy(
            isFirstLine,
            (run.properties.fontSize ?? defaultFontSize) * fontScale,
            getLineSpacing(para, lnSpcReduction),
            para,
            true,
          );
          tspans.push(
            `<tspan x="${xPos}" dy="${dy}" text-anchor="${anchorValue}" ${styles}>${escapeXml(run.text)}</tspan>`,
          );
        } else {
          tspans.push(`<tspan ${styles}>${escapeXml(run.text)}</tspan>`);
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
  );
  if (bodyProperties.anchor === "ctr") {
    yStart = Math.max(marginTop, (height - totalTextHeight) / 2);
  } else if (bodyProperties.anchor === "b") {
    yStart = Math.max(marginTop, height - totalTextHeight - marginBottom);
  }
  yStart += scaledDefaultFontSize;

  return `<text x="0" y="${yStart}">${tspans.join("")}</text>`;
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

function buildStyleAttrs(props: RunProperties, fontScale: number = 1): string {
  const styles: string[] = [];

  if (props.fontSize) {
    const scaledSize = props.fontSize * fontScale;
    styles.push(`font-size="${scaledSize}pt"`);
  }
  if (props.fontFamily) {
    styles.push(`font-family="${escapeXml(props.fontFamily)}"`);
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
): number {
  let totalHeight = 0;

  for (let pIdx = 0; pIdx < paragraphs.length; pIdx++) {
    const para = paragraphs[pIdx];
    const lineSpacing = getLineSpacing(para, lnSpcReduction);
    const lineHeight = defaultFontSize * PX_PER_PT * lineSpacing;

    let lineCount: number;
    if (shouldWrap && para.runs.length > 0 && para.runs.some((r) => r.text.length > 0)) {
      const wrappedLines = wrapParagraph(para, textWidth, defaultFontSize);
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
