import { getMetricsFallbackFont } from "../data/font-metrics.js";
import type { TextPathFontResolver } from "../font/text-path-context.js";
import type { Transform } from "../model/shape.js";
import type {
  AutoNumScheme,
  BodyProperties,
  Paragraph,
  ParagraphProperties,
  RunProperties,
  SpacingValue,
  TextBody,
  TextVerticalType,
} from "../model/text.js";
import { EMU_PER_INCH } from "../utils/constants.js";
import { emuToPixels } from "../utils/emu.js";
import { wrapParagraph } from "../utils/text-wrap.js";
import type { Emu } from "../utils/unit-types.js";
import { asEmu } from "../utils/unit-types.js";
import type { RendererContext } from "./render-context.js";
import {
  createLegacyRendererContext,
  getJpanFallbackFontFromContext,
  getMappedFontFromContext,
} from "./render-context.js";

const PX_PER_PT = 96 / 72;
const DEFAULT_LINE_SPACING = 1.0;
const DEFAULT_FONT_SIZE_PT = 18;

function isVerticalText(vert: TextVerticalType): boolean {
  return vert === "vert" || vert === "eaVert" || vert === "wordArtVert" || vert === "mongolianVert";
}

function isVert270Text(vert: TextVerticalType): boolean {
  return vert === "vert270";
}

/**
 * Replace dimensions and margins when writing vertically.
 * Configure the layout space so that the text appears in the correct position after the rotation transformation.
 */
function resolveTextDimensions(
  bodyProperties: BodyProperties,
  originalWidth: number,
  originalHeight: number,
): {
  width: number;
  height: number;
  marginLeftPx: number;
  marginRightPx: number;
  marginTopPx: number;
  marginBottomPx: number;
} {
  const vert = bodyProperties.vert;

  if (isVerticalText(vert)) {
    // vert (90° CW): Layout space is HxW, margins are swapped according to rotation
    return {
      width: originalHeight,
      height: originalWidth,
      marginLeftPx: emuToPixels(bodyProperties.marginTop),
      marginRightPx: emuToPixels(bodyProperties.marginBottom),
      marginTopPx: emuToPixels(bodyProperties.marginRight),
      marginBottomPx: emuToPixels(bodyProperties.marginLeft),
    };
  }

  if (isVert270Text(vert)) {
    // vert270 (90° CCW): Layout space is HxW, margins are swapped in the opposite direction
    return {
      width: originalHeight,
      height: originalWidth,
      marginLeftPx: emuToPixels(bodyProperties.marginBottom),
      marginRightPx: emuToPixels(bodyProperties.marginTop),
      marginTopPx: emuToPixels(bodyProperties.marginLeft),
      marginBottomPx: emuToPixels(bodyProperties.marginRight),
    };
  }

  // horizontal text
  return {
    width: originalWidth,
    height: originalHeight,
    marginLeftPx: emuToPixels(bodyProperties.marginLeft),
    marginRightPx: emuToPixels(bodyProperties.marginRight),
    marginTopPx: emuToPixels(bodyProperties.marginTop),
    marginBottomPx: emuToPixels(bodyProperties.marginBottom),
  };
}

export function renderTextBody(
  textBody: TextBody,
  transform: Transform,
  context: RendererContext = createLegacyRendererContext(),
): string {
  const fontResolver = context.textPathFontResolver;
  if (fontResolver) {
    return renderTextBodyAsPath(textBody, transform, fontResolver, context);
  }

  const { bodyProperties, paragraphs } = textBody;
  const originalWidth = emuToPixels(transform.extentWidth);
  const originalHeight = emuToPixels(transform.extentHeight);

  const { width, height, marginLeftPx, marginRightPx, marginTopPx, marginBottomPx } =
    resolveTextDimensions(bodyProperties, originalWidth, originalHeight);

  const hasText = paragraphs.some((p) => p.runs.some((r) => r.text.length > 0));
  if (!hasText) return "";

  const fullTextWidth = width - marginLeftPx - marginRightPx;
  const numCol = bodyProperties.numCol ?? 1;
  const textWidth = numCol > 1 ? fullTextWidth / numCol : fullTextWidth;
  const defaultFontSize = getDefaultFontSize(paragraphs);
  const shouldWrap = bodyProperties.wrap !== "none";

  let fontScale = bodyProperties.fontScale;
  const lnSpcReduction = bodyProperties.lnSpcReduction;

  // normAutofit: dynamically reduce fontScale if text extends off shape
  if (bodyProperties.autoFit === "normAutofit" && shouldWrap) {
    const availableHeight = height - marginTopPx - marginBottomPx;
    fontScale = computeShrinkToFitScale(
      paragraphs,
      defaultFontSize,
      fontScale,
      lnSpcReduction,
      textWidth,
      availableHeight,
      context,
    );
  }

  const scaledDefaultFontSizePt = defaultFontSize * fontScale;

  // Default font line height ratio
  const defaultLineHeightRatio = getDefaultLineHeightRatio(paragraphs, context);
  const defaultAscenderRatio = getDefaultAscenderRatio(paragraphs, context);
  const defaultNaturalHeightPt = scaledDefaultFontSizePt * defaultLineHeightRatio;

  const tspans: string[] = [];
  let isFirstLine = true;

  // For serial number management
  const autoNumCounters = new Map<string, number>();

  // spaceAfter (pixel resolved) in previous paragraph
  let prevSpaceAfterPx = 0;

  for (const para of paragraphs) {
    const paraMarginLeft = emuToPixels(para.properties.marginLeft ?? asEmu(0));
    const paraIndent = emuToPixels(para.properties.indent ?? asEmu(0));

    // Text start position = bodyMarginLeft + paraMarginLeft
    const textStartX = marginLeftPx + paraMarginLeft;
    // Bullet point position = textStartX + indent (indent is usually a negative value)
    const bulletX = textStartX + paraIndent;

    const effectiveTextWidth = textWidth - paraMarginLeft;

    const bulletText = resolveBulletText(para.properties, autoNumCounters);

    const { xPos, anchorValue } = getAlignmentInfo(
      para.properties.alignment,
      textStartX,
      effectiveTextWidth,
      width,
      marginRightPx,
    );

    // Paragraph spacing calculation: max(previous paragraph spaceAfter, current paragraph spaceBefore)
    const paraFontSizePt = getParagraphFontSize(para, defaultFontSize) * fontScale;
    const spaceBeforePx = resolveSpacingPx(para.properties.spaceBefore, paraFontSizePt);
    const paragraphGapPx = Math.max(prevSpaceAfterPx, spaceBeforePx);

    if (para.runs.length === 0 || !para.runs.some((r) => r.text.length > 0)) {
      const emptyParaHeightPt = paraFontSizePt > 0 ? paraFontSizePt : defaultNaturalHeightPt;
      const dy = computeDy(
        isFirstLine,
        getLineHeightPx(para, emptyParaHeightPt, lnSpcReduction),
        paragraphGapPx,
      );
      tspans.push(`<tspan x="${xPos}" dy="${dy}" text-anchor="${anchorValue}"> </tspan>`);
      isFirstLine = false;
      prevSpaceAfterPx = resolveSpacingPx(para.properties.spaceAfter, paraFontSizePt);
      continue;
    }

    if (shouldWrap) {
      const wrappedLines = wrapParagraph(
        para,
        effectiveTextWidth,
        scaledDefaultFontSizePt,
        fontScale,
        context.textMeasurer,
      );
      for (let lineIdx = 0; lineIdx < wrappedLines.length; lineIdx++) {
        const line = wrappedLines[lineIdx];
        const lineGapPx = lineIdx === 0 ? paragraphGapPx : 0;
        if (line.segments.length === 0) {
          const dy = computeDy(
            isFirstLine,
            getLineHeightPx(para, defaultNaturalHeightPt, lnSpcReduction),
            lineGapPx,
          );
          tspans.push(`<tspan x="${xPos}" dy="${dy}" text-anchor="${anchorValue}"> </tspan>`);
          isFirstLine = false;
          continue;
        }

        // Insert bullet point on first line
        if (lineIdx === 0 && bulletText) {
          const lineFontSize = getLineFontSize(line.segments, defaultFontSize) * fontScale;
          const lineNaturalHeightPt = computeLineNaturalHeight(
            line.segments,
            defaultFontSize,
            fontScale,
            context,
          );
          const dy = computeDy(
            isFirstLine,
            getLineHeightPx(para, lineNaturalHeightPt, lnSpcReduction),
            paragraphGapPx,
          );
          const firstSeg = line.segments[0];
          const bulletFontChain = buildBulletFontChain(
            para.properties,
            firstSeg?.properties.fontFamily,
            firstSeg?.properties.fontFamilyEa,
          );
          const bulletStyles = buildBulletStyleAttrs(
            para.properties,
            lineFontSize,
            fontScale,
            bulletFontChain,
            context,
          );
          context.fontUsageCollector?.record(bulletFontChain, bulletText);
          tspans.push(
            `<tspan x="${bulletX}" dy="${dy}" text-anchor="start" ${bulletStyles}>${escapeXml(bulletText)}</tspan>`,
          );
          // Continue text after bullet point (set x to start text on same line)
          for (let segIdx = 0; segIdx < line.segments.length; segIdx++) {
            const seg = line.segments[segIdx];
            const prefix = segIdx === 0 ? `x="${xPos}" text-anchor="${anchorValue}" ` : "";
            tspans.push(renderSegment(seg.text, seg.properties, fontScale, prefix, context));
          }
        } else {
          for (let segIdx = 0; segIdx < line.segments.length; segIdx++) {
            const seg = line.segments[segIdx];

            if (segIdx === 0) {
              const lineNaturalHeightPt = computeLineNaturalHeight(
                line.segments,
                defaultFontSize,
                fontScale,
                context,
              );
              const dy = computeDy(
                isFirstLine,
                getLineHeightPx(para, lineNaturalHeightPt, lnSpcReduction),
                lineGapPx,
              );
              const prefix = `x="${xPos}" dy="${dy}" text-anchor="${anchorValue}" `;
              tspans.push(renderSegment(seg.text, seg.properties, fontScale, prefix, context));
            } else {
              tspans.push(renderSegment(seg.text, seg.properties, fontScale, "", context));
            }
          }
        }
        isFirstLine = false;
      }
    } else {
      // wrap="none": no wrapping
      let firstRunRendered = false;
      if (bulletText) {
        const firstRun = para.runs.find((r) => r.text.length > 0);
        const fontSize = (firstRun?.properties.fontSize ?? defaultFontSize) * fontScale;
        const naturalHeightPt = computeLineNaturalHeight(
          para.runs,
          defaultFontSize,
          fontScale,
          context,
        );
        const dy = computeDy(
          isFirstLine,
          getLineHeightPx(para, naturalHeightPt, lnSpcReduction),
          paragraphGapPx,
        );
        const bulletFontChain = buildBulletFontChain(
          para.properties,
          firstRun?.properties.fontFamily,
          firstRun?.properties.fontFamilyEa,
        );
        const bulletStyles = buildBulletStyleAttrs(
          para.properties,
          fontSize,
          fontScale,
          bulletFontChain,
          context,
        );
        context.fontUsageCollector?.record(bulletFontChain, bulletText);
        tspans.push(
          `<tspan x="${bulletX}" dy="${dy}" text-anchor="start" ${bulletStyles}>${escapeXml(bulletText)}</tspan>`,
        );
      }

      for (let i = 0; i < para.runs.length; i++) {
        const run = para.runs[i];
        if (run.text.length === 0) continue;

        if (!firstRunRendered) {
          if (bulletText) {
            // If there is a bullet, the text specifies the x position (no dy, same line as the bullet)
            const prefix = `x="${xPos}" text-anchor="${anchorValue}" `;
            tspans.push(renderSegment(run.text, run.properties, fontScale, prefix, context));
          } else {
            const naturalHeightPt = computeLineNaturalHeight(
              para.runs,
              defaultFontSize,
              fontScale,
              context,
            );
            const dy = computeDy(
              isFirstLine,
              getLineHeightPx(para, naturalHeightPt, lnSpcReduction),
              paragraphGapPx,
            );
            const prefix = `x="${xPos}" dy="${dy}" text-anchor="${anchorValue}" `;
            tspans.push(renderSegment(run.text, run.properties, fontScale, prefix, context));
          }
          firstRunRendered = true;
        } else {
          tspans.push(renderSegment(run.text, run.properties, fontScale, "", context));
        }
      }
      isFirstLine = false;
    }

    prevSpaceAfterPx = resolveSpacingPx(para.properties.spaceAfter, paraFontSizePt);
  }

  // Vertical position calculation
  let yStart = marginTopPx;
  const totalTextHeight = estimateTextHeight(
    paragraphs,
    defaultFontSize,
    shouldWrap,
    textWidth,
    lnSpcReduction,
    fontScale,
    context,
  );
  if (bodyProperties.anchor === "ctr") {
    yStart = Math.max(marginTopPx, (height - totalTextHeight) / 2);
  } else if (bodyProperties.anchor === "b") {
    yStart = Math.max(marginTopPx, height - totalTextHeight - marginBottomPx);
  }
  const firstParaFontSizePt = getParagraphFontSize(paragraphs[0], defaultFontSize) * fontScale;
  const firstLineBaselineOffsetPt = firstParaFontSizePt * defaultAscenderRatio;
  yStart += firstLineBaselineOffsetPt * PX_PER_PT;

  const textElement = `<text x="0" y="${yStart}" xml:space="preserve">${tspans.join("")}</text>`;

  if (isVerticalText(bodyProperties.vert)) {
    return `<g transform="translate(${originalWidth}, 0) rotate(90)">${textElement}</g>`;
  }
  if (isVert270Text(bodyProperties.vert)) {
    return `<g transform="translate(0, ${originalHeight}) rotate(-90)">${textElement}</g>`;
  }
  return textElement;
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

/**
 * Build a font chain for bullet points.
 * bulletFont If not specified, falls back to the text run font (same rules as path drawing).
 */
function buildBulletFontChain(
  props: ParagraphProperties,
  runFontFamily: string | null | undefined,
  runFontFamilyEa: string | null | undefined,
): (string | null)[] {
  return [props.bulletFont, runFontFamily ?? null, runFontFamilyEa ?? null];
}

function buildBulletStyleAttrs(
  props: ParagraphProperties,
  textFontSizePt: number,
  fontScale: number,
  bulletFontChain: (string | null)[],
  context: RendererContext = createLegacyRendererContext(),
): string {
  const styles: string[] = [];

  if (props.bulletSizePct !== null) {
    const size = textFontSizePt * (props.bulletSizePct / 100000);
    styles.push(`font-size="${size}pt"`);
  }

  const fontFamilyValue = buildFontFamilyValue(bulletFontChain, context);
  if (fontFamilyValue) {
    styles.push(`font-family="${fontFamilyValue}"`);
  }

  if (props.bulletColor) {
    styles.push(`fill="${props.bulletColor.hex}"`);
    if (props.bulletColor.alpha < 1) {
      styles.push(`fill-opacity="${props.bulletColor.alpha}"`);
    }
  }

  // If fontScale is not 1, use textFontSizePt as is if size is not specified.
  // (textFontSizePt already has fontScale applied)
  void fontScale;

  return styles.join(" ");
}

function getAlignmentInfo(
  alignment: "l" | "ctr" | "r" | "just" | null,
  marginLeftPx: number,
  textWidth: number,
  width: number,
  marginRightPx: number,
): { xPos: number; anchorValue: string } {
  if (alignment === "ctr") return { xPos: marginLeftPx + textWidth / 2, anchorValue: "middle" };
  if (alignment === "r") return { xPos: width - marginRightPx, anchorValue: "end" };
  return { xPos: marginLeftPx, anchorValue: "start" };
}

/**
 * Returns the height of a paragraph in pixels.
 * If lnSpc is spcPts (fixed line spacing), it is a fixed value independent of font size,
 * spcPct (magnification)/If not specified, calculate by naturalHeightPt x magnification.
 */
function getLineHeightPx(
  para: Paragraph,
  naturalHeightPt: number,
  lnSpcReduction: number = 0,
): number {
  const lineSpacing = para.properties.lineSpacing;
  if (lineSpacing?.type === "pts") {
    return (lineSpacing.value / 100) * PX_PER_PT * (1 - lnSpcReduction);
  }
  const factor =
    lineSpacing !== null ? Math.max(0.5, lineSpacing.value / 100000) : DEFAULT_LINE_SPACING;
  return naturalHeightPt * PX_PER_PT * factor * (1 - lnSpcReduction);
}

function resolveSpacingPx(spacing: SpacingValue, fontSizePt: number): number {
  if (spacing.type === "pts") {
    return (spacing.value / 100) * PX_PER_PT;
  }
  // pct: val / 100000 is the ratio to the font size
  return fontSizePt * (spacing.value / 100000) * PX_PER_PT;
}

function getParagraphFontSize(para: Paragraph, defaultFontSize: number): number {
  for (const run of para.runs) {
    if (run.text.length > 0 && run.properties.fontSize) {
      return run.properties.fontSize;
    }
  }
  if (para.endParaRunProperties?.fontSize) {
    return para.endParaRunProperties.fontSize;
  }
  return defaultFontSize;
}

function computeDy(isFirstLine: boolean, lineHeightPx: number, paragraphGapPx: number): string {
  if (isFirstLine) return "0";

  return (lineHeightPx + paragraphGapPx).toFixed(2);
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

function computeLineNaturalHeight(
  segments: { properties: RunProperties }[],
  defaultFontSize: number,
  fontScale: number,
  context: RendererContext = createLegacyRendererContext(),
): number {
  let maxHeight = 0;
  for (const seg of segments) {
    const fontSize = (seg.properties.fontSize ?? defaultFontSize) * fontScale;
    const ratio = context.textMeasurer.getLineHeightRatio(
      seg.properties.fontFamily,
      seg.properties.fontFamilyEa,
    );
    maxHeight = Math.max(maxHeight, fontSize * ratio);
  }
  return maxHeight > 0 ? maxHeight : defaultFontSize * fontScale * 1.2;
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

function getGenericFamily(fontFamily: string): string {
  const lower = fontFamily.toLowerCase();
  if (
    lower.includes("mincho") ||
    lower.includes("明朝") ||
    lower === "times new roman" ||
    lower === "georgia" ||
    lower === "cambria" ||
    (lower.includes("serif") && !lower.includes("sans"))
  ) {
    return "serif";
  }
  return "sans-serif";
}

function escapeFontName(name: string): string {
  return name
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function buildFontFamilyValue(
  fonts: (string | null)[],
  context: RendererContext = createLegacyRendererContext(),
): string | null {
  const uniqueFonts: string[] = [];
  const seen = new Set<string>();

  for (const font of fonts) {
    if (font && !seen.has(font)) {
      seen.add(font);
      uniqueFonts.push(font);

      // Add OSS alternative font from mapping table
      const mapped = getMappedFontFromContext(font, context);
      if (mapped && !seen.has(mapped)) {
        seen.add(mapped);
        uniqueFonts.push(mapped);
      }

      // Added metrics compatible OSS font as fallback
      const fallback = getMetricsFallbackFont(font);
      if (fallback && !seen.has(fallback)) {
        seen.add(fallback);
        uniqueFonts.push(fallback);
      }
    }
  }

  if (uniqueFonts.length === 0) return null;

  const genericFamily = getGenericFamily(uniqueFonts[0]);

  const parts = uniqueFonts.map((f) => {
    const escaped = escapeFontName(f);
    return f.includes(" ") ? `'${escaped}'` : escaped;
  });
  parts.push(genericFamily);

  return parts.join(", ");
}

function buildStyleAttrs(
  props: RunProperties,
  fontScale: number = 1,
  fontFamilies?: (string | null)[],
  context: RendererContext = createLegacyRendererContext(),
): string {
  const styles: string[] = [];

  if (props.fontSize) {
    const scaledSize = props.fontSize * fontScale;
    styles.push(`font-size="${scaledSize}pt"`);
  }
  const fonts = fontFamilies ?? [props.fontFamily, props.fontFamilyEa];
  const fontFamilyValue = buildFontFamilyValue(fonts, context);
  if (fontFamilyValue) {
    styles.push(`font-family="${fontFamilyValue}"`);
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

  if (props.outline) {
    const strokeWidth = emuToPixels(props.outline.width);
    styles.push(`stroke="${props.outline.color.hex}"`);
    styles.push(`stroke-width="${strokeWidth}"`);
    if (props.outline.color.alpha < 1) {
      styles.push(`stroke-opacity="${props.outline.color.alpha}"`);
    }
    styles.push(`paint-order="stroke"`);
  }

  return styles.join(" ");
}

function renderSegment(
  text: string,
  props: RunProperties,
  fontScale: number,
  prefix: string,
  context: RendererContext,
): string {
  let tspanContent: string;
  if (!needsScriptSplit(props)) {
    const styles = buildStyleAttrs(props, fontScale, undefined, context);
    context.fontUsageCollector?.record([props.fontFamily, props.fontFamilyEa], text);
    tspanContent = `<tspan ${prefix}${styles}>${escapeXml(text)}</tspan>`;
  } else {
    const parts = splitByScript(text);
    const result: string[] = [];
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i];
      const fonts = part.isEa
        ? [props.fontFamilyEa, getJpanFallbackFontFromContext(context), props.fontFamily]
        : [props.fontFamily, props.fontFamilyEa];
      const styles = buildStyleAttrs(props, fontScale, fonts, context);
      context.fontUsageCollector?.record(fonts, part.text);
      if (i === 0) {
        result.push(`<tspan ${prefix}${styles}>${escapeXml(part.text)}</tspan>`);
      } else {
        result.push(`<tspan ${styles}>${escapeXml(part.text)}</tspan>`);
      }
    }
    tspanContent = result.join("");
  }

  if (props.hyperlink) {
    const href = escapeXml(props.hyperlink.url);
    return `<a href="${href}">${tspanContent}</a>`;
  }
  return tspanContent;
}

function getDefaultFontSize(paragraphs: TextBody["paragraphs"]): number {
  for (const p of paragraphs) {
    for (const r of p.runs) {
      if (r.properties.fontSize) return r.properties.fontSize;
    }
  }
  return DEFAULT_FONT_SIZE_PT;
}

function getDefaultLineHeightRatio(
  paragraphs: TextBody["paragraphs"],
  context: RendererContext = createLegacyRendererContext(),
): number {
  for (const p of paragraphs) {
    for (const r of p.runs) {
      if (r.properties.fontFamily || r.properties.fontFamilyEa) {
        return context.textMeasurer.getLineHeightRatio(
          r.properties.fontFamily,
          r.properties.fontFamilyEa,
        );
      }
    }
  }
  return 1.2;
}

function getDefaultAscenderRatio(
  paragraphs: TextBody["paragraphs"],
  context: RendererContext = createLegacyRendererContext(),
): number {
  for (const p of paragraphs) {
    for (const r of p.runs) {
      if (r.properties.fontFamily || r.properties.fontFamilyEa) {
        return context.textMeasurer.getAscenderRatio(
          r.properties.fontFamily,
          r.properties.fontFamilyEa,
        );
      }
    }
  }
  return 1.0;
}

/**
 * spAutofit: Calculates the required shape height (EMU) depending on the amount of text.
 * Returns null if the text fits within the original shape.
 */
export function computeSpAutofitHeight(
  textBody: TextBody,
  transform: Transform,
  context: RendererContext = createLegacyRendererContext(),
): Emu | null {
  const { bodyProperties, paragraphs } = textBody;

  const hasText = paragraphs.some((p) => p.runs.some((r) => r.text.length > 0));
  if (!hasText) return null;

  const originalWidth = emuToPixels(transform.extentWidth);
  const originalHeight = emuToPixels(transform.extentHeight);

  const { width, height, marginLeftPx, marginRightPx, marginTopPx, marginBottomPx } =
    resolveTextDimensions(bodyProperties, originalWidth, originalHeight);

  const fullTextWidth = width - marginLeftPx - marginRightPx;
  const numCol = bodyProperties.numCol ?? 1;
  const textWidth = numCol > 1 ? fullTextWidth / numCol : fullTextWidth;
  const defaultFontSize = getDefaultFontSize(paragraphs);
  const shouldWrap = bodyProperties.wrap !== "none";

  const textHeight = estimateTextHeight(
    paragraphs,
    defaultFontSize,
    shouldWrap,
    textWidth,
    undefined,
    undefined,
    context,
  );
  const requiredHeightPx = textHeight + marginTopPx + marginBottomPx;

  if (requiredHeightPx <= height) return null;

  const DEFAULT_DPI = 96;
  return asEmu((requiredHeightPx / DEFAULT_DPI) * EMU_PER_INCH);
}

function computeShrinkToFitScale(
  paragraphs: TextBody["paragraphs"],
  defaultFontSize: number,
  fontScale: number,
  lnSpcReduction: number,
  textWidth: number,
  availableHeight: number,
  context: RendererContext = createLegacyRendererContext(),
): number {
  if (availableHeight <= 0) return fontScale;

  const minScale = fontScale * 0.1;
  let scale = fontScale;
  const maxIterations = 5;

  for (let i = 0; i < maxIterations; i++) {
    const textHeight = estimateTextHeight(
      paragraphs,
      defaultFontSize,
      true,
      textWidth,
      lnSpcReduction,
      scale,
      context,
    );
    if (textHeight <= availableHeight) break;
    const newScale = scale * (availableHeight / textHeight);
    scale = Math.max(newScale, minScale);
    if (scale <= minScale) break;
  }

  return scale;
}

function estimateTextHeight(
  paragraphs: TextBody["paragraphs"],
  defaultFontSize: number,
  shouldWrap: boolean,
  textWidth: number,
  lnSpcReduction: number = 0,
  fontScale: number = 1,
  context: RendererContext = createLegacyRendererContext(),
): number {
  let totalHeight = 0;
  const defaultRatio = getDefaultLineHeightRatio(paragraphs, context);
  let prevSpaceAfterPx = 0;
  // wrapParagraph expects a pre-scaled defaultFontSize
  const scaledDefaultForWrap = defaultFontSize * fontScale;

  for (let pIdx = 0; pIdx < paragraphs.length; pIdx++) {
    const para = paragraphs[pIdx];
    const isEmpty = !para.runs.some((r) => r.text.length > 0);
    const naturalHeightPt =
      isEmpty && para.endParaRunProperties?.fontSize
        ? para.endParaRunProperties.fontSize * fontScale * defaultRatio
        : computeLineNaturalHeight(para.runs, defaultFontSize, fontScale, context);
    const lineHeight = getLineHeightPx(
      para,
      naturalHeightPt > 0 ? naturalHeightPt : defaultFontSize * fontScale * defaultRatio,
      lnSpcReduction,
    );

    let lineCount: number;
    if (shouldWrap && para.runs.length > 0 && para.runs.some((r) => r.text.length > 0)) {
      const wrappedLines = wrapParagraph(
        para,
        textWidth,
        scaledDefaultForWrap,
        fontScale,
        context.textMeasurer,
      );
      lineCount = wrappedLines.length;
    } else {
      lineCount = para.runs.some((r) => r.text.length > 0) ? 1 : 1;
    }

    totalHeight += lineCount * lineHeight;

    if (pIdx > 0) {
      const paraFontSizePt = getParagraphFontSize(para, defaultFontSize) * fontScale;
      const spaceBeforePx = resolveSpacingPx(para.properties.spaceBefore, paraFontSizePt);
      totalHeight += Math.max(prevSpaceAfterPx, spaceBeforePx);
    }

    const paraFontSizeForAfterPt = getParagraphFontSize(para, defaultFontSize) * fontScale;
    prevSpaceAfterPx = resolveSpacingPx(para.properties.spaceAfter, paraFontSizeForAfterPt);
  }

  return totalHeight;
}

// ============================================================
// Text -> path conversion (Satori method)
// ============================================================

/**
 * Calculates the starting x position of the line based on alignment.
 * An alternative to text-anchor for tspan rendering.
 */
function computePathLineX(
  alignment: "l" | "ctr" | "r" | "just" | null,
  textStartX: number,
  effectiveTextWidth: number,
  width: number,
  marginRightPx: number,
  lineWidth: number,
): number {
  if (alignment === "ctr") return textStartX + (effectiveTextWidth - lineWidth) / 2;
  if (alignment === "r") return width - marginRightPx - lineWidth;
  return textStartX;
}

/**
 * Measures the total width of all segments in a row.
 * If a fontResolver is given, use the correct width with font.getAdvanceWidth().
 */
function measureLineWidth(
  segments: { text: string; properties: RunProperties }[],
  defaultFontSize: number,
  fontScale: number,
  fontResolver?: TextPathFontResolver | null,
  context: RendererContext = createLegacyRendererContext(),
): number {
  let totalWidth = 0;
  const jpanFallback = fontResolver ? getJpanFallbackFontFromContext(context) : null;
  for (const seg of segments) {
    const fontSize = (seg.properties.fontSize ?? defaultFontSize) * fontScale;
    if (fontResolver) {
      const fontSizePx = fontSize * PX_PER_PT;
      const font = fontResolver.resolveFont(
        seg.properties.fontFamily,
        seg.properties.fontFamilyEa,
        jpanFallback,
        context,
      );
      if (font) {
        totalWidth += font.getAdvanceWidth(seg.text, fontSizePx);
        continue;
      }
    }
    totalWidth += context.textMeasurer.measureTextWidth(
      seg.text,
      fontSize,
      seg.properties.bold,
      seg.properties.fontFamily,
      seg.properties.fontFamilyEa,
    );
  }
  return totalWidth;
}

/**
 * Construct the fill attribute of the path element.
 */
function buildPathFillAttrs(props: RunProperties): string {
  const attrs: string[] = [];
  if (props.color) {
    attrs.push(`fill="${props.color.hex}"`);
    if (props.color.alpha < 1) {
      attrs.push(`fill-opacity="${props.color.alpha}"`);
    }
  } else {
    attrs.push('fill="#000000"');
  }
  return attrs.join(" ");
}

/**
 * Draw underlines and strikethroughs as SVG line elements.
 */
function renderTextDecorations(
  x: number,
  y: number,
  segmentWidth: number,
  fontSizePx: number,
  props: RunProperties,
): string[] {
  const lines: string[] = [];
  const strokeColor = props.color?.hex ?? "#000000";
  const strokeWidth = Math.max(1, fontSizePx * 0.05);
  const opacityAttr =
    props.color && props.color.alpha < 1 ? ` stroke-opacity="${props.color.alpha}"` : "";

  if (props.underline) {
    const underlineY = y + fontSizePx * 0.15;
    lines.push(
      `<line x1="${x.toFixed(2)}" y1="${underlineY.toFixed(2)}" x2="${(x + segmentWidth).toFixed(2)}" y2="${underlineY.toFixed(2)}" stroke="${strokeColor}" stroke-width="${strokeWidth.toFixed(2)}"${opacityAttr}/>`,
    );
  }

  if (props.strikethrough) {
    const strikeY = y - fontSizePx * 0.3;
    lines.push(
      `<line x1="${x.toFixed(2)}" y1="${strikeY.toFixed(2)}" x2="${(x + segmentWidth).toFixed(2)}" y2="${strikeY.toFixed(2)}" stroke="${strokeColor}" stroke-width="${strokeWidth.toFixed(2)}"${opacityAttr}/>`,
    );
  }

  return lines;
}

/**
 * Renders a single text segment to a path element and returns the width.
 */
function renderSegmentAsPath(
  text: string,
  props: RunProperties,
  x: number,
  y: number,
  fontScale: number,
  defaultFontSize: number,
  fontResolver: TextPathFontResolver,
  context: RendererContext,
  vert?: TextVerticalType,
): { svg: string; width: number } {
  const fontSize = (props.fontSize ?? defaultFontSize) * fontScale;
  const fontSizePx = fontSize * PX_PER_PT;
  const parts: string[] = [];
  let totalWidth = 0;

  // Tab -> Space conversion
  const processedText = text.replace(/\t/g, "    ");

  // baseline-shift processing
  let yOffset = 0;
  if (props.baseline > 0) yOffset = -fontSizePx * 0.35;
  else if (props.baseline < 0) yOffset = fontSizePx * 0.2;
  const effectiveY = y + yOffset;

  const jpanFallback = getJpanFallbackFontFromContext(context);

  const processSegment = (
    segText: string,
    fontFamily: string | null,
    fontFamilyEa: string | null,
  ) => {
    if (segText.length === 0) return;

    const font = fontResolver.resolveFont(fontFamily, fontFamilyEa, jpanFallback, context);
    const segWidth = font
      ? font.getAdvanceWidth(segText, fontSizePx)
      : context.textMeasurer.measureTextWidth(
          segText,
          fontSize,
          props.bold,
          props.fontFamily,
          props.fontFamilyEa,
        );

    if (font) {
      const path = font.getPath(segText, x + totalWidth, effectiveY, fontSizePx);
      const pathData = path.toPathData(2);

      if (pathData && pathData.length > 0) {
        const fillAttrs = buildPathFillAttrs(props);
        parts.push(`<path d="${pathData}" ${fillAttrs}/>`);
      }
    }

    // Underline/strikethrough
    if (props.underline || props.strikethrough) {
      parts.push(...renderTextDecorations(x + totalWidth, effectiveY, segWidth, fontSizePx, props));
    }

    totalWidth += segWidth;
  };

  /**
   * CJK character upright rendering during eaVert.
   * Render each CJK character individually and with -90° counter rotation
   * Cancels the 90° CW rotation of the group to make it appear upright.
   */
  const processCjkUpright = (
    segText: string,
    fontFamily: string | null,
    fontFamilyEa: string | null,
  ) => {
    if (segText.length === 0) return;

    const font = fontResolver.resolveFont(fontFamily, fontFamilyEa, jpanFallback, context);
    const fillAttrs = buildPathFillAttrs(props);

    for (const char of segText) {
      const charWidth = font
        ? font.getAdvanceWidth(char, fontSizePx)
        : context.textMeasurer.measureTextWidth(
            char,
            fontSize,
            props.bold,
            props.fontFamily,
            props.fontFamilyEa,
          );

      if (font) {
        const charX = x + totalWidth;
        const path = font.getPath(char, charX, effectiveY, fontSizePx);
        const pathData = path.toPathData(2);

        if (pathData && pathData.length > 0) {
          // Rotation center: center point of the character
          const cx = charX + charWidth / 2;
          const cy =
            effectiveY - (fontSizePx * (font.ascender + font.descender)) / 2 / font.unitsPerEm;
          parts.push(
            `<g transform="rotate(-90, ${cx.toFixed(2)}, ${cy.toFixed(2)})"><path d="${pathData}" ${fillAttrs}/></g>`,
          );
        }
      }

      // Underline/strikethrough (placed outside the counter rotation)
      if (props.underline || props.strikethrough) {
        parts.push(
          ...renderTextDecorations(x + totalWidth, effectiveY, charWidth, fontSizePx, props),
        );
      }

      totalWidth += charWidth;
    }
  };

  // eaVert: CJK characters upright, non-CJK 90° CW with group rotation
  if (vert === "eaVert") {
    const scriptParts = splitByScript(processedText);
    for (const part of scriptParts) {
      const ff = part.isEa ? (props.fontFamilyEa ?? props.fontFamily) : props.fontFamily;
      const ffEa = part.isEa ? props.fontFamilyEa : props.fontFamilyEa;
      if (part.isEa) {
        processCjkUpright(part.text, ff, ffEa);
      } else {
        processSegment(part.text, ff, ffEa);
      }
    }
  } else if (needsScriptSplit(props)) {
    // CJK/Latin script split
    const scriptParts = splitByScript(processedText);
    for (const part of scriptParts) {
      const ff = part.isEa ? props.fontFamilyEa : props.fontFamily;
      const ffEa = part.isEa ? props.fontFamily : props.fontFamilyEa;
      processSegment(part.text, ff, ffEa);
    }
  } else {
    processSegment(processedText, props.fontFamily, props.fontFamilyEa);
  }

  let svg = parts.join("");
  if (props.hyperlink && svg.length > 0) {
    const href = escapeXml(props.hyperlink.url);
    svg = `<a href="${href}">${svg}</a>`;
  }

  return { svg, width: totalWidth };
}

/**
 * Render bullet points as paths.
 */
function renderBulletAsPath(
  bulletText: string,
  x: number,
  y: number,
  paraProps: ParagraphProperties,
  textFontSizePt: number,
  fontScale: number,
  fontResolver: TextPathFontResolver,
  context: RendererContext,
  runFontFamily?: string | null,
  runFontFamilyEa?: string | null,
): string[] {
  let bulletFontSize = textFontSizePt;
  if (paraProps.bulletSizePct !== null) {
    bulletFontSize = textFontSizePt * (paraProps.bulletSizePct / 100000);
  }
  const fontSizePx = bulletFontSize * PX_PER_PT;

  // Use bulletFont if specified, fallback to text run font if not specified
  const font = paraProps.bulletFont
    ? fontResolver.resolveFont(paraProps.bulletFont, null, undefined, context)
    : fontResolver.resolveFont(runFontFamily ?? null, runFontFamilyEa ?? null, undefined, context);
  if (!font) return [];

  const path = font.getPath(bulletText, x, y, fontSizePx);
  const pathData = path.toPathData(2);
  if (!pathData || pathData.length === 0) return [];

  const attrs: string[] = [];
  if (paraProps.bulletColor) {
    attrs.push(`fill="${paraProps.bulletColor.hex}"`);
    if (paraProps.bulletColor.alpha < 1) {
      attrs.push(`fill-opacity="${paraProps.bulletColor.alpha}"`);
    }
  } else {
    attrs.push('fill="#000000"');
  }

  return [`<path d="${pathData}" ${attrs.join(" ")}/>`];
}

/**
 * Draw text as an SVG path element (Satori method).
 * Called only if a font buffer is provided.
 */
function renderTextBodyAsPath(
  textBody: TextBody,
  transform: Transform,
  fontResolver: TextPathFontResolver,
  context: RendererContext,
): string {
  const { bodyProperties, paragraphs } = textBody;
  const originalWidth = emuToPixels(transform.extentWidth);
  const originalHeight = emuToPixels(transform.extentHeight);

  const { width, height, marginLeftPx, marginRightPx, marginTopPx, marginBottomPx } =
    resolveTextDimensions(bodyProperties, originalWidth, originalHeight);

  const hasText = paragraphs.some((p) => p.runs.some((r) => r.text.length > 0));
  if (!hasText) return "";

  const fullTextWidth = width - marginLeftPx - marginRightPx;
  const numCol = bodyProperties.numCol ?? 1;
  const textWidth = numCol > 1 ? fullTextWidth / numCol : fullTextWidth;
  const defaultFontSize = getDefaultFontSize(paragraphs);
  const shouldWrap = bodyProperties.wrap !== "none";

  let fontScale = bodyProperties.fontScale;
  const lnSpcReduction = bodyProperties.lnSpcReduction;

  if (bodyProperties.autoFit === "normAutofit" && shouldWrap) {
    const availableHeight = height - marginTopPx - marginBottomPx;
    fontScale = computeShrinkToFitScale(
      paragraphs,
      defaultFontSize,
      fontScale,
      lnSpcReduction,
      textWidth,
      availableHeight,
      context,
    );
  }

  const scaledDefaultFontSizePt = defaultFontSize * fontScale;
  const defaultLineHeightRatio = getDefaultLineHeightRatio(paragraphs, context);
  const defaultAscenderRatio = getDefaultAscenderRatio(paragraphs, context);
  const defaultNaturalHeightPt = scaledDefaultFontSizePt * defaultLineHeightRatio;

  // Vertical position calculation (reusing existing logic)
  let yStart = marginTopPx;
  const totalTextHeight = estimateTextHeight(
    paragraphs,
    defaultFontSize,
    shouldWrap,
    textWidth,
    lnSpcReduction,
    fontScale,
    context,
  );
  if (bodyProperties.anchor === "ctr") {
    yStart = Math.max(marginTopPx, (height - totalTextHeight) / 2);
  } else if (bodyProperties.anchor === "b") {
    yStart = Math.max(marginTopPx, height - totalTextHeight - marginBottomPx);
  }
  const firstParaFontSizePt = getParagraphFontSize(paragraphs[0], defaultFontSize) * fontScale;
  const firstLineBaselineOffsetPt = firstParaFontSizePt * defaultAscenderRatio;
  yStart += firstLineBaselineOffsetPt * PX_PER_PT;

  // path rendering
  const elements: string[] = [];
  let currentY = yStart;
  let isFirstLine = true;
  const autoNumCounters = new Map<string, number>();
  let prevSpaceAfterPx = 0;

  for (const para of paragraphs) {
    const paraMarginLeft = emuToPixels(para.properties.marginLeft ?? asEmu(0));
    const paraIndent = emuToPixels(para.properties.indent ?? asEmu(0));
    const textStartX = marginLeftPx + paraMarginLeft;
    const bulletX = textStartX + paraIndent;
    const effectiveTextWidth = textWidth - paraMarginLeft;

    const bulletText = resolveBulletText(para.properties, autoNumCounters);

    const paraFontSizePt = getParagraphFontSize(para, defaultFontSize) * fontScale;
    const spaceBeforePx = resolveSpacingPx(para.properties.spaceBefore, paraFontSizePt);
    const paragraphGapPx = Math.max(prevSpaceAfterPx, spaceBeforePx);

    // empty paragraph
    if (para.runs.length === 0 || !para.runs.some((r) => r.text.length > 0)) {
      if (!isFirstLine) {
        const emptyParaHeightPt = paraFontSizePt > 0 ? paraFontSizePt : defaultNaturalHeightPt;
        currentY += getLineHeightPx(para, emptyParaHeightPt, lnSpcReduction) + paragraphGapPx;
      }
      isFirstLine = false;
      prevSpaceAfterPx = resolveSpacingPx(para.properties.spaceAfter, paraFontSizePt);
      continue;
    }

    if (shouldWrap) {
      const wrappedLines = wrapParagraph(
        para,
        effectiveTextWidth,
        scaledDefaultFontSizePt,
        fontScale,
        context.textMeasurer,
      );

      for (let lineIdx = 0; lineIdx < wrappedLines.length; lineIdx++) {
        const line = wrappedLines[lineIdx];
        const lineGapPx = lineIdx === 0 ? paragraphGapPx : 0;

        if (line.segments.length === 0) {
          if (!isFirstLine) {
            currentY += getLineHeightPx(para, defaultNaturalHeightPt, lnSpcReduction) + lineGapPx;
          }
          isFirstLine = false;
          continue;
        }

        // Row height calculation and y position update
        const lineNaturalHeightPt = computeLineNaturalHeight(
          line.segments,
          defaultFontSize,
          fontScale,
          context,
        );
        if (!isFirstLine) {
          currentY += getLineHeightPx(para, lineNaturalHeightPt, lnSpcReduction) + lineGapPx;
        }

        // Calculate x position for alignment by measuring line width
        const lineWidth = measureLineWidth(
          line.segments,
          defaultFontSize,
          fontScale,
          fontResolver,
          context,
        );
        const lineStartX = computePathLineX(
          para.properties.alignment,
          textStartX,
          effectiveTextWidth,
          width,
          marginRightPx,
          lineWidth,
        );

        let currentX = lineStartX;

        // Bullet mark (first line only)
        if (lineIdx === 0 && bulletText) {
          const lineFontSize = getLineFontSize(line.segments, defaultFontSize) * fontScale;
          const firstSeg = line.segments[0];
          elements.push(
            ...renderBulletAsPath(
              bulletText,
              bulletX,
              currentY,
              para.properties,
              lineFontSize,
              fontScale,
              fontResolver,
              context,
              firstSeg?.properties.fontFamily,
              firstSeg?.properties.fontFamilyEa,
            ),
          );
        }

        // Render each segment into a path
        for (const seg of line.segments) {
          const result = renderSegmentAsPath(
            seg.text,
            seg.properties,
            currentX,
            currentY,
            fontScale,
            defaultFontSize,
            fontResolver,
            context,
            bodyProperties.vert,
          );
          if (result.svg) elements.push(result.svg);
          currentX += result.width;
        }

        isFirstLine = false;
      }
    } else {
      // wrap="none": no wrapping
      const naturalHeightPt = computeLineNaturalHeight(
        para.runs,
        defaultFontSize,
        fontScale,
        context,
      );
      if (!isFirstLine) {
        currentY += getLineHeightPx(para, naturalHeightPt, lnSpcReduction) + paragraphGapPx;
      }

      // Measure line width
      const runsAsSegments = para.runs
        .filter((r) => r.text.length > 0)
        .map((r) => ({ text: r.text, properties: r.properties }));
      const lineWidth = measureLineWidth(
        runsAsSegments,
        defaultFontSize,
        fontScale,
        fontResolver,
        context,
      );
      const lineStartX = computePathLineX(
        para.properties.alignment,
        textStartX,
        effectiveTextWidth,
        width,
        marginRightPx,
        lineWidth,
      );

      let currentX = lineStartX;

      // bullet point symbol
      if (bulletText) {
        const firstRun = para.runs.find((r) => r.text.length > 0);
        const fontSize = (firstRun?.properties.fontSize ?? defaultFontSize) * fontScale;
        elements.push(
          ...renderBulletAsPath(
            bulletText,
            bulletX,
            currentY,
            para.properties,
            fontSize,
            fontScale,
            fontResolver,
            context,
            firstRun?.properties.fontFamily,
            firstRun?.properties.fontFamilyEa,
          ),
        );
      }

      // Render each run into a pass
      for (const run of para.runs) {
        if (run.text.length === 0) continue;
        const result = renderSegmentAsPath(
          run.text,
          run.properties,
          currentX,
          currentY,
          fontScale,
          defaultFontSize,
          fontResolver,
          context,
          bodyProperties.vert,
        );
        if (result.svg) elements.push(result.svg);
        currentX += result.width;
      }

      isFirstLine = false;
    }

    prevSpaceAfterPx = resolveSpacingPx(para.properties.spaceAfter, paraFontSizePt);
  }

  const content = elements.join("");
  if (content.length === 0) return "";

  if (isVerticalText(bodyProperties.vert)) {
    return `<g transform="translate(${originalWidth}, 0) rotate(90)">${content}</g>`;
  }
  if (isVert270Text(bodyProperties.vert)) {
    return `<g transform="translate(0, ${originalHeight}) rotate(-90)">${content}</g>`;
  }
  return content;
}

/** Default tab width (number of spaces) */
const TAB_SPACES = "    "; // 4 spaces

function escapeXml(str: string): string {
  return str
    .replace(/\t/g, TAB_SPACES)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
