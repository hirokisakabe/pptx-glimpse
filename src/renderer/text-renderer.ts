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
import type { Transform } from "../model/shape.js";
import { EMU_PER_INCH } from "../utils/constants.js";
import { emuToPixels } from "../utils/emu.js";
import { wrapParagraph } from "../utils/text-wrap.js";
import { getMetricsFallbackFont } from "../data/font-metrics.js";
import { getTextMeasurer } from "../font/text-measurer.js";
import { getCurrentMappedFont } from "../font/font-mapping-context.js";
import type { TextPathFontResolver } from "../font/text-path-context.js";
import { getTextPathFontResolver } from "../font/text-path-context.js";

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
 * 縦書き時の次元・マージン入れ替えを行う。
 * 回転変換後にテキストが正しい位置に表示されるよう、レイアウト空間を構成する。
 */
function resolveTextDimensions(
  bodyProperties: BodyProperties,
  originalWidth: number,
  originalHeight: number,
): {
  width: number;
  height: number;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
} {
  const vert = bodyProperties.vert;

  if (isVerticalText(vert)) {
    // vert (90° CW): レイアウト空間は H×W、マージンを回転に合わせて入れ替え
    return {
      width: originalHeight,
      height: originalWidth,
      marginLeft: emuToPixels(bodyProperties.marginTop),
      marginRight: emuToPixels(bodyProperties.marginBottom),
      marginTop: emuToPixels(bodyProperties.marginRight),
      marginBottom: emuToPixels(bodyProperties.marginLeft),
    };
  }

  if (isVert270Text(vert)) {
    // vert270 (90° CCW): レイアウト空間は H×W、マージンを逆方向に入れ替え
    return {
      width: originalHeight,
      height: originalWidth,
      marginLeft: emuToPixels(bodyProperties.marginBottom),
      marginRight: emuToPixels(bodyProperties.marginTop),
      marginTop: emuToPixels(bodyProperties.marginLeft),
      marginBottom: emuToPixels(bodyProperties.marginRight),
    };
  }

  // 水平テキスト
  return {
    width: originalWidth,
    height: originalHeight,
    marginLeft: emuToPixels(bodyProperties.marginLeft),
    marginRight: emuToPixels(bodyProperties.marginRight),
    marginTop: emuToPixels(bodyProperties.marginTop),
    marginBottom: emuToPixels(bodyProperties.marginBottom),
  };
}

export function renderTextBody(textBody: TextBody, transform: Transform): string {
  const fontResolver = getTextPathFontResolver();
  if (fontResolver) {
    return renderTextBodyAsPath(textBody, transform, fontResolver);
  }

  const { bodyProperties, paragraphs } = textBody;
  const originalWidth = emuToPixels(transform.extentWidth);
  const originalHeight = emuToPixels(transform.extentHeight);

  const { width, height, marginLeft, marginRight, marginTop, marginBottom } = resolveTextDimensions(
    bodyProperties,
    originalWidth,
    originalHeight,
  );

  const hasText = paragraphs.some((p) => p.runs.some((r) => r.text.length > 0));
  if (!hasText) return "";

  const fullTextWidth = width - marginLeft - marginRight;
  const numCol = bodyProperties.numCol ?? 1;
  const textWidth = numCol > 1 ? fullTextWidth / numCol : fullTextWidth;
  const defaultFontSize = getDefaultFontSize(paragraphs);
  const shouldWrap = bodyProperties.wrap !== "none";

  let fontScale = bodyProperties.fontScale;
  const lnSpcReduction = bodyProperties.lnSpcReduction;

  // normAutofit: テキストが図形からはみ出す場合に fontScale を動的に縮小
  if (bodyProperties.autoFit === "normAutofit" && shouldWrap) {
    const availableHeight = height - marginTop - marginBottom;
    fontScale = computeShrinkToFitScale(
      paragraphs,
      defaultFontSize,
      fontScale,
      lnSpcReduction,
      textWidth,
      availableHeight,
    );
  }

  const scaledDefaultFontSize = defaultFontSize * fontScale;

  // デフォルトフォントの行高さ比率
  const defaultLineHeightRatio = getDefaultLineHeightRatio(paragraphs);
  const defaultNaturalHeight = scaledDefaultFontSize * defaultLineHeightRatio;

  const tspans: string[] = [];
  let isFirstLine = true;

  // 連番管理用
  const autoNumCounters = new Map<string, number>();

  // 前の段落の spaceAfter（ピクセル解決済み）
  let prevSpaceAfterPx = 0;

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

    // 段落間隔の計算: max(前段落のspaceAfter, 現段落のspaceBefore)
    const paraFontSize = getParagraphFontSize(para, defaultFontSize) * fontScale;
    const spaceBeforePx = resolveSpacingPx(para.properties.spaceBefore, paraFontSize);
    const paragraphGap = Math.max(prevSpaceAfterPx, spaceBeforePx);

    if (para.runs.length === 0 || !para.runs.some((r) => r.text.length > 0)) {
      const emptyParaHeight = paraFontSize > 0 ? paraFontSize : defaultNaturalHeight;
      const dy = computeDy(isFirstLine, emptyParaHeight, DEFAULT_LINE_SPACING, paragraphGap);
      tspans.push(`<tspan x="${xPos}" dy="${dy}" text-anchor="${anchorValue}"> </tspan>`);
      isFirstLine = false;
      prevSpaceAfterPx = resolveSpacingPx(para.properties.spaceAfter, paraFontSize);
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
        const lineGap = lineIdx === 0 ? paragraphGap : 0;
        if (line.segments.length === 0) {
          const dy = computeDy(
            isFirstLine,
            defaultNaturalHeight,
            getLineSpacing(para, lnSpcReduction),
            lineGap,
          );
          tspans.push(`<tspan x="${xPos}" dy="${dy}" text-anchor="${anchorValue}"> </tspan>`);
          isFirstLine = false;
          continue;
        }

        // 最初の行に箇条書き記号を挿入
        if (lineIdx === 0 && bulletText) {
          const lineFontSize = getLineFontSize(line.segments, defaultFontSize) * fontScale;
          const lineNaturalHeight = computeLineNaturalHeight(
            line.segments,
            defaultFontSize,
            fontScale,
          );
          const dy = computeDy(
            isFirstLine,
            lineNaturalHeight,
            getLineSpacing(para, lnSpcReduction),
            paragraphGap,
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
              const lineNaturalHeight = computeLineNaturalHeight(
                line.segments,
                defaultFontSize,
                fontScale,
              );
              const dy = computeDy(
                isFirstLine,
                lineNaturalHeight,
                getLineSpacing(para, lnSpcReduction),
                lineGap,
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
        const naturalHeight = computeLineNaturalHeight(para.runs, defaultFontSize, fontScale);
        const dy = computeDy(
          isFirstLine,
          naturalHeight,
          getLineSpacing(para, lnSpcReduction),
          paragraphGap,
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
            const naturalHeight = computeLineNaturalHeight(para.runs, defaultFontSize, fontScale);
            const dy = computeDy(
              isFirstLine,
              naturalHeight,
              getLineSpacing(para, lnSpcReduction),
              paragraphGap,
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

    prevSpaceAfterPx = resolveSpacingPx(para.properties.spaceAfter, paraFontSize);
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
  const firstParaFontSize = getParagraphFontSize(paragraphs[0], defaultFontSize) * fontScale;
  const firstLineNaturalHeight = firstParaFontSize * defaultLineHeightRatio;
  yStart += firstLineNaturalHeight * PX_PER_PT;

  const textElement = `<text x="0" y="${yStart}">${tspans.join("")}</text>`;

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

function resolveSpacingPx(spacing: SpacingValue, fontSizePt: number): number {
  if (spacing.type === "pts") {
    return (spacing.value / 100) * PX_PER_PT;
  }
  // pct: val / 100000 がフォントサイズに対する比率
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

function computeDy(
  isFirstLine: boolean,
  fontSizePt: number,
  lineSpacingFactor: number,
  paragraphGap: number,
): string {
  if (isFirstLine) return "0";

  const lineHeight = fontSizePt * PX_PER_PT * lineSpacingFactor;
  const dy = lineHeight + paragraphGap;

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

function computeLineNaturalHeight(
  segments: { properties: RunProperties }[],
  defaultFontSize: number,
  fontScale: number,
): number {
  let maxHeight = 0;
  for (const seg of segments) {
    const fontSize = (seg.properties.fontSize ?? defaultFontSize) * fontScale;
    const ratio = getTextMeasurer().getLineHeightRatio(
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

export function buildFontFamilyValue(fonts: (string | null)[]): string | null {
  const uniqueFonts: string[] = [];
  const seen = new Set<string>();

  for (const font of fonts) {
    if (font && !seen.has(font)) {
      seen.add(font);
      uniqueFonts.push(font);

      // マッピングテーブルから OSS 代替フォントを追加
      const mapped = getCurrentMappedFont(font);
      if (mapped && !seen.has(mapped)) {
        seen.add(mapped);
        uniqueFonts.push(mapped);
      }

      // メトリクス互換 OSS フォントをフォールバックとして追加
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
): string {
  const styles: string[] = [];

  if (props.fontSize) {
    const scaledSize = props.fontSize * fontScale;
    styles.push(`font-size="${scaledSize}pt"`);
  }
  const fonts = fontFamilies ?? [props.fontFamily, props.fontFamilyEa];
  const fontFamilyValue = buildFontFamilyValue(fonts);
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
): string {
  let tspanContent: string;
  if (!needsScriptSplit(props)) {
    const styles = buildStyleAttrs(props, fontScale);
    tspanContent = `<tspan ${prefix}${styles}>${escapeXml(text)}</tspan>`;
  } else {
    const parts = splitByScript(text);
    const result: string[] = [];
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i];
      const fonts = part.isEa
        ? [props.fontFamilyEa, props.fontFamily]
        : [props.fontFamily, props.fontFamilyEa];
      const styles = buildStyleAttrs(props, fontScale, fonts);
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

function getDefaultLineHeightRatio(paragraphs: TextBody["paragraphs"]): number {
  for (const p of paragraphs) {
    for (const r of p.runs) {
      if (r.properties.fontFamily || r.properties.fontFamilyEa) {
        return getTextMeasurer().getLineHeightRatio(
          r.properties.fontFamily,
          r.properties.fontFamilyEa,
        );
      }
    }
  }
  return 1.2;
}

/**
 * spAutofit: テキスト量に応じた必要な図形の高さ (EMU) を計算する。
 * テキストが元の図形に収まる場合は null を返す。
 */
export function computeSpAutofitHeight(textBody: TextBody, transform: Transform): number | null {
  const { bodyProperties, paragraphs } = textBody;

  const hasText = paragraphs.some((p) => p.runs.some((r) => r.text.length > 0));
  if (!hasText) return null;

  const originalWidth = emuToPixels(transform.extentWidth);
  const originalHeight = emuToPixels(transform.extentHeight);

  const { width, height, marginLeft, marginRight, marginTop, marginBottom } = resolveTextDimensions(
    bodyProperties,
    originalWidth,
    originalHeight,
  );

  const fullTextWidth = width - marginLeft - marginRight;
  const numCol = bodyProperties.numCol ?? 1;
  const textWidth = numCol > 1 ? fullTextWidth / numCol : fullTextWidth;
  const defaultFontSize = getDefaultFontSize(paragraphs);
  const shouldWrap = bodyProperties.wrap !== "none";

  const textHeight = estimateTextHeight(paragraphs, defaultFontSize, shouldWrap, textWidth);
  const requiredHeightPx = textHeight + marginTop + marginBottom;

  if (requiredHeightPx <= height) return null;

  const DEFAULT_DPI = 96;
  return (requiredHeightPx / DEFAULT_DPI) * EMU_PER_INCH;
}

function computeShrinkToFitScale(
  paragraphs: TextBody["paragraphs"],
  defaultFontSize: number,
  fontScale: number,
  lnSpcReduction: number,
  textWidth: number,
  availableHeight: number,
): number {
  if (availableHeight <= 0) return fontScale;

  const minScale = fontScale * 0.1;
  let scale = fontScale;
  const maxIterations = 5;

  for (let i = 0; i < maxIterations; i++) {
    const scaledDefault = defaultFontSize * scale;
    const textHeight = estimateTextHeight(
      paragraphs,
      scaledDefault,
      true,
      textWidth,
      lnSpcReduction,
      scale,
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
): number {
  let totalHeight = 0;
  const defaultRatio = getDefaultLineHeightRatio(paragraphs);
  let prevSpaceAfterPx = 0;

  for (let pIdx = 0; pIdx < paragraphs.length; pIdx++) {
    const para = paragraphs[pIdx];
    const lineSpacing = getLineSpacing(para, lnSpcReduction);
    const isEmpty = !para.runs.some((r) => r.text.length > 0);
    const naturalHeight =
      isEmpty && para.endParaRunProperties?.fontSize
        ? para.endParaRunProperties.fontSize * fontScale * defaultRatio
        : computeLineNaturalHeight(para.runs, defaultFontSize, fontScale);
    const lineHeight =
      (naturalHeight > 0 ? naturalHeight : defaultFontSize * fontScale * defaultRatio) *
      PX_PER_PT *
      lineSpacing;

    let lineCount: number;
    if (shouldWrap && para.runs.length > 0 && para.runs.some((r) => r.text.length > 0)) {
      const wrappedLines = wrapParagraph(para, textWidth, defaultFontSize, fontScale);
      lineCount = wrappedLines.length;
    } else {
      lineCount = para.runs.some((r) => r.text.length > 0) ? 1 : 1;
    }

    totalHeight += lineCount * lineHeight;

    if (pIdx > 0) {
      const paraFontSize = getParagraphFontSize(para, defaultFontSize) * fontScale;
      const spaceBeforePx = resolveSpacingPx(para.properties.spaceBefore, paraFontSize);
      totalHeight += Math.max(prevSpaceAfterPx, spaceBeforePx);
    }

    const paraFontSizeForAfter = getParagraphFontSize(para, defaultFontSize) * fontScale;
    prevSpaceAfterPx = resolveSpacingPx(para.properties.spaceAfter, paraFontSizeForAfter);
  }

  return totalHeight;
}

// ============================================================
// テキスト→パス変換 (Satori 方式)
// ============================================================

/**
 * alignment に基づいて行の開始 x 位置を計算する。
 * tspan レンダリングの text-anchor に代わるもの。
 */
function computePathLineX(
  alignment: "l" | "ctr" | "r" | "just",
  textStartX: number,
  effectiveTextWidth: number,
  width: number,
  marginRight: number,
  lineWidth: number,
): number {
  switch (alignment) {
    case "ctr":
      return textStartX + (effectiveTextWidth - lineWidth) / 2;
    case "r":
      return width - marginRight - lineWidth;
    default:
      return textStartX;
  }
}

/**
 * 行内の全セグメントの幅合計を計測する。
 */
function measureLineWidth(
  segments: { text: string; properties: RunProperties }[],
  defaultFontSize: number,
  fontScale: number,
): number {
  let totalWidth = 0;
  for (const seg of segments) {
    const fontSize = (seg.properties.fontSize ?? defaultFontSize) * fontScale;
    totalWidth += getTextMeasurer().measureTextWidth(
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
 * path 要素の fill 属性を構築する。
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
 * 下線・取り消し線を SVG line 要素として描画する。
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
 * 単一テキストセグメントを path 要素にレンダリングし、幅を返す。
 */
function renderSegmentAsPath(
  text: string,
  props: RunProperties,
  x: number,
  y: number,
  fontScale: number,
  defaultFontSize: number,
  fontResolver: TextPathFontResolver,
): { svg: string; width: number } {
  const fontSize = (props.fontSize ?? defaultFontSize) * fontScale;
  const fontSizePx = fontSize * PX_PER_PT;
  const parts: string[] = [];
  let totalWidth = 0;

  // タブ→スペース変換
  const processedText = text.replace(/\t/g, "    ");

  // baseline-shift 処理
  let yOffset = 0;
  if (props.baseline > 0) yOffset = -fontSizePx * 0.35;
  else if (props.baseline < 0) yOffset = fontSizePx * 0.2;
  const effectiveY = y + yOffset;

  const processSegment = (
    segText: string,
    fontFamily: string | null,
    fontFamilyEa: string | null,
  ) => {
    if (segText.length === 0) return;

    const font = fontResolver.resolveFont(fontFamily, fontFamilyEa);
    const segWidth = getTextMeasurer().measureTextWidth(
      segText,
      fontSize,
      props.bold,
      fontFamily,
      fontFamilyEa,
    );

    if (font) {
      const path = font.getPath(segText, x + totalWidth, effectiveY, fontSizePx);
      const pathData = path.toPathData(2);

      if (pathData && pathData.length > 0) {
        const fillAttrs = buildPathFillAttrs(props);
        parts.push(`<path d="${pathData}" ${fillAttrs}/>`);
      }
    }

    // 下線・取り消し線
    if (props.underline || props.strikethrough) {
      parts.push(...renderTextDecorations(x + totalWidth, effectiveY, segWidth, fontSizePx, props));
    }

    totalWidth += segWidth;
  };

  // CJK/ラテンのスクリプト分割
  if (needsScriptSplit(props)) {
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
 * 箇条書き記号をパスとしてレンダリングする。
 */
function renderBulletAsPath(
  bulletText: string,
  x: number,
  y: number,
  paraProps: ParagraphProperties,
  textFontSizePt: number,
  fontScale: number,
  fontResolver: TextPathFontResolver,
): string[] {
  let bulletFontSize = textFontSizePt;
  if (paraProps.bulletSizePct !== null) {
    bulletFontSize = textFontSizePt * (paraProps.bulletSizePct / 100000);
  }
  const fontSizePx = bulletFontSize * PX_PER_PT;

  const font = fontResolver.resolveFont(paraProps.bulletFont, null);
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
 * テキストを SVG path 要素として描画する（Satori 方式）。
 * フォントバッファが提供されている場合にのみ呼ばれる。
 */
function renderTextBodyAsPath(
  textBody: TextBody,
  transform: Transform,
  fontResolver: TextPathFontResolver,
): string {
  const { bodyProperties, paragraphs } = textBody;
  const originalWidth = emuToPixels(transform.extentWidth);
  const originalHeight = emuToPixels(transform.extentHeight);

  const { width, height, marginLeft, marginRight, marginTop, marginBottom } = resolveTextDimensions(
    bodyProperties,
    originalWidth,
    originalHeight,
  );

  const hasText = paragraphs.some((p) => p.runs.some((r) => r.text.length > 0));
  if (!hasText) return "";

  const fullTextWidth = width - marginLeft - marginRight;
  const numCol = bodyProperties.numCol ?? 1;
  const textWidth = numCol > 1 ? fullTextWidth / numCol : fullTextWidth;
  const defaultFontSize = getDefaultFontSize(paragraphs);
  const shouldWrap = bodyProperties.wrap !== "none";

  let fontScale = bodyProperties.fontScale;
  const lnSpcReduction = bodyProperties.lnSpcReduction;

  if (bodyProperties.autoFit === "normAutofit" && shouldWrap) {
    const availableHeight = height - marginTop - marginBottom;
    fontScale = computeShrinkToFitScale(
      paragraphs,
      defaultFontSize,
      fontScale,
      lnSpcReduction,
      textWidth,
      availableHeight,
    );
  }

  const scaledDefaultFontSize = defaultFontSize * fontScale;
  const defaultLineHeightRatio = getDefaultLineHeightRatio(paragraphs);
  const defaultNaturalHeight = scaledDefaultFontSize * defaultLineHeightRatio;

  // 垂直位置の計算（既存ロジック再利用）
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
  const firstParaFontSize = getParagraphFontSize(paragraphs[0], defaultFontSize) * fontScale;
  const firstLineNaturalHeight = firstParaFontSize * defaultLineHeightRatio;
  yStart += firstLineNaturalHeight * PX_PER_PT;

  // パスレンダリング
  const elements: string[] = [];
  let currentY = yStart;
  let isFirstLine = true;
  const autoNumCounters = new Map<string, number>();
  let prevSpaceAfterPx = 0;

  for (const para of paragraphs) {
    const paraMarginLeft = emuToPixels(para.properties.marginLeft);
    const paraIndent = emuToPixels(para.properties.indent);
    const textStartX = marginLeft + paraMarginLeft;
    const bulletX = textStartX + paraIndent;
    const effectiveTextWidth = textWidth - paraMarginLeft;

    const bulletText = resolveBulletText(para.properties, autoNumCounters);

    const paraFontSize = getParagraphFontSize(para, defaultFontSize) * fontScale;
    const spaceBeforePx = resolveSpacingPx(para.properties.spaceBefore, paraFontSize);
    const paragraphGap = Math.max(prevSpaceAfterPx, spaceBeforePx);

    // 空段落
    if (para.runs.length === 0 || !para.runs.some((r) => r.text.length > 0)) {
      if (!isFirstLine) {
        const emptyParaHeight = paraFontSize > 0 ? paraFontSize : defaultNaturalHeight;
        currentY += emptyParaHeight * PX_PER_PT * DEFAULT_LINE_SPACING + paragraphGap;
      }
      isFirstLine = false;
      prevSpaceAfterPx = resolveSpacingPx(para.properties.spaceAfter, paraFontSize);
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
        const lineGap = lineIdx === 0 ? paragraphGap : 0;

        if (line.segments.length === 0) {
          if (!isFirstLine) {
            currentY +=
              defaultNaturalHeight * PX_PER_PT * getLineSpacing(para, lnSpcReduction) + lineGap;
          }
          isFirstLine = false;
          continue;
        }

        // 行の高さ計算と y 位置更新
        const lineNaturalHeight = computeLineNaturalHeight(
          line.segments,
          defaultFontSize,
          fontScale,
        );
        if (!isFirstLine) {
          currentY +=
            lineNaturalHeight * PX_PER_PT * getLineSpacing(para, lnSpcReduction) + lineGap;
        }

        // 行の幅を計測して alignment 用の x 位置を計算
        const lineWidth = measureLineWidth(line.segments, defaultFontSize, fontScale);
        const lineStartX = computePathLineX(
          para.properties.alignment,
          textStartX,
          effectiveTextWidth,
          width,
          marginRight,
          lineWidth,
        );

        let currentX = lineStartX;

        // 箇条書き記号（最初の行のみ）
        if (lineIdx === 0 && bulletText) {
          const lineFontSize = getLineFontSize(line.segments, defaultFontSize) * fontScale;
          elements.push(
            ...renderBulletAsPath(
              bulletText,
              bulletX,
              currentY,
              para.properties,
              lineFontSize,
              fontScale,
              fontResolver,
            ),
          );
        }

        // 各セグメントをパスにレンダリング
        for (const seg of line.segments) {
          const result = renderSegmentAsPath(
            seg.text,
            seg.properties,
            currentX,
            currentY,
            fontScale,
            defaultFontSize,
            fontResolver,
          );
          if (result.svg) elements.push(result.svg);
          currentX += result.width;
        }

        isFirstLine = false;
      }
    } else {
      // wrap="none": 折り返しなし
      const naturalHeight = computeLineNaturalHeight(para.runs, defaultFontSize, fontScale);
      if (!isFirstLine) {
        currentY += naturalHeight * PX_PER_PT * getLineSpacing(para, lnSpcReduction) + paragraphGap;
      }

      // 行の幅を計測
      const runsAsSegments = para.runs
        .filter((r) => r.text.length > 0)
        .map((r) => ({ text: r.text, properties: r.properties }));
      const lineWidth = measureLineWidth(runsAsSegments, defaultFontSize, fontScale);
      const lineStartX = computePathLineX(
        para.properties.alignment,
        textStartX,
        effectiveTextWidth,
        width,
        marginRight,
        lineWidth,
      );

      let currentX = lineStartX;

      // 箇条書き記号
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
          ),
        );
      }

      // 各ランをパスにレンダリング
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
        );
        if (result.svg) elements.push(result.svg);
        currentX += result.width;
      }

      isFirstLine = false;
    }

    prevSpaceAfterPx = resolveSpacingPx(para.properties.spaceAfter, paraFontSize);
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

/** デフォルトタブ幅（スペース数） */
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
