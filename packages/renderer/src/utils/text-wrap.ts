import type { TextMeasurer } from "../font/text-measurer.js";
import { getTextMeasurer } from "../font/text-measurer.js";
import type { Paragraph, RunProperties } from "../model/text.js";

interface LineSegment {
  text: string;
  properties: RunProperties;
}

interface WrappedLine {
  segments: LineSegment[];
}

interface Token {
  text: string;
  properties: RunProperties;
  width: number;
  breakable: boolean;
  forceBreak?: boolean;
}

const DEFAULT_FONT_SIZE = 18;

/**
 * Wrapping tolerance to absorb approximation errors in font metrics.
 * Specify as a ratio to availableWidth.
 * Width overestimation by alternative font metrics (e.g. Meiryo -> Noto Sans JP)
 * Alleviates the problem where the last character of text that would normally fit on one line is sent to the next line.
 */
const WRAP_TOLERANCE_RATIO = 0.02;

function isCjk(codePoint: number): boolean {
  return (
    (codePoint >= 0x3000 && codePoint <= 0x9fff) ||
    (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
    (codePoint >= 0xff01 && codePoint <= 0xff60) ||
    (codePoint >= 0x20000 && codePoint <= 0x2a6df)
  );
}

function isWhitespace(codePoint: number): boolean {
  return codePoint === 0x20 || codePoint === 0x09 || codePoint === 0x0a || codePoint === 0x0d;
}

/**
 * Split text into breakable units.
 * - Blank: Individual token (breakable)
 * - CJK characters: one character at a time (breakable)
 * - Sequence of Latin characters: 1 word (breakable except at the beginning)
 */
function splitTextIntoFragments(text: string): { fragment: string; breakable: boolean }[] {
  const fragments: { fragment: string; breakable: boolean }[] = [];
  let current = "";
  let currentType: "latin" | "cjk" | "space" | null = null;

  for (const char of text) {
    const cp = char.codePointAt(0)!;

    if (isWhitespace(cp)) {
      if (current && currentType !== "space") {
        fragments.push({ fragment: current, breakable: currentType === "cjk" });
        current = "";
      }
      currentType = "space";
      current += char;
    } else if (isCjk(cp)) {
      if (current) {
        fragments.push({
          fragment: current,
          breakable: currentType === "cjk" || currentType === "space",
        });
        current = "";
      }
      // Each CJK character is an independent token.
      fragments.push({ fragment: char, breakable: true });
      currentType = "cjk";
      current = "";
    } else {
      // latin letters
      if (current && currentType !== "latin") {
        fragments.push({ fragment: current, breakable: currentType === "space" });
        current = "";
      }
      currentType = "latin";
      current += char;
    }
  }

  if (current) {
    fragments.push({
      fragment: current,
      breakable: currentType === "space" || currentType === "cjk",
    });
  }

  return fragments;
}

function tokenizeRuns(
  runs: Paragraph["runs"],
  defaultFontSize: number,
  fontScale: number,
  textMeasurer: TextMeasurer,
): Token[] {
  const tokens: Token[] = [];
  let isFirst = true;

  for (const run of runs) {
    if (run.text.length === 0) continue;

    // Runs containing \n are split and forced newline tokens are inserted.
    if (run.text.includes("\n")) {
      const parts = run.text.split("\n");
      for (let pi = 0; pi < parts.length; pi++) {
        if (pi > 0) {
          // forced line break token
          tokens.push({
            text: "",
            properties: run.properties,
            width: 0,
            breakable: true,
            forceBreak: true,
          });
          isFirst = false;
        }
        const part = parts[pi];
        if (part.length === 0) continue;
        const fontSize = run.properties.fontSize
          ? run.properties.fontSize * fontScale
          : defaultFontSize;
        const bold = run.properties.bold;
        const fontFamily = run.properties.fontFamily;
        const fontFamilyEa = run.properties.fontFamilyEa;
        const fragments = splitTextIntoFragments(part);
        for (const { fragment, breakable } of fragments) {
          const width = textMeasurer.measureTextWidth(
            fragment,
            fontSize,
            bold,
            fontFamily,
            fontFamilyEa,
          );
          tokens.push({
            text: fragment,
            properties: run.properties,
            width,
            breakable: isFirst ? false : breakable,
          });
          isFirst = false;
        }
      }
      continue;
    }

    const fontSize = run.properties.fontSize
      ? run.properties.fontSize * fontScale
      : defaultFontSize;
    const bold = run.properties.bold;
    const fontFamily = run.properties.fontFamily;
    const fontFamilyEa = run.properties.fontFamilyEa;
    const fragments = splitTextIntoFragments(run.text);

    for (const { fragment, breakable } of fragments) {
      const width = textMeasurer.measureTextWidth(
        fragment,
        fontSize,
        bold,
        fontFamily,
        fontFamilyEa,
      );
      tokens.push({
        text: fragment,
        properties: run.properties,
        width,
        breakable: isFirst ? false : breakable,
      });
      isFirst = false;
    }
  }

  return tokens;
}

function isSpaceOnly(text: string): boolean {
  for (const char of text) {
    if (!isWhitespace(char.codePointAt(0)!)) return false;
  }
  return true;
}

/**
 * Forcibly split long tokens by character
 */
function splitTokenByChars(
  token: Token,
  availableWidth: number,
  defaultFontSize: number,
  fontScale: number,
  textMeasurer: TextMeasurer,
): Token[][] {
  const lines: Token[][] = [];
  let currentLine: Token[] = [];
  let currentWidth = 0;
  const fontSize = token.properties.fontSize
    ? token.properties.fontSize * fontScale
    : defaultFontSize;
  const bold = token.properties.bold;
  const fontFamily = token.properties.fontFamily;
  const fontFamilyEa = token.properties.fontFamilyEa;

  for (const char of token.text) {
    const charWidth = textMeasurer.measureTextWidth(char, fontSize, bold, fontFamily, fontFamilyEa);

    if (currentWidth + charWidth > availableWidth && currentLine.length > 0) {
      lines.push(currentLine);
      currentLine = [];
      currentWidth = 0;
    }

    currentLine.push({
      text: char,
      properties: token.properties,
      width: charWidth,
      breakable: false,
    });
    currentWidth += charWidth;
  }

  if (currentLine.length > 0) {
    lines.push(currentLine);
  }

  return lines;
}

function mergeSegments(tokens: Token[]): LineSegment[] {
  const segments: LineSegment[] = [];

  for (const token of tokens) {
    if (segments.length > 0) {
      const last = segments[segments.length - 1];
      if (last.properties === token.properties) {
        last.text += token.text;
        continue;
      }
    }
    segments.push({ text: token.text, properties: token.properties });
  }

  return segments;
}

function trimTrailingSpaces(segments: LineSegment[]): LineSegment[] {
  if (segments.length === 0) return segments;

  const last = segments[segments.length - 1];
  const trimmed = last.text.replace(/\s+$/, "");

  if (trimmed.length === 0) {
    const result = segments.slice(0, -1);
    return trimTrailingSpaces(result);
  }

  if (trimmed !== last.text) {
    return [...segments.slice(0, -1), { text: trimmed, properties: last.properties }];
  }

  return segments;
}

function layoutTokensIntoLines(
  tokens: Token[],
  availableWidth: number,
  defaultFontSize: number,
  fontScale: number,
  textMeasurer: TextMeasurer,
): WrappedLine[] {
  if (tokens.length === 0) return [{ segments: [] }];

  const lines: WrappedLine[] = [];
  let currentLine: Token[] = [];
  let currentWidth = 0;
  const tolerance = availableWidth * WRAP_TOLERANCE_RATIO;

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];

    // Force newline token: immediately end the line
    if (token.forceBreak) {
      const segments = trimTrailingSpaces(mergeSegments(currentLine));
      lines.push({ segments: segments.length > 0 ? segments : [] });
      currentLine = [];
      currentWidth = 0;
      continue;
    }

    if (currentWidth + token.width <= availableWidth + tolerance) {
      currentLine.push(token);
      currentWidth += token.width;
    } else if (currentLine.length === 0) {
      // The line is empty and doesn't even contain a single token -> Forced division by character
      if (isSpaceOnly(token.text)) {
        // Skip blank tokens
        continue;
      }
      const splitLines = splitTokenByChars(
        token,
        availableWidth,
        defaultFontSize,
        fontScale,
        textMeasurer,
      );
      for (let j = 0; j < splitLines.length; j++) {
        if (j < splitLines.length - 1) {
          const segments = trimTrailingSpaces(mergeSegments(splitLines[j]));
          if (segments.length > 0) lines.push({ segments });
        } else {
          // the last chunk becomes the beginning of the next line
          currentLine = splitLines[j];
          currentWidth = splitLines[j].reduce((sum, t) => sum + t.width, 0);
        }
      }
    } else if (token.breakable) {
      // Line breaks at possible line breaks
      const segments = trimTrailingSpaces(mergeSegments(currentLine));
      if (segments.length > 0) lines.push({ segments });

      if (isSpaceOnly(token.text)) {
        // Skip leading spaces
        currentLine = [];
        currentWidth = 0;
      } else {
        currentLine = [token];
        currentWidth = token.width;
      }
    } else {
      // Not breakable but does not fit on the line -> Break the line in front and move this token to the next line
      const segments = trimTrailingSpaces(mergeSegments(currentLine));
      if (segments.length > 0) lines.push({ segments });
      currentLine = [token];
      currentWidth = token.width;
    }
  }

  if (currentLine.length > 0) {
    const segments = trimTrailingSpaces(mergeSegments(currentLine));
    if (segments.length > 0) lines.push({ segments });
  }

  return lines.length > 0 ? lines : [{ segments: [] }];
}

/**
 * Convert a run of paragraphs to a wrapped line array
 */
export function wrapParagraph(
  paragraph: Paragraph,
  availableWidth: number,
  defaultFontSize: number = DEFAULT_FONT_SIZE,
  fontScale: number = 1,
  textMeasurer: TextMeasurer = getTextMeasurer(),
): WrappedLine[] {
  if (paragraph.runs.length === 0 || !paragraph.runs.some((r) => r.text.length > 0)) {
    return [{ segments: [] }];
  }

  const safeWidth = Math.max(availableWidth, 1);
  const tokens = tokenizeRuns(paragraph.runs, defaultFontSize, fontScale, textMeasurer);

  if (tokens.length === 0) return [{ segments: [] }];

  return layoutTokensIntoLines(tokens, safeWidth, defaultFontSize, fontScale, textMeasurer);
}
