import type { Paragraph, RunProperties } from "../model/text.js";
import { measureTextWidth } from "./text-measure.js";

export interface LineSegment {
  text: string;
  properties: RunProperties;
}

export interface WrappedLine {
  segments: LineSegment[];
}

interface Token {
  text: string;
  properties: RunProperties;
  width: number;
  breakable: boolean;
}

const DEFAULT_FONT_SIZE = 18;

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
 * テキストをブレーク可能な単位に分割する。
 * - 空白: 個別トークン (breakable)
 * - CJK 文字: 1文字ずつ (breakable)
 * - ラテン文字の連続: 1ワード (先頭以外で breakable)
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
      // CJK 文字は1文字ずつ独立トークン
      fragments.push({ fragment: char, breakable: true });
      currentType = "cjk";
      current = "";
    } else {
      // ラテン文字
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

function tokenizeRuns(runs: Paragraph["runs"], defaultFontSize: number): Token[] {
  const tokens: Token[] = [];
  let isFirst = true;

  for (const run of runs) {
    if (run.text.length === 0) continue;

    const fontSize = run.properties.fontSize ?? defaultFontSize;
    const bold = run.properties.bold;
    const fontFamily = run.properties.fontFamily;
    const fontFamilyEa = run.properties.fontFamilyEa;
    const fragments = splitTextIntoFragments(run.text);

    for (const { fragment, breakable } of fragments) {
      const width = measureTextWidth(fragment, fontSize, bold, fontFamily, fontFamilyEa);
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
 * 長いトークンを文字単位で強制分割する
 */
function splitTokenByChars(
  token: Token,
  availableWidth: number,
  defaultFontSize: number,
): Token[][] {
  const lines: Token[][] = [];
  let currentLine: Token[] = [];
  let currentWidth = 0;
  const fontSize = token.properties.fontSize ?? defaultFontSize;
  const bold = token.properties.bold;
  const fontFamily = token.properties.fontFamily;
  const fontFamilyEa = token.properties.fontFamilyEa;

  for (const char of token.text) {
    const charWidth = measureTextWidth(char, fontSize, bold, fontFamily, fontFamilyEa);

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
): WrappedLine[] {
  if (tokens.length === 0) return [{ segments: [] }];

  const lines: WrappedLine[] = [];
  let currentLine: Token[] = [];
  let currentWidth = 0;

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];

    if (currentWidth + token.width <= availableWidth) {
      currentLine.push(token);
      currentWidth += token.width;
    } else if (currentLine.length === 0) {
      // 行が空で1トークンすら入らない → 文字単位で強制分割
      if (isSpaceOnly(token.text)) {
        // 空白だけのトークンはスキップ
        continue;
      }
      const splitLines = splitTokenByChars(token, availableWidth, defaultFontSize);
      for (let j = 0; j < splitLines.length; j++) {
        if (j < splitLines.length - 1) {
          const segments = trimTrailingSpaces(mergeSegments(splitLines[j]));
          if (segments.length > 0) lines.push({ segments });
        } else {
          // 最後のチャンクは次の行の先頭になる
          currentLine = splitLines[j];
          currentWidth = splitLines[j].reduce((sum, t) => sum + t.width, 0);
        }
      }
    } else if (token.breakable) {
      // 改行可能な位置で改行
      const segments = trimTrailingSpaces(mergeSegments(currentLine));
      if (segments.length > 0) lines.push({ segments });

      if (isSpaceOnly(token.text)) {
        // 行頭の空白はスキップ
        currentLine = [];
        currentWidth = 0;
      } else {
        currentLine = [token];
        currentWidth = token.width;
      }
    } else {
      // breakable でないが行に収まらない → 手前で改行してこのトークンを次の行へ
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
 * 段落のランを折り返して行配列に変換する
 */
export function wrapParagraph(
  paragraph: Paragraph,
  availableWidth: number,
  defaultFontSize: number = DEFAULT_FONT_SIZE,
): WrappedLine[] {
  if (paragraph.runs.length === 0 || !paragraph.runs.some((r) => r.text.length > 0)) {
    return [{ segments: [] }];
  }

  const safeWidth = Math.max(availableWidth, 1);
  const tokens = tokenizeRuns(paragraph.runs, defaultFontSize);

  if (tokens.length === 0) return [{ segments: [] }];

  return layoutTokensIntoLines(tokens, safeWidth, defaultFontSize);
}
