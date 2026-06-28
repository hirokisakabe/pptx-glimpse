import { describe, expect, it } from "vitest";

import type { Paragraph, RunProperties } from "../model/text.js";
import { wrapParagraph } from "./text-wrap.js";

function makeRunProps(overrides: Partial<RunProperties> = {}): RunProperties {
  return {
    fontSize: 18,
    fontFamily: null,
    fontFamilyEa: null,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    color: null,
    baseline: 0,
    ...overrides,
  };
}

function makeParagraph(texts: string[], props?: Partial<RunProperties>): Paragraph {
  return {
    runs: texts.map((text) => ({ text, properties: makeRunProps(props) })),
    properties: {
      alignment: "l",
      lineSpacing: null,
      spaceBefore: 0,
      spaceAfter: 0,
      level: 0,
    },
  };
}

describe("wrapParagraph", () => {
  it("Do not wrap short text", () => {
    const para = makeParagraph(["Hi"]);
    const lines = wrapParagraph(para, 500, 18);
    expect(lines).toHaveLength(1);
    expect(lines[0].segments).toHaveLength(1);
    expect(lines[0].segments[0].text).toBe("Hi");
  });

  it("Wrap long English text at word boundaries", () => {
    const para = makeParagraph(["The quick brown fox jumps over the lazy dog"]);
    // Narrow the width to cause wrapping
    const lines = wrapParagraph(para, 100, 18);
    expect(lines.length).toBeGreaterThan(1);
    // Each line starts with a word in uppercase or lowercase
    for (const line of lines) {
      expect(line.segments.length).toBeGreaterThan(0);
    }
  });

  it("Wrap long CJK text at character boundaries", () => {
    const para = makeParagraph(["本日は晴天なり今日もいい天気です"]);
    const lines = wrapParagraph(para, 100, 18);
    expect(lines.length).toBeGreaterThan(1);
  });

  it("Empty paragraphs return empty lines", () => {
    const para = makeParagraph([]);
    const lines = wrapParagraph(para, 500, 18);
    expect(lines).toHaveLength(1);
    expect(lines[0].segments).toHaveLength(0);
  });

  it("Paragraphs with only runs of empty text return empty lines", () => {
    const para = makeParagraph([""]);
    const lines = wrapParagraph(para, 500, 18);
    expect(lines).toHaveLength(1);
    expect(lines[0].segments).toHaveLength(0);
  });

  it("Multiple TextRuns are correctly segmented", () => {
    const boldProps = makeRunProps({ bold: true });
    const normalProps = makeRunProps({ bold: false });
    const para: Paragraph = {
      runs: [
        { text: "Bold ", properties: boldProps },
        { text: "Normal", properties: normalProps },
      ],
      properties: {
        alignment: "l",
        lineSpacing: null,
        spaceBefore: 0,
        spaceAfter: 0,
        level: 0,
      },
    };
    const lines = wrapParagraph(para, 500, 18);
    expect(lines).toHaveLength(1);
    // bold and normal result in at least two segments or are not merged
    const allText = lines[0].segments.map((s) => s.text).join("");
    expect(allText).toBe("Bold Normal");
  });

  it("Wrapping across TextRun boundaries with different styles", () => {
    const boldProps = makeRunProps({ bold: true });
    const normalProps = makeRunProps({ bold: false });
    const para: Paragraph = {
      runs: [
        { text: "First part is bold and ", properties: boldProps },
        { text: "second part is normal text", properties: normalProps },
      ],
      properties: {
        alignment: "l",
        lineSpacing: null,
        spaceBefore: 0,
        spaceAfter: 0,
        level: 0,
      },
    };
    const lines = wrapParagraph(para, 150, 18);
    expect(lines.length).toBeGreaterThan(1);
    // Join each line with a space to ensure all text is preserved
    // (The spaces at the end of the line are trimmed, so the spaces between the lines are supplemented and combined)
    const allText = lines.map((l) => l.segments.map((s) => s.text).join("")).join(" ");
    expect(allText).toContain("First part");
    expect(allText).toContain("bold and");
    expect(allText).toContain("second part");
    expect(allText).toContain("normal text");
  });

  it("Place at least one character even if availableWidth is very small", () => {
    const para = makeParagraph(["AB"]);
    const lines = wrapParagraph(para, 1, 18);
    expect(lines.length).toBeGreaterThanOrEqual(1);
    const allText = lines.flatMap((l) => l.segments.map((s) => s.text)).join("");
    expect(allText).toBe("AB");
  });

  it("Whitespace at the end of the line is trimmed", () => {
    const para = makeParagraph(["Hello World"]);
    // Wrap at Hello(~60px) + space(~7px) = ~67px
    const lines = wrapParagraph(para, 80, 18);
    expect(lines.length).toBeGreaterThanOrEqual(1);
    for (const line of lines) {
      const lastSeg = line.segments[line.segments.length - 1];
      if (lastSeg) {
        expect(lastSeg.text).not.toMatch(/\s+$/);
      }
    }
  });

  it("fontScale is applied to run-specific fontSize", () => {
    // Set explicit fontSize=36 in run and measure width with fontScale=0.5
    // After applying fontScale, it will be equivalent to fontSize=18, so the wrapping result should be the same as fontSize=18.
    const para = makeParagraph(["The quick brown fox jumps over the lazy dog"], {
      fontSize: 36,
    });
    const linesWithScale = wrapParagraph(para, 200, 18, 0.5);
    const paraSmall = makeParagraph(["The quick brown fox jumps over the lazy dog"], {
      fontSize: 18,
    });
    const linesSmall = wrapParagraph(paraSmall, 200, 18, 1);
    expect(linesWithScale.length).toBe(linesSmall.length);
  });

  it("Do not wrap CJK text that slightly protrudes due to font metrics error", () => {
    // Measurement width of "ABC system" ≈ 135.74px (Carlito ABC + NotoSansJP CJKx4, 18pt)
    // availableWidth=134 -> Extrusion 1.74px, tolerance=134*0.02=2.68px -> Fits in one line
    const props = makeRunProps({
      fontFamily: "Calibri",
      fontFamilyEa: "Meiryo",
      fontSize: 18,
    });
    const para: Paragraph = {
      runs: [{ text: "ABCシステム", properties: props }],
      properties: {
        alignment: "l",
        lineSpacing: null,
        spaceBefore: 0,
        spaceAfter: 0,
        level: 0,
      },
    };
    const lines = wrapParagraph(para, 134, 18);
    expect(lines).toHaveLength(1);
    expect(lines[0].segments.map((s) => s.text).join("")).toBe("ABCシステム");
  });

  it("Wraps correctly when width exceeds tolerance", () => {
    // Measurement width of "ABC system" ≈ 135.74px
    // availableWidth=120 -> Extrusion 15.74px, tolerance=120*0.02=2.4px -> Wrapped
    const props = makeRunProps({
      fontFamily: "Calibri",
      fontFamilyEa: "Meiryo",
      fontSize: 18,
    });
    const para: Paragraph = {
      runs: [{ text: "ABCシステム", properties: props }],
      properties: {
        alignment: "l",
        lineSpacing: null,
        spaceBefore: 0,
        spaceAfter: 0,
        level: 0,
      },
    };
    const lines = wrapParagraph(para, 120, 18);
    expect(lines.length).toBeGreaterThan(1);
  });

  it("When fontScale is 1, run's fontSize is used as is.", () => {
    const para = makeParagraph(["Hello World"], { fontSize: 36 });
    const linesNoScale = wrapParagraph(para, 200, 18, 1);
    // fontSize=36 is wider than fontSize=18 so it should wrap to more lines
    const paraSmall = makeParagraph(["Hello World"], { fontSize: 18 });
    const linesSmall = wrapParagraph(paraSmall, 200, 18, 1);
    expect(linesNoScale.length).toBeGreaterThanOrEqual(linesSmall.length);
  });
});
