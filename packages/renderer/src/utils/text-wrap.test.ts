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
  it("covers text-wrap behavior 1", () => {
    const para = makeParagraph(["Hi"]);
    const lines = wrapParagraph(para, 500, 18);
    expect(lines).toHaveLength(1);
    expect(lines[0].segments).toHaveLength(1);
    expect(lines[0].segments[0].text).toBe("Hi");
  });

  it("covers text-wrap behavior 2", () => {
    const para = makeParagraph(["The quick brown fox jumps over the lazy dog"]);
    // Test note.
    const lines = wrapParagraph(para, 100, 18);
    expect(lines.length).toBeGreaterThan(1);
    // Test note.
    for (const line of lines) {
      expect(line.segments.length).toBeGreaterThan(0);
    }
  });

  it("covers text-wrap behavior 3", () => {
    const para = makeParagraph(["本日は晴天なり今日もいい天気です"]);
    const lines = wrapParagraph(para, 100, 18);
    expect(lines.length).toBeGreaterThan(1);
  });

  it("covers text-wrap behavior 4", () => {
    const para = makeParagraph([]);
    const lines = wrapParagraph(para, 500, 18);
    expect(lines).toHaveLength(1);
    expect(lines[0].segments).toHaveLength(0);
  });

  it("covers text-wrap behavior 5", () => {
    const para = makeParagraph([""]);
    const lines = wrapParagraph(para, 500, 18);
    expect(lines).toHaveLength(1);
    expect(lines[0].segments).toHaveLength(0);
  });

  it("covers text-wrap behavior 6", () => {
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
    // Test note.
    const allText = lines[0].segments.map((s) => s.text).join("");
    expect(allText).toBe("Bold Normal");
  });

  it("covers text-wrap behavior 7", () => {
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
    // Test note.
    // Test note.
    const allText = lines.map((l) => l.segments.map((s) => s.text).join("")).join(" ");
    expect(allText).toContain("First part");
    expect(allText).toContain("bold and");
    expect(allText).toContain("second part");
    expect(allText).toContain("normal text");
  });

  it("covers text-wrap behavior 8", () => {
    const para = makeParagraph(["AB"]);
    const lines = wrapParagraph(para, 1, 18);
    expect(lines.length).toBeGreaterThanOrEqual(1);
    const allText = lines.flatMap((l) => l.segments.map((s) => s.text)).join("");
    expect(allText).toBe("AB");
  });

  it("covers text-wrap behavior 9", () => {
    const para = makeParagraph(["Hello World"]);
    // Test note.
    const lines = wrapParagraph(para, 80, 18);
    expect(lines.length).toBeGreaterThanOrEqual(1);
    for (const line of lines) {
      const lastSeg = line.segments[line.segments.length - 1];
      if (lastSeg) {
        expect(lastSeg.text).not.toMatch(/\s+$/);
      }
    }
  });

  it("covers text-wrap behavior 10", () => {
    // Test note.
    // Test note.
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

  it("covers text-wrap behavior 11", () => {
    // Test note.
    // Test note.
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

  it("covers text-wrap behavior 12", () => {
    // Test note.
    // Test note.
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

  it("covers text-wrap behavior 13", () => {
    const para = makeParagraph(["Hello World"], { fontSize: 36 });
    const linesNoScale = wrapParagraph(para, 200, 18, 1);
    // Test note.
    const paraSmall = makeParagraph(["Hello World"], { fontSize: 18 });
    const linesSmall = wrapParagraph(paraSmall, 200, 18, 1);
    expect(linesNoScale.length).toBeGreaterThanOrEqual(linesSmall.length);
  });
});
