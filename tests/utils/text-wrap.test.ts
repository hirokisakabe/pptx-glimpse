import { describe, it, expect } from "vitest";
import { wrapParagraph } from "../../src/utils/text-wrap.js";
import type { Paragraph, RunProperties } from "../../src/model/text.js";

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
  it("短いテキストは折り返さない", () => {
    const para = makeParagraph(["Hi"]);
    const lines = wrapParagraph(para, 500, 18);
    expect(lines).toHaveLength(1);
    expect(lines[0].segments).toHaveLength(1);
    expect(lines[0].segments[0].text).toBe("Hi");
  });

  it("長い英語テキストをワード境界で折り返す", () => {
    const para = makeParagraph(["The quick brown fox jumps over the lazy dog"]);
    // 幅を狭くして折り返しを発生させる
    const lines = wrapParagraph(para, 100, 18);
    expect(lines.length).toBeGreaterThan(1);
    // 各行の先頭が大文字 or 小文字の単語で始まる
    for (const line of lines) {
      expect(line.segments.length).toBeGreaterThan(0);
    }
  });

  it("長い CJK テキストを文字境界で折り返す", () => {
    const para = makeParagraph(["本日は晴天なり今日もいい天気です"]);
    const lines = wrapParagraph(para, 100, 18);
    expect(lines.length).toBeGreaterThan(1);
  });

  it("空の段落は空の行を返す", () => {
    const para = makeParagraph([]);
    const lines = wrapParagraph(para, 500, 18);
    expect(lines).toHaveLength(1);
    expect(lines[0].segments).toHaveLength(0);
  });

  it("空テキストのランのみの段落は空の行を返す", () => {
    const para = makeParagraph([""]);
    const lines = wrapParagraph(para, 500, 18);
    expect(lines).toHaveLength(1);
    expect(lines[0].segments).toHaveLength(0);
  });

  it("複数の TextRun が正しくセグメントに分割される", () => {
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
    // bold と normal で少なくとも2つのセグメントになるか、マージされない
    const allText = lines[0].segments.map((s) => s.text).join("");
    expect(allText).toBe("Bold Normal");
  });

  it("スタイルが異なる TextRun の境界をまたぐ折り返し", () => {
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
    // 各行をスペースで結合して全テキストが保持されていることを確認
    // (行末空白はトリムされるので行間にスペースを補って結合)
    const allText = lines.map((l) => l.segments.map((s) => s.text).join("")).join(" ");
    expect(allText).toContain("First part");
    expect(allText).toContain("bold and");
    expect(allText).toContain("second part");
    expect(allText).toContain("normal text");
  });

  it("availableWidth が非常に小さい場合でも少なくとも1文字は配置する", () => {
    const para = makeParagraph(["AB"]);
    const lines = wrapParagraph(para, 1, 18);
    expect(lines.length).toBeGreaterThanOrEqual(1);
    const allText = lines.flatMap((l) => l.segments.map((s) => s.text)).join("");
    expect(allText).toBe("AB");
  });

  it("行末の空白がトリムされる", () => {
    const para = makeParagraph(["Hello World"]);
    // Hello(~60px) + space(~7px) = ~67px で折り返し
    const lines = wrapParagraph(para, 80, 18);
    expect(lines.length).toBeGreaterThanOrEqual(1);
    for (const line of lines) {
      const lastSeg = line.segments[line.segments.length - 1];
      if (lastSeg) {
        expect(lastSeg.text).not.toMatch(/\s+$/);
      }
    }
  });
});
